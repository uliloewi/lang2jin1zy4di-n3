using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Net;
using System.IO;

namespace lang2jin1zy4diän3
{
    public partial class Form1 : Form
    {
        ///dict 形如
        /// Key(字)    | Value（所在頁）
        /// 詉         |  31
        /// _________________
        /// 𧦮         |  1
        ///            |  31
        /// _________________
        /// 閫         | 1091
        ///...
        Dictionary<string, List<int>> dict = new Dictionary<string, List<int>>();                       
        public Form1()
        {
            InitializeComponent();
            button3.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "word files (*.doc;*.docx)|*.doc;*.docx|All files (*.*)|*.*";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                object readOnly = true;
                object miss = System.Reflection.Missing.Value;
                object path = openFileDialog1.FileName;
                Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path);
                string totaltext = "";
                for (int i = 0; i < docs.Paragraphs.Count; i++)
                {
                    Range paraRange = docs.Paragraphs[i + 1].Range;
                    string paragraphText = paraRange.Text;
                    List<string> chars = getChracters(paragraphText);
                    int page = paraRange.Sentences.First.Information[WdInformation.wdActiveEndAdjustedPageNumber];   // 這些字所在頁數               
                    foreach (string ch in chars)
                    {//字和他所在頁碼
                        if (dict.Keys.Contains(ch))
                        {
                            dict[ch].Add(page);
                        }
                        else
                        {
                            List<int> pageNo = new List<int>();
                            pageNo.Add(page);
                            dict.Add(ch, pageNo);
                        }
                    }
                    totaltext += " \r\n " + paragraphText;
                }
                Console.WriteLine(totaltext);
                docs.Close();
                word.Quit();
                button3.Enabled = true;
            }
        }



        /// <summary>
        /// 選字頭。比如從
        /// “閫(梱) kuen3 門檻、門限。”中選出“閫”“梱”。
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private List<string> getChracters(string paragraph)
        {
            List<string> chars = new List<string>();
            char[] charList = new char[] { 'a', 'ä', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'ü', 'v', 'w', 'x', 'y', 'z', ' ', };
            int checkSurrogate = 1;
            char firstSurrogate = ' ';
            foreach (char ch in paragraph)
            {
               if (ch != '(' && ch != ')' && !charList.Contains(ch))
                {
                    if (char.IsHighSurrogate(ch) && checkSurrogate == 1)
                    {//複雜字的前半個
                        firstSurrogate = ch;
                        checkSurrogate++;
                    }
                    else if (checkSurrogate == 2)
                    {//複雜字的後半個
                        string complexChar = String.Concat(firstSurrogate, ch);
                        chars.Add(complexChar);
                        checkSurrogate = 1;
                    }
                    else
                    {//簡單字直接加
                        chars.Add(ch.ToString());
                    }
                }
                else if (charList.Contains(ch))//空格或拼音就終止
                    break;
           }
            return chars;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ///dictUnicode 形如
            /// Key（萬國碼）   | Value（字和所在頁）
            /// 35401          | 詉  31
            /// _____________________________
            ///162222          | 𧦮  1
            ///                      31
            /// _____________________________
            ///38315           | 閫  1091
            ///...
            Dictionary<int, KeyValuePair<string, List<int>>> dictUnicode = new Dictionary<int, KeyValuePair<string, List<int>>>();
            foreach (KeyValuePair<string, List<int>> kvp in dict)
            {//爲每個字配上萬國碼
                string hexCode = GetInfo(kvp.Key, "http://www.unicode.org/cgi-bin/GetUnihanData.pl?codepoint=", "<title>Unihan data for U+", "</title>");//找出十六進制萬國碼
                int unicode = Int32.Parse(hexCode, System.Globalization.NumberStyles.HexNumber);//十六進制萬國碼轉爲十進制
                dictUnicode.Add(unicode, kvp);
            }

            ///radicalDict 形如
            ///Key 部首和筆畫             | Value
            ///言 5 畫                   | 35401   詉  31 
            ///                          | 162222  𧦮  1
            ///                          |            31
           ///________________________________________________________                                       
           ///門 7 畫                   | 38315  閫  1091
           ///...
           Dictionary<string, Dictionary<int, KeyValuePair<string, List<int>>>> radicalDict = new Dictionary<string, Dictionary<int, KeyValuePair<string, List<int>>>>();
           foreach (var item in dictUnicode.OrderBy(i => i.Key))
            {
                string stroke = GetInfo(item.Value.Key, "https://zh.wiktionary.org/zh-hant/", "部首索引 ", "畫</li>");//找出部首和筆畫 “言" class="mw-redirect">言</a></b> ＋ 5”
                stroke = stroke.Substring(23).Replace("</a></b> ＋", "") + "畫";
                if (!radicalDict.Keys.Contains(stroke))
                {//偏旁和筆畫                   
                    Dictionary<int, KeyValuePair<string, List<int>>> uniCodes = new Dictionary<int, KeyValuePair<string, List<int>>>();
                    uniCodes.Add(item.Key, item.Value);
                    radicalDict.Add(stroke, uniCodes);
                }
                else
                {
                    radicalDict[stroke].Add(item.Key, item.Value);
                }
          }


            using (System.IO.StreamWriter file = new System.IO.StreamWriter(openFileDialog1.FileName.Replace(".doc", "") + ".txt"))
            {//製作部首檢字表
                foreach (var rd in radicalDict.OrderBy(i => i.Value.Keys.Min()))
                {
                   file.WriteLine(rd.Key.Substring(2));
                    foreach (var localDict in rd.Value.OrderBy(j => j.Key))
                    {
                        int cnt = 0;
                        foreach (int page in localDict.Value.Value)
                        {
                            if (cnt == 0)
                                file.WriteLine(localDict.Value.Key + "  " + page);// + "  " + localDict.Key);
                            else
                                file.WriteLine("    " + page);
                            cnt++;
                        }
                    }
                }
            }
        }

        ///// <summary>
        /////
        ///// </summary>
        ///// <param name="str"></param>
        ///// <returns></returns>
        //private int GetUniCode(string str)
        //{
        //    int pos = str.IndexOf("<title>Unihan data for U+");
        //    int endPos = str.IndexOf("</title>");
        //    int distance = endPos - pos;
        //    string hexCode = distance == 30 ? str.Substring(pos + 25, 5) : str.Substring(pos + 25, 4);
        //    int num = Int32.Parse(hexCode, System.Globalization.NumberStyles.HexNumber);
        //    return num;

        //}



        private string GetInfo(string character, string webName, string pattern, string endPattern)
        {
            string info = "";
            string url = webName + character;
            HttpWebRequest request = HttpWebRequest.Create(url) as HttpWebRequest;
            request.Method = "GET";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream receiveStream = response.GetResponseStream();
            Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
            // Pipes the stream to a higher level stream reader with the required encoding format.
            StreamReader readStream = new StreamReader(receiveStream, encode);

            Char[] read = new Char[256];
            // Reads 256 characters at a time.    
            int count = readStream.Read(read, 0, 256);
            while (count > 0)
            {
                String strBlock = new String(read, 0, count);
                if (strBlock.Contains(pattern))
                {
                    info = ExtractInfo(strBlock, pattern, endPattern);
                    break;
                }
                count = readStream.Read(read, 0, 256);
            }
            // Releases the resources of the response.
            response.Close();
            // Releases the resources of the Stream.
            readStream.Close();
            return info;
        }

        private string ExtractInfo(string text, string startPattern, string endPattern)
        {
            int pos = text.IndexOf(startPattern);
            int endPos = text.IndexOf(endPattern);
            int endOfStartPattern = pos + startPattern.Length;
            string info = "";
            if (endPos > endOfStartPattern)
                info = text.Substring(endOfStartPattern, endPos - endOfStartPattern);
            else
                info = text.Substring(endOfStartPattern);
            return info;
        }
        /// <summary>
        /// 有些常用字不在廣韻中，打開常用字表，找出這些字，保存在“……uei4lu5chang2iong4zy4.txt”
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            List<string> res = new List<string>();
            openFileDialog1.Filter = "text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {//打開常用字表，一個txt文件
                string[] usualChars = System.IO.File.ReadAllLines(openFileDialog1.FileName);

                foreach (string str in usualChars)
                {
                    foreach (char ch in str)
                   {
                       if (!dict.Keys.Contains(ch.ToString()))//如果廣韻不含此常用字
                        {
                            res.Add(ch.ToString());
                        }
                    }
                }
                string[] resText = res.ToArray();
                string ss = openFileDialog1.InitialDirectory;
                System.IO.File.WriteAllLines(openFileDialog1.FileName.Replace(".txt", "uei4lu5chang2iong4zy4.txt"), resText);
                button3.Enabled = false;
            }
        }

        /// <summary>
        /// 插入簡體字
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            List<string> res = new List<string>();
            Dictionary<string, string> tradToSim = new Dictionary<string, string>();
            openFileDialog1.Filter = "text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {//打開繁簡字表，一個txt文件
                string[] rows = System.IO.File.ReadAllLines(openFileDialog1.FileName);

                foreach (string row in rows)
                {
                    List<string> elem = row.Split('〕').ToList();
                    foreach (string str in elem)
                    {
                        if (str.Length == 3 && !tradToSim.Keys.Contains(str.Substring(str.Length - 1, 1)))
                            tradToSim.Add(str.Substring(str.Length - 1, 1), str.Substring(0, 1)); //繁體字做key，簡體字做value

                    }
                }
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(openFileDialog1.FileName.Replace(".txt", "foundTradChar") + ".txt"))
            {//
                foreach (string ch in dict.Keys)
                {
                    if (tradToSim.Keys.Contains(ch))
                    {
                        string pages = "";
                        foreach (int pn in dict[ch])
                            pages += pn.ToString() + ",";//所有頁碼
                        file.WriteLine(ch + " pages: " + pages + " °" + tradToSim[ch]); //         
                    }
                }
            }

            //openFileDialog1.Filter = "word files (*.doc;*.docx)|*.doc;*.docx|All files (*.*)|*.*";
            //Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            //if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    object missing = System.Reflection.Missing.Value;
            //    object readOnly = false;
            //    object isVisible = true;
            //    word.Visible = true;
            //    object path = openFileDialog1.FileName;
            //    Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path,ref missing,ref readOnly,ref missing,ref missing,ref missing,ref missing,ref missing,
            //        ref missing,ref missing,ref missing,isVisible);
            //    docs.Activate();
            //    for (int i = 0; i < docs.Paragraphs.Count; i++)
            //    {
            //        Range paraRange = docs.Paragraphs[i + 1].Range;
            //        string paragraphText = paraRange.Text;
            //        string charHead = getChracters(paragraphText).FirstOrDefault();
            //        if (tradToSim.Keys.Contains(charHead))
            //        {
            //            paraRange.Text=paragraphText.Insert(1, tradToSim[charHead]);
            //        }
            //    }
            //    docs.Save();
            //}
        }
    }
}

