using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace SpeechExcel.Execute
{
    public class Parser
    {
        private string pattern = "[第]{0,1}([零一二三四五六七八九十百千]+)列([到至和]{0,1})";

        private List<int> _parse(string strLikeField) 
		{
            List<int> fields = new List<int>();
            if (strLikeField.Contains("所有"))
            {
                fields.Add(0);
                return fields;
            }

            MatchCollection mcs = Regex.Matches(strLikeField, pattern);
            
            if (mcs != null)
            {
                foreach (Match mc in mcs)
                {
                    fields.Add(_toInt(mc.Groups[1].Value));
                    if (mc.Groups[2].Value == "到" || mc.Groups[2].Value == "至")
                    {
                        // this means that this filed is a block
                        fields.Add(0);
                    }
                }
                // Console.WriteLine(strLikeField);
                return fields;
            }
            else
            {
                return null;
            }
		}

        /// <summary>
        /// 将中文转成数字
        /// </summary>
        /// <returns>最后的解析结果</returns>
        /// <param name="word">表示需要被解析的中文文字</param>
        public static int _toInt(string word)
        {
            string num = "零一二三四五六七八九";
            string factor = "十百千";

            int a = 0;

            // 如果最后一个汉字为零，a=10
            if (word.IndexOf(factor[0]) == 0) a = 10;


            word = Regex.Replace(word, num[0].ToString(), "");

            if (Regex.IsMatch(word, "([" + num + "])$"))
            {
                a += num.IndexOf(Regex.Match(word, "([" + num + "])$").Value[0]);
            }

            if (Regex.IsMatch(word, "([" + num + "])" + factor[0]))
            {
                a += num.IndexOf(Regex.Match(word, "([" + num + "])" + factor[0]).Value[0]) * 10;
            }

            if (Regex.IsMatch(word, "([" + num + "])" + factor[1]))
            {
                a += num.IndexOf(Regex.Match(word, "([" + num + "])" + factor[1]).Value[0]) * 100;
            }

            if (Regex.IsMatch(word, "([" + num + "])" + factor[2]))
            {
                a += num.IndexOf(Regex.Match(word, "([" + num + "])" + factor[2]).Value[0]) * 1000;
            }

            return a;
        }

        /// <summary
        /// Parsers the json to dict.
        /// </summary>
        /// <returns>The json to dict.</returns>
        /// <param name="jsonText">Json text.</param>
        private Dictionary<string, object> _parseJsonToDict(string jsonText)
        {
            JsonSerializer jss = new JsonSerializer();
            JsonReader reader = new JsonTextReader(new StringReader(jsonText));
            try
            {
                return jss.Deserialize<Dictionary<string, object>>(reader);
            }
            catch (Exception e)
            {
                return null;
            }
        }

        /// <summary>
        /// A class object holds Parse result.
        /// </summary>
        public class ParseResult
        {
            public JArray entities;
            public string intent;
            public string queryText;
            
            public ParseResult()
            {
                this.intent = "";
                this.queryText = "";
            }
        }

        /// <summary>
        /// Gets the parse result.
        /// </summary>
        /// <returns>The parse result.</returns>
        /// <param name="luisText">Luis text.</param>
        public ParseResult getParseResult(string luisText)
        {
            ParseResult parseResult = new ParseResult();
            Dictionary<string, object> jsonResult = this._parseJsonToDict(luisText);  // 解析成的json结果
            JArray entities = (JArray)jsonResult["entities"];    // 单个Intent对应的entities
            JArray intents = (JArray)jsonResult["intents"];
            parseResult.queryText = jsonResult["query"] as string;  // 获取的querytext
            parseResult.intent = intents[0]["intent"].ToString();   // 获取的最大概率的intent
            parseResult.entities = entities;    // 包含的entities
            return parseResult;
        }

        public class ReplaceNode
        {
            public int startIndex, endIndex;
            public int Row, Column;
            public string content;
            public ReplaceNode(int startIndex, int endIndex, int Row, int Column,string content)
            {
                this.startIndex = startIndex; this.endIndex = endIndex;
                this.Column = Column; this.Row = Row;
                this.content = content;
            }
        }

        public class ReplaceNodeAscent : IComparer<ReplaceNode>
        {
            public int Compare(ReplaceNode x, ReplaceNode y)
            {
                return x.startIndex.CompareTo(y.startIndex);
            }
        }

        /// <summary>
        /// 将query的部分结果替换
        /// </summary>
        /// <param name="dataRange"></param>
        /// <param name="speech_text"></param>
        /// <param name="range_list"></param>
        /// <returns></returns>
        public static string replace(String speech_text, out List<ReplaceNode> replace_list)
        {
            //把文本中出现的表格中存在的内容替换成[cell_content]，并将其对应的Range对象存在range_list
            //此函数非常慢，大文件勿用
            //Stopwatch sw = new Stopwatch();
            //sw.Start();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook as Excel.Workbook;
            Excel.Worksheet sheet = wb.ActiveSheet as Excel.Worksheet;
            replace_list = new List<ReplaceNode>();
            //MessageBox.Show(wb.FullName);
            string extension = System.IO.Path.GetExtension(wb.FullName);
            string filePath = @"Temp" + extension;
            wb.SaveCopyAs(filePath);
            try
            {
                FileStream fs = File.OpenRead(filePath);
                
                IWorkbook wk = null;
                if(extension.Equals(".xls"))
                {
                    wk = new HSSFWorkbook(fs);
                }
                else
                {
                    wk = new XSSFWorkbook(fs);
                }
                ISheet isheet = wk.GetSheet(sheet.Name);
                for (int i = 0; i < isheet.LastRowNum; i++)
                {
                    IRow row = isheet.GetRow(i);
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        string s = row.GetCell(j).ToString();
                        if (s.Length > 0 && speech_text.Contains(s))
                        {
                            int start = speech_text.IndexOf(s);
                            int end = start + s.Length;
                            replace_list.Add(new ReplaceNode(start, end, i + 1, j + 1, s));
                            speech_text = speech_text.Replace(s, "[cell_content]");
                        }
                    }
                }
                ReplaceNodeAscent sort_ascent = new ReplaceNodeAscent();
                replace_list.Sort(sort_ascent);
                return speech_text;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return "";
            }
        }

    }
}
