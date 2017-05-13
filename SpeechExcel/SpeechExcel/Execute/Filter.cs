using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Cognitive.LUIS;
using System.Windows;

namespace SpeechExcel.Execute
{
    class Filter
    {
        /// <summary>
        /// 取消筛选（显示原图）
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="queryText"></param>
        public static string cancelFilter(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            cancelFilter();
            return "";
        }

        /// <summary>
        /// 筛选功能
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="queryText"></param>
        public static string filter(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            //string valueWord = "";
            //string operate = "";
            string errorMessage= "";
            List<string> valueWord = new List<string>();
            List<string> operate = new List<string>();
            bool valueFilter = false;
            bool logicOr = false;
            foreach (var item in res.GetAllEntities())
            {
                if (item.Name == "FilterOperator::greater_equal")
                {
                    //operate = ">=";
                    operate.Add(">=");
                    valueFilter = true;
                }
                else if (item.Name == "FilterOperator::greater_than")
                {
                    //operate = ">";
                    operate.Add(">");
                    valueFilter = true;
                }
                else if (item.Name == "FilterOperator::less_than")
                {
                    //operate = "<";
                    operate.Add("<");
                    valueFilter = true;
                }
                else if (item.Name == "FilterOperator::less_equal")
                {
                    //operate = "<=";
                    operate.Add("<=");
                    valueFilter = true;
                }
                else if (item.Name == "builtin.number")
                {
                    //valueWord = item.Resolution["value"].ToString();
                    valueWord.Add(item.Resolution["value"].ToString());
                }
                else if (item.Name == "logic::or")
                {
                    logicOr = true;
                }
            }
            if (valueFilter)
            {
                int column = 0;
                foreach (var cell in replace_list)
                {
                    if (cell.Row == 1)
                    {
                        column = cell.Column;
                    }
                }
                if (column == 0)
                {
                    // MessageBox.Show("找不到要对应的列");
                      errorMessage = "找不到对应的列";    
                }
                // else if (valueWord == "")
                else if (valueWord.Count == 0)
                {
                    // MessageBox.Show("无法确定筛选范围");
                    errorMessage = "无法确定筛选范围";
                }
                else
                {
                    // 预处理 
                    //if (operate.Count == 1)
                    //    operate.Add("");
                    //if (valueWord.Count == 1)
                    //    valueWord.Add("");
                    string[] operateArray = new string[2];
                    for (int i = 0; i < operate.Count; i++)
                    {
                        if (valueWord.Count == 2 && operate.Count == 2)
                        {
                            // 或者
                            if (logicOr)
                            {
                                if ((operate[i][0] == '>'))
                                {
                                    string value = Convert.ToDouble(valueWord[0]) > Convert.ToDouble(valueWord[1]) ? valueWord[0] : valueWord[1];
                                    operateArray[i] = operate[i] + value;
                                }
                                else if ((operate[i][0] == '<'))
                                {
                                    string value = Convert.ToDouble(valueWord[0]) < Convert.ToDouble(valueWord[1]) ? valueWord[0] : valueWord[1];
                                    operateArray[i] = operate[i] + value;
                                }
                            }
                            // 并且
                            else
                            {
                                if ((operate[i][0] == '>'))
                                {
                                    string value = Convert.ToDouble(valueWord[0]) < Convert.ToDouble(valueWord[1]) ? valueWord[0] : valueWord[1];
                                    operateArray[i] = operate[i] + value;
                                }
                                else if ((operate[i][0] == '<'))
                                {
                                    string value = Convert.ToDouble(valueWord[0]) > Convert.ToDouble(valueWord[1]) ? valueWord[0] : valueWord[1];
                                    operateArray[i] = operate[i] + value;
                                }
                            }
                        }
                        else
                        {
                            operateArray[i] = operate[i] + valueWord[i];
                        }
                    }
                    //OperateFilter(column, operate + valueWord);
                    if (operate.Count == valueWord.Count)
                        OperateFilter(column, logicOr, operateArray[0], operateArray[1]);
                    else
                        errorMessage = "我不太懂你的意思";
                }
            }
            else
            {
                int column = 0;
                HashSet<string> typeName = new HashSet<string>();
                HashSet<int> columnList = new HashSet<int>();
                //int column = parseColumnName(res.OriginalQuery, typeName);
                foreach (var cell in replace_list)
                {
                    if (cell.Row != 1)
                    {
                        string s = cell.content;
                        typeName.Add(s);
                        column = cell.Column;
                        columnList.Add(cell.Column);
                    }
                    //if (cell.Row == 1)
                    //{
                    //    column = cell.Column;
                    //}
                }
                if (column == 0)
                {
                    //MessageBox.Show("找不到要对应的列");
                    errorMessage = "找不到对应的列";
                }
                else if (typeName.Count == 0)
                {
                    //MessageBox.Show("找不到要筛选的类别");
                    errorMessage = "找不到要筛选的类别";
                }else if (columnList.Count > 1)
                {
                    errorMessage = "筛选的对象在多个列中出现。麻烦筛选的时候说一下是哪一列。";
                }
                else
                {
                    TypeFilter(column, typeName);
                }
            }
            return errorMessage;
        }

        /// <summary>
        /// 值筛选
        /// </summary>
        /// <param name="column"></param>
        /// <param name="criteral"></param>
        public static Excel.Range OperateFilter(int column, bool logicOr, string criteral1, string criteral2 = "")
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.XlAutoFilterOperator opt = logicOr ? Excel.XlAutoFilterOperator.xlOr : Excel.XlAutoFilterOperator.xlAnd;
            Excel.Range rng = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
            rng = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
            if (criteral2 == "" || criteral2 == null)
                rng.AutoFilter(column, criteral1, opt, Type.Missing, true);
            else
                rng.AutoFilter(column, criteral1, opt, criteral2, true);
            return ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);

        }
        /// <summary>
        /// 找出speech_text中的列名的id和要筛选的行名字
        /// </summary>
        /// <param name="speech_text">请求的原文</param>
        /// <param name="rowSet">接收返回的分类名</param>
        /// <returns>请求中包含的列id</returns>
        public static int parseColumnName(string speech_text, HashSet<string> rowSet)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = ws.UsedRange;
            int column = 0;
            for (int i = 1; i <= rng.Columns.Count; i++)
            {
                string s = ((Excel.Range)rng.Cells[1, i]).Text.ToString();
                if (s.Length > 0 && speech_text.Contains(s))
                {
                    column = i;
                    break;
                }
            }
            for (int i = 1; i <= rng.Rows.Count; i++)
            {
                string s = ((Excel.Range)rng.Cells[i, column]).Text.ToString();
                if (s.Length > 0 && speech_text.Contains(s))
                {
                    rowSet.Add(s);
                }
            }
            return column;

        }
        /// <summary>
        /// 根据源文本获取列id
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static int columnInText(string text)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = ws.UsedRange;
            int column = 0;
            for (int i = 1; i <= rng.Columns.Count; i++)
            {
                string s = ((Excel.Range)rng.Cells[1, i]).Text.ToString();
                if (s.Length > 0 && text.Contains(s))
                {
                    column = i;
                    break;
                }
            }
            return column;
        }
        /// <summary>
        /// 分类筛选
        /// </summary>
        /// <param name="column"> 要筛选的列id</param>
        /// <param name="typeName"> 筛选出来的类别名字的list</param>
        /// <returns></returns>
        static public Excel.Range TypeFilter(int column, HashSet<string> typeName)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = ws.UsedRange;
            rng.AutoFilter(column, typeName.Count > 0 ? typeName.ToArray() : Type.Missing, Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
            return ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
        }

        /// <summary>
        /// 取消筛选
        /// </summary>
        public static void cancelFilter()
        {
            try
            {
                Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                ws.ShowAllData();
            }
            catch (Exception e)
            {

            }
        }
    }
}
