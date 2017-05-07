using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Cognitive.LUIS;
using System.Windows;
using System.Linq;

namespace SpeechExcel.Execute
{
    static class SheetOpe
    {
        /// <summary>
        /// 取消筛选（显示原图）
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="queryText"></param>
        public static void cancelFilter(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            cancelFilter();
        }

        /// <summary>
        /// 筛选功能
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="queryText"></param>
        public static void filter(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            string valueWord = "";
            string operate = "";
            MessageBox.Show("Filter");
            bool valueFilter = false;
            foreach (var item in res.GetAllEntities())
            {
                if (item.Name == "FilterOperator::greater_equal")
                {
                    operate = ">=";
                    valueFilter = true;
                }
                else if (item.Name == "FilterOperator::greater_than")
                {
                    operate = ">";
                    valueFilter = true;
                }
                else if (item.Name == "FilterOperator::less_than")
                {
                    operate = "<";
                    valueFilter = true;
                }
                else if (item.Name == "FilterOperator::less_equal")
                {
                    operate = "<=";
                    valueFilter = true;
                }
                else if (item.Name == "builtin.number")
                {
                    valueWord = item.Resolution["value"].ToString();
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
                    MessageBox.Show("找不到要对应的列");
                }
                else if (valueWord == "")
                {
                    MessageBox.Show("无法确定筛选范围");
                }
                else
                {
                    OperateFilter(column, operate + valueWord);
                }
            }
            else
            {
                int column = 0;
                HashSet<string> typeName = new HashSet<string>();
                //int column = parseColumnName(res.OriginalQuery, typeName);
                foreach (var cell in replace_list)
                {
                    if (cell.Row != 1)
                    {
                        string s = cell.content;
                        typeName.Add(s);
                    }
                    if (cell.Row == 1)
                    {
                        column = cell.Column;
                    }
                }
                if (column == 0)
                {
                    MessageBox.Show("找不到要对应的列");
                }
                else if (typeName.Count == 0)

                {
                    MessageBox.Show("找不到要筛选的类别");
                }
                else
                {
                    TypeFilter(column, typeName);
                }
            }
        }

        /// <summary>
        /// 值筛选
        /// </summary>
        /// <param name="column"></param>
        /// <param name="criteral"></param>
        public static void OperateFilter(int column, string criteral)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = ws.UsedRange;
            rng.AutoFilter(column, criteral, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

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
        static public void TypeFilter(int column, HashSet<string> typeName)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = ws.UsedRange;
            rng.AutoFilter(column, typeName.Count > 0 ? typeName.ToArray() : Type.Missing, Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
        }


        /// <summary>
        /// 取消筛选
        /// </summary>
        public static void cancelFilter()
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            ws.ShowAllData();
        }


        public static string get_value(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range dataRange = worksheet.UsedRange;

            //MessageBox.Show("It's OK");
            if (replace_list.Count == 3)
            {
                replace_list.RemoveAt(0);
            }
            string ret = "";
            if (replace_list.Count == 2)
            {
                HashSet<string> types = new HashSet<string>();
                types.Add(replace_list[0].content);
                //cancelFilter();
                //TypeFilter(replace_list[0].Column, types);
                int row_id = replace_list[0].Row;
                int column_id = replace_list[1].Column;
                //MessageBox.Show(replace_list[0].content);
                //MessageBox.Show(column_id.ToString());
                //MessageBox.Show("Row:" + row_id.ToString() + ", Col:" + column_id.ToString());
                ret = _get_value(dataRange, row_id, column_id);
            }
            return ret;
        }


        /// <summary>
        /// 获取某个
        /// </summary>
        /// <param name="DataRange">数据来源</param>
        /// <param name="row_id">行号</param>
        /// <param name="column_id">列号</param>
        /// <returns></returns>
        public static String _get_value(Excel.Range dataRange, int row_id, int column_id)
        {
            ((Excel.Range)dataRange.Cells[row_id, column_id]).Activate();
            return ((Excel.Range)dataRange.Cells[row_id, column_id]).Text.ToString();
        }


        public static String get_sum(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            String ret = "";
            if (replace_list.Count == 3)
            {
                replace_list.RemoveAt(0);
            }
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range dataRange = worksheet.UsedRange;
            Excel.Range rng = dataRange;
            if (replace_list.Count == 2)
            {
                HashSet<String> types = new HashSet<string>();
                types.Add(replace_list[0].content);
                int column_id = replace_list[0].Column;
                rng = Filter.TypeFilter(column_id, types);
                replace_list.RemoveAt(0);
            }
            if(replace_list.Count == 1)
            {
                int column_id = replace_list[0].Column;
                rng = (Excel.Range)dataRange.Columns[column_id];
                ret = Globals.ThisAddIn.Application.WorksheetFunction.Subtotal(109, rng).ToString();
            }
            return ret;
        }

        public static string sort(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range dataRange = worksheet.UsedRange;
            Excel.XlSortOrder sort_order = Excel.XlSortOrder.xlAscending;
            foreach (var item in res.GetAllEntities())
            {
                if (item.Name == "SortOrder::Ascending")
                {
                    sort_order = Excel.XlSortOrder.xlAscending;
                }
                else if (item.Name == "SortOrder::Descending")
                {
                    sort_order = Excel.XlSortOrder.xlDescending;
                }
            }
            string ret = "";
            if (replace_list.Count == 0)
            {
                ret = "Can't Parser";
            }
            else
            {
                int column_id = replace_list[0].Column;
                ret = _sort_by_column_id(dataRange, column_id, sort_order);
            }
            return ret;
        }

        /// <summary>
        /// 对指定列排序，DataRange:Range对象指定范围，column_id：指定列号, sort_order：排序方式
        /// </summary>
        /// <param name="DataRange"></param>
        /// <param name="column_id"></param>
        /// <param name="sort_order"></param>
        public static string _sort_by_column_id(Excel.Range DataRange, int column_id, Excel.XlSortOrder sort_order)
        {
            //对指定列排序，DataRange:Range对象指定范围，column_id：指定列号, sort_order：排序方式
            Boolean oldFresh = Globals.ThisAddIn.Application.ScreenUpdating;
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                DataRange.Sort(DataRange.Columns[column_id], sort_order,
                    Type.Missing, Type.Missing, sort_order, Type.Missing, sort_order,
                    Excel.XlYesNoGuess.xlGuess);
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }
            return "";
        }


        public static string find_min_max(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range dataRange = worksheet.UsedRange;
            int min_max = -1;
            foreach (var item in res.GetAllEntities())
            {
                if (item.Name == "maximum")
                {
                    min_max = 1;
                }
                else if (item.Name == "minimum")
                {
                    min_max = 0;
                }
            }
            string ret = "";
            if (min_max == -1)
            {
                ret = Properties.Resources.unkown;
            }
            if (replace_list.Count == 0)
            {
                ret = "抱歉，我不知道是第几列";
            }
            else
            {
                int column_id = replace_list[0].Column;
                ret = _find_min_max(dataRange, column_id, min_max);
                HashSet<string> filter_content = new HashSet<string>();
                filter_content.Add(ret);
                //cancelFilter();
                //TypeFilter(column_id, filter_content);
            }
            return ret;
        }

        /// <summary>
        /// 最值查找
        /// </summary>
        /// <param name="DataRange"></param>
        /// <param name="column_id"></param>
        /// <param name="min_max"></param>
        /// <returns></returns>
        public static string _find_min_max(Excel.Range dataRange, int column_id, int min_max)
        {
            //找到对应列的最大最小值,DataRange:指定Range范围，column_id:指定列号，min_max:0为找最小值，1为找最大值
            Boolean oldFresh = Globals.ThisAddIn.Application.ScreenUpdating;
            try
            {
                //Globals.ThisAddIn.Application.ScreenUpdating = false;
                dynamic t;
                if (min_max == 0)
                {
                    // min
                    t = Globals.ThisAddIn.Application.WorksheetFunction.Min(dataRange.Columns[column_id]);
                }
                else
                {
                    // max
                    t = Globals.ThisAddIn.Application.WorksheetFunction.Max(dataRange.Columns[column_id]);
                }
                //MessageBox.Show(((Excel.Range)dataRange[2, column_id]).GetType().ToString());

                return t.ToString();

            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }
        }

    }
}
