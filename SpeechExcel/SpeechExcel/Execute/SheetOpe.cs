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
       
        public static void get_value(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range dataRange = worksheet.UsedRange;

            //MessageBox.Show("It's OK");
            if (replace_list.Count == 3)
            {
                replace_list.RemoveAt(0);
            }
            if (replace_list.Count == 2)
            {
                int row_id = replace_list[0].Row;
                int column_id = replace_list[1].Column;
                MessageBox.Show("Row:" + row_id.ToString() + ", Col:" + column_id.ToString());
                MessageBox.Show(_get_value(dataRange, row_id, column_id));
            }
            return;
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
            return ((Excel.Range)dataRange.Cells[row_id, column_id]).Text.ToString();
        }

        public static void sort(LuisResult res, List<Parser.ReplaceNode> replace_list)
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
            if (replace_list.Count == 0)
            {
                return;
            }
            else
            {
                int column_id = replace_list[0].Column;
                _sort_by_column_id(dataRange, column_id, sort_order);
            }
        }

        /// <summary>
        /// 对指定列排序，DataRange:Range对象指定范围，column_id：指定列号, sort_order：排序方式
        /// </summary>
        /// <param name="DataRange"></param>
        /// <param name="column_id"></param>
        /// <param name="sort_order"></param>
        public static void _sort_by_column_id(Excel.Range DataRange, int column_id, Excel.XlSortOrder sort_order)
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
        }


        public static void find_min_max(LuisResult res, List<Parser.ReplaceNode> replace_list)
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
            if (min_max == -1)
            {
                MessageBox.Show("没有找到maximum或者minimum实体");
                return;
            }
            if (replace_list.Count == 0)
            {
                MessageBox.Show("不能确定列");
                return;
            }
            else
            {
                int column_id = replace_list[0].Column;
                MessageBox.Show(_find_min_max(dataRange, column_id, min_max));
            }
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
                Globals.ThisAddIn.Application.ScreenUpdating = false;
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

                return t.ToString();

            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }
        }

    }
}
