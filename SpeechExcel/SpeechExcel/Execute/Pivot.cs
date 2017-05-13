using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using Microsoft.Cognitive.LUIS;

namespace SpeechExcel.Execute
{
    class Pivot
    {
        public static string mess = "";

        public struct PivotData
        {
            public int type;   // type == 0: row, type == 1: column, type == 2: data
            public object name;    // string or int
            public int add_in;
            /// <summary>
            /// Pivot Construct
            /// </summary>
            /// <param name="type">field's type, 0: row, 1: column, 2: data</param>
            /// <param name="name">field's name</param>
            public PivotData(int type, object name, int add_in = 4) { this.type = type; this.name = name; this.add_in = add_in; }
        }

        /// <summary>
        /// 图表名称和chartype的映射
        /// </summary>
        public static Dictionary<string, Excel.XlChartType> chartName = new Dictionary<string, Excel.XlChartType>()
        {
            { "chartype::1", Excel.XlChartType.xl3DColumn },
        };

        public static Dictionary<int, string> colNameDict = null;

        /// <summary>
        /// 创建透视图表
        /// </summary>
        /// <param name="entities"></param>
        public static string CreatePivot(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            mess = "";
            Excel.XlChartType chartType = Excel.XlChartType.xlColumnStacked;
            List<int> colIdxes = new List<int>();
            bool state = false;
            // 使用[col]和[col]的数据绘制成（[charType]）
            foreach (Entity item in res.GetAllEntities())
            {
                if (item.Name == "builtin.ordinal")
                {
                    colIdxes.Add(Parser._toInt(item.Value.Substring(1)));
                }
                else if (item.Name == "pivotele")
                {
                    state = true;
                }
            }
            if (!state)
            {
                return "Sorry, I do not know what you wanna do.";
            }

            Boolean oldFresh = Globals.ThisAddIn.Application.ScreenUpdating;
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                createPivot(chartType); // create table and pivot chart
                addColumns(colIdxes);   // add columns to table and chart, if count of colIdxes is 0, do nothing
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }
            return mess;
        }

        /// <summary>
        /// Add more column to pivot table, 
        /// </summary>
        /// <param name="res"></param>
        /// <param name="dataList"></param>
        public static string AddColumn(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            mess = "";
            // firstly, check table type whether pivot
            List<int> colIdxes = new List<int>();
            bool state = false;
            foreach (Entity item in res.GetAllEntities())
            {
                if (item.Name == "builtin.ordinal") colIdxes.Add(Parser._toInt(item.Value.Substring(1)));
                else if (item.Name == "action::add") state = true;
            }
            if (!state || colIdxes.Count < 1)
            {
                return "抱歉，我不能添加列。";
            }

            Boolean oldFresh = Globals.ThisAddIn.Application.ScreenUpdating;
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                addColumns(colIdxes);
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }
            return mess;
            
        }

        /// <summary>
        /// 修改某个列值的统计方式，前提是这个列值得是DataFeild
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="queryText"></param>
        public static string ChangeFunc(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            // 将[col]的统计方式[modify]{成}[functionName]
            return "对透视图修改了统计函数。";
        }

        public static string OriInterFace(List<int> idx, Excel.XlChartType charType)
        {
            mess = "";
            Boolean oldFresh = Globals.ThisAddIn.Application.ScreenUpdating;
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                createPivot(charType);
                if (charType == Excel.XlChartType.xlArea) addColumns(idx, true);
                else addColumns(idx);
            }
            catch
            {
                mess = "抱歉，由于某些原因我不能创建透视图。";
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }
            return mess;
        }

        /// <summary>
        /// Create pivot table and pivot chart
        /// </summary>
        /// <param name="charType">char type</param>
        public static void createPivot(Excel.XlChartType charType)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet curSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Worksheet newSheet = workbook.Sheets.Add() as Excel.Worksheet;
            Excel.PivotTable table = workbook.PivotCaches().Create(SourceType: Excel.XlPivotTableSourceType.xlDatabase,
                SourceData: curSheet.UsedRange, Version: 6).CreatePivotTable(TableDestination: Globals.ThisAddIn.Application.ActiveCell,
                DefaultVersion: 6);
            newSheet.Select();
            newSheet.Shapes.AddChart2(201, charType).Select();
            workbook.ActiveChart.SetSourceData(Source: newSheet.UsedRange);
            newSheet.Shapes.Item(1).ScaleWidth(1.4f, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);
            newSheet.Shapes.Item(1).ScaleHeight(1.6f, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);
            // draw pivot chart
            Globals.ThisAddIn.Application.ActiveWorkbook.ShowPivotChartActiveFields = true;
        }

        /// <summary>
        /// Add columns to chart and table
        /// </summary>
        /// <param name="colIdxes">列的序列号集合</param>
        public static void addColumns(List<int> colIdxes, bool group=false)
        {
            if (colIdxes.Count == 0) return;
            Excel.PivotTable table1 = Globals.ThisAddIn.Application.ActiveChart.PivotLayout.PivotTable; // get current pivot table (active)
                                                                                                        // Create and populate the PivotTable.

            if (group)
            {
                // the seconde column should be parsed as row
                table1.AddDataField(table1.PivotFields(colIdxes[0]),
                    Type.Missing, Excel.XlConsolidationFunction.xlCount);
                Excel.PivotField customerField = null;
                if (colIdxes.Count >= 2) customerField = (Excel.PivotField)table1.PivotFields(colIdxes[1]);
                else customerField = (Excel.PivotField)table1.PivotFields(colIdxes[0]);
                customerField.Orientation =
                        Excel.XlPivotFieldOrientation.xlRowField;
                customerField.Position = 1;
                Excel.Range rng = customerField.DataRange.Cells[1] as Excel.Range;
                rng.Group(true, true, 15);
            }
            else
            {
                Excel.PivotField customerField =
                    (Excel.PivotField)table1.PivotFields(colIdxes[0]);
                customerField.Orientation =
                    Excel.XlPivotFieldOrientation.xlRowField;
                customerField.Position = 1;

                int count = colIdxes.Count;
                for (int i = 1; i < count; i++)
                {
                    table1.AddDataField(table1.PivotFields(colIdxes[i]),
                        Type.Missing, Excel.XlConsolidationFunction.xlSum);
                }
                if (count > 3)
                {
                    table1.DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                    table1.DataPivotField.Position = 1;
                }
            }
        }
    }
}
