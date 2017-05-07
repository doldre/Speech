using Microsoft.Cognitive.LUIS;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpeechExcel.Execute
{
    static class Other
    {
        private static Excel.Workbook workbook;
        private static Excel.Worksheet sheet;
        private static Excel.Worksheet curSheet;

        private static string mss;

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

        public static string UseTemplate(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            mss = "";
            sheetTemplate();
            return mss;
        }

        /// <summary>
        /// 创建财报分析表
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="queryText"></param>
        public static void sheetTemplate()
        {
            Boolean oldFresh = Globals.ThisAddIn.Application.ScreenUpdating;
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                sheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                curSheet = workbook.Worksheets.Add() as Excel.Worksheet;
                curSheet.Select();
                Excel.Range headingBar = curSheet.Range["A22:D22"];

                // part1: 年度开支对比
                // setting heading
                string title = "年度开支对比";
                createHead(headingBar, title);
                // setting table
                createPivotTable(1, curSheet, curSheet.Range["A23"], title,
                    curSheet.Range["A5"], curSheet.Range["G21"], sheet.UsedRange, Excel.XlChartType.xlColumnClustered, 11);
                // setting fields
                List<PivotData> fields = new List<PivotData>() { new PivotData(1, "日期", 6), new PivotData(0, "类别"), new PivotData(2, "实际成本") };
                addField(workbook.ActiveChart.PivotLayout.PivotTable, fields);
                // part2: 年度开支分类占比
                headingBar = curSheet.Range["H22:M22"];
                title = "2016-2017年开支分类占比";
                createHead(headingBar, title);
                createPivotTable(2, curSheet, curSheet.Range["H23"], title,
                    curSheet.Range["H5"], curSheet.Range["O21"], sheet.UsedRange, Excel.XlChartType.xlPieOfPie, 6);
                List<PivotData> fields2 = new List<PivotData>() { new PivotData(1, "日期"), new PivotData(0, "类别"), new PivotData(2, "实际成本") };
                addField(workbook.ActiveChart.PivotLayout.PivotTable, fields2);
                // part3: 每月开支趋势图
                headingBar = curSheet.Range["P22:Q22"];
                title = "每月开支趋势";
                createHead(headingBar, "每月开支趋势");
                createPivotTable(3, curSheet, curSheet.Range["P23"], title,
                    curSheet.Range["P5"], curSheet.Range["W21"], sheet.UsedRange, Excel.XlChartType.xlColumnClustered, 11);
                List<PivotData> fields3 = new List<PivotData>() { new PivotData(0, "日期"), new PivotData(2, "实际成本") };
                addField(workbook.ActiveChart.PivotLayout.PivotTable, fields3);
                // setting sheet's heading
                headingBar = curSheet.Range["A1:V3"];
                createHead(headingBar, "2016-2017年度开支报表概览", 13.5, fontSize: 18);
            }
            catch
            {
                mss = "在使用报表分析功能时出错，请确定原表中具有：类别，实际成本和时间";
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }
        }

        /// <summary>
        /// 为每个透视表设定heading
        /// </summary>
        /// <param name="headingBar">range的范围</param>
        /// <param name="headingText">heading的内容</param>
        public static void createHead(Excel.Range headingBar, string headingText, double height = 27.0, int fontSize = 12)
        {
            headingBar.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headingBar.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headingBar.WrapText = false;
            headingBar.Orientation = 0;
            headingBar.AddIndent = false;
            headingBar.IndentLevel = 0;
            headingBar.ShrinkToFit = false;
            headingBar.ReadingOrder = (int)Excel.Constants.xlContext;
            headingBar.MergeCells = true;
            headingBar.RowHeight = height;
            headingBar.Font.Size = fontSize;
            headingBar.Font.FontStyle = "Microsoft YaHei UI";
            headingBar.Select();
            headingBar.FormulaR1C1 = headingText;
        }

        /// <summary>
        /// Create pivot table template
        /// </summary>
        /// <param name="idx">current index</param>
        /// <param name="curSheet">current pivot sheet</param>
        /// <param name="topLeft">the table's topleft</param>
        /// <param name="chartTopLeft">the chart's topleft</param>
        /// <param name="bottomRight">the chart's topleft</param>
        /// <param name="source">the table's source data range</param>
        /// <param name="charType">chart type, Excel.XlChartType</param>
        public static void createPivotTable(int idx, Excel.Worksheet curSheet, Excel.Range topLeft, string title,
            Excel.Range chartTopLeft, Excel.Range bottomRight, Excel.Range source, Excel.XlChartType charType, int layOutParam)
        {
            Excel.PivotTable table = workbook.PivotCaches().Create(SourceType: Excel.XlPivotTableSourceType.xlDatabase,
                SourceData: source, Version: 6).CreatePivotTable(TableDestination: topLeft, TableName: title, DefaultVersion: 6);
            // create pivot chart, and set its position and its size
            createPivotChart(idx, curSheet, chartTopLeft, bottomRight, table.DataBodyRange, charType, title, layOutParam);
        }

        /// <summary>
        /// Add new field to pivot table
        /// </summary>
        /// <param name="table">pivot table object</param>
        /// <param name="fields">field list</param>
        public static void addField(Excel.PivotTable table, List<PivotData> fields)
        {
            Excel.PivotField field = null;
            foreach (PivotData item in fields)
            {
                switch (item.type)
                {
                    case 0: // row
                        field = table.PivotFields(item.name) as Excel.PivotField;
                        field.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                        field.Position = 1;
                        break;
                    case 1: // column
                        field = table.PivotFields(item.name) as Excel.PivotField;
                        field.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                        field.Position = 1;
                        break;
                    case 2: // data
                        field = table.AddDataField(table.PivotFields(item.name), Type.Missing, Excel.XlConsolidationFunction.xlSum);
                        break;
                    default:
                        break;
                }
                // if field is not null and field's data type is xlDate
                if (null != field && field.DataType == Excel.XlPivotFieldDataType.xlDate)
                {
                    Excel.Range rng = field.DataRange.Cells[1] as Excel.Range;
                    bool[] periods = new bool[7] { false, false, false, false, false, false, false };
                    periods[item.add_in] = true;
                    rng.Group(true, true, Type.Missing, periods);
                }
            }
            //table.PivotCache().Refresh();
        }

        public static void createPivotChart(int idx, Excel.Worksheet curSheet, Excel.Range topLeft, Excel.Range bottomRight, Excel.Range source, Excel.XlChartType charType, string chartName, int layOutParam)
        {
            curSheet.Shapes.AddChart2(201, charType).Select();
            workbook.ActiveChart.ChartTitle.Text = chartName;
            workbook.ActiveChart.SetSourceData(Source: source);
            workbook.ActiveChart.ApplyLayout(layOutParam);
            Excel.Shape baseShape = curSheet.Shapes.Item(idx);
            baseShape.Top = (float)topLeft.Top;
            baseShape.Left = (float)topLeft.Left;
            baseShape.Height = (float)bottomRight.Top - baseShape.Top;
            baseShape.Width = (float)bottomRight.Left - baseShape.Left;
        }
    }
}
