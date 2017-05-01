using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using Microsoft.Cognitive.LUIS;

namespace SpeechExcel.Execute
{
    class Pivot
    {
        /// <summary>
        /// 图标名称和chartype的映射
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
        public static void CreatePivot(LuisResult res, List<Parser.ReplaceNode> dataList)
        {

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
                MessageBox.Show("Sorry, I do not know what you wanna do.");
                return;
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
        }

        /// <summary>
        /// Add more column to pivot table, 
        /// </summary>
        /// <param name="res"></param>
        /// <param name="dataList"></param>
        public static void AddColumn(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
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
                MessageBox.Show("Cannot add any column!");
                return;
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
            
        }

        /// <summary>
        /// 修改某个列值的统计方式，前提是这个列值得是DataFeild
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="queryText"></param>
        public static void ChangeFunc(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            // 将[col]的统计方式[modify]{成}[functionName]
            MessageBox.Show("Happy for modify Pivot function!");
        }

        public static string ConvertColName(List<int> idxes)
        {
            char iAx;
            string ans = "";
            for (int i = 0; i < idxes.Count; i++)
            {
                idxes[i] = (idxes[i] - 1) % 26;
                iAx = (char)(65 + idxes[i]);
                ans += iAx + ":" + iAx + ",";
            }
            return ans.Substring(0, ans.Length - 1);
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
            // draw pivot chart
            Globals.ThisAddIn.Application.ActiveWorkbook.ShowPivotChartActiveFields = true;
        }

        /// <summary>
        /// Add columns to chart and table
        /// </summary>
        /// <param name="colIdxes">列的序列号集合</param>
        public static void addColumns(List<int> colIdxes)
        {
            if (colIdxes.Count == 0) return;
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Excel.PivotTable table1 = Globals.ThisAddIn.Application.ActiveChart.PivotLayout.PivotTable; // get current pivot table (active)
            // Create and populate the PivotTable.
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

        /// <summary>
        /// 修改某个指定的列的统计函数，eg：
        /// </summary>
        /// <param name="colName">列名</param>
        /// <param name="newFunc">新的统计函数</param>
        public static void changeAnalysis(string colName, Excel.XlConsolidationFunction newFunc = Excel.XlConsolidationFunction.xlCount)
        {
            Excel.PivotField field = Globals.ThisAddIn.Application.ActiveChart.PivotLayout.PivotTable.PivotFields(colName) as Excel.PivotField;
            field.Caption = "Count: Level";
            field.Function = newFunc;
        }
    }
}
