using System;
using System.Collections.Generic;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using Microsoft.Cognitive.LUIS;

namespace SpeechExcel.Execute
{
    static class OrdinaryChart
    {
        public static Dictionary<string, Excel.XlChartType> chartMap = new Dictionary<string, Excel.XlChartType>()
        {
            { "3Dcolumn", Excel.XlChartType.xl3DColumn },
            { "pie", Excel.XlChartType.xlPie },
            { "line", Excel.XlChartType.xlLineMarkers }
        };

        /// <summary>
        /// Create ordinary chart
        /// </summary>
        /// <param name="res"></param>
        /// <param name="dataList"></param>
        public static void CreateChart(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            Excel.XlChartType chartType = Excel.XlChartType.xlColumnStacked;
            string rangeBlock = "", selectType = "column";
            // 使用[col]和[col]的数据绘制成（[charType]）
            foreach (Entity item in res.GetAllEntities())
            {
                if (item.Name == "builtin.ordinal")
                {
                    int col = Parser._toInt(item.Value.Substring(1));
                    char nameCol = (char)(64 + col);
                    rangeBlock += nameCol + ":" + nameCol + ",";
                }
                else if (item.Name.Contains("chartype"))
                {
                    chartType = chartMap[item.Name.Split(':')[2]];  // substract the chart type
                }
            }
            // check whether we find rows
            if (rangeBlock == "")
            {
                selectType = "row";
                foreach (var cell in dataList)
                {
                    if (cell.Row != 1)
                    {
                        // transfer to row
                        char nameRow = (char)(48 + cell.Row);
                        rangeBlock += nameRow + ":" + nameRow + ",";
                    }
                    else if (cell.Row == 1)
                    {
                        // transfer to column
                        char nameCol = (char)(64 + cell.Column);
                        rangeBlock += nameCol + ':' + nameCol + ',';
                    }
                }
            }
            
            if (rangeBlock == "")
            {
                MessageBox.Show("Sorry, I cannot substract any data range to plot a chart.");
                return;
            }
            // plot chart
            createChart(selectType, rangeBlock.Substring(0, rangeBlock.Length - 1), chartType);
        }

        public static void ModifyChart(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            Excel.XlChartType chartType = Excel.XlChartType.xlColumnStacked;
            // 使用[col]和[col]的数据绘制成（[charType]）
            foreach (Entity item in res.GetAllEntities())
            {
                if (item.Name.Contains("chartType"))
                {
                    chartType = chartMap[item.Name.Split(':')[2]];  // substract the chart type
                }
            }
            // check whether we find rows 
            Excel.Workbook curWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            // replace slection chart with a new chart type
            modifyChart(curWorkbook.ActiveChart, chartType);
        }

        /// <summary>
        /// Plot chart by select range
        /// </summary>
        /// <param name="selectType">row or column</param>
        /// <param name="rangeBlock">select range, accept like string</param>
        /// <param name="chartType">plot-chart type</param>
        private static void createChart(string selectType, string rangeBlock, Excel.XlChartType chartType)
        {

            Boolean oldFresh = Globals.ThisAddIn.Application.ScreenUpdating;
            Excel.XlRowCol plotby = Excel.XlRowCol.xlRows;

            if (selectType == "column") plotby = Excel.XlRowCol.xlColumns;

            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                Excel.Worksheet curSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                curSheet.Shapes.AddChart2(227, chartType).Select();
                workbook.ActiveChart.SetSourceData(Source: curSheet.UsedRange.Range[rangeBlock], PlotBy: plotby);
                curSheet.Shapes.Item(1).ScaleWidth(1.4f, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);
                curSheet.Shapes.Item(1).ScaleHeight(1.6f, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);

            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.ToString());
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }
        }

        /// <summary>
        /// 更换图标类型
        /// </summary>
        /// <param name="oriChart">必须是Active的chart</param>
        /// <param name="newType">新的图类型</param>
        private static void modifyChart(Excel.Chart oriChart, Excel.XlChartType newType)
        {
            Boolean oldFresh = Globals.ThisAddIn.Application.ScreenUpdating;

            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                oriChart.ChartType = newType;
                oriChart.Refresh();
            }
            catch
            {
                MessageBox.Show("请选择你要修改的图表！");
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }

        }
    }
}
