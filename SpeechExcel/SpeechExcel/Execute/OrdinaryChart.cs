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
        private static string mess;
        public static Dictionary<string, Excel.XlChartType> chartMap = new Dictionary<string, Excel.XlChartType>()
        {
            { "3Dcolumn", Excel.XlChartType.xlColumnClustered },
            { "pie", Excel.XlChartType.xlPie },
            { "line", Excel.XlChartType.xlLineMarkers }
        };

        /// <summary>
        /// Create ordinary chart
        /// </summary>
        /// <param name="res"></param>
        /// <param name="dataList"></param>
        public static string CreateChart(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            mess = "";
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
            // 如果绘制的是透视图，调用pivot中的OriChart接口
            if (chartType == chartMap["pivot"])
            {
                List<int> idx = new List<int>();
                foreach (var cell in dataList)
                {
                    if (cell.Row == 1)
                    {
                        // transfer to column
                        idx.Add(cell.Column);
                    }
                }
                mess = Pivot.OriInterFace(idx, chartType);
            }
            else if (rangeBlock == "") // 如果没有找到列，那么使用datalist给出的列
            {
                selectType = "row";
                foreach (var cell in dataList)
                {
                    if (cell.Row != 1)// 如果不是行号不是1，说明这是一个筛选过程
                    {
                        // transfer to row
                        char nameRow = (char)(48 + cell.Row);
                        rangeBlock += nameRow + ":" + nameRow + ",";
                    }
                    else if (cell.Row == 1)//如果行号是1，说明这是一个列
                    {
                        // transfer to column
                        char nameCol = (char)(64 + cell.Column);
                        rangeBlock += nameCol + ":" + nameCol + ",";
                    }
                }
                if (rangeBlock == "") mess = "抱歉，我不能提取到有效的数据来绘制图表。";
                else createChart(selectType, rangeBlock.Substring(0, rangeBlock.Length - 1), chartType);
            }

            return mess;
        }

        /// <summary>
        /// Modify Current Chart to another chart
        /// </summary>
        /// <param name="res"></param>
        /// <param name="dataList"></param>
        /// <returns></returns>
        public static string ModifyChart(LuisResult res, List<Parser.ReplaceNode> dataList)
        {
            mess = "";
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
            return mess;
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
                //MessageBox.Show("Error: " + e.ToString());
                MessageBox.Show("error at Orichart: " + e.ToString());
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
                //MessageBox.Show("请选择你要修改的图表！");
                mess = "请选择你要修改的图表！";
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = oldFresh;
            }

        }
    }
}
