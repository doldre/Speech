using System;
using System.Collections.Generic;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace SpeechExcel.Execute
{
    static class OrdinaryChart
    {
        public static Dictionary<string, Excel.XlChartType> chartMap = new Dictionary<string, Excel.XlChartType>()
        {
            { "column", Excel.XlChartType.xl3DColumn },
            { "surface", Excel.XlChartType.xlSurface },
            { "scatter", Excel.XlChartType.xlXYScatter },
            { "pie", Excel.XlChartType.xlPie },
            { "bubble", Excel.XlChartType.xlBubble },
        };

        private static void funcs(JArray entities)
        {
            Excel.Worksheet s = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            return;
        }

        /// <summary>
        /// 创建图表
        /// </summary>
        /// <param name="rng">图表数据来源</param>
        /// <param name="type">图类型</param>
        /// <param name="x1">位置</param>
        /// <param name="y1">位置</param>
        /// <param name="x2">位置</param>
        /// <param name="y2">位置</param>
        public static void CreateChart(Excel.Range rng, Excel.XlChartType chartType, int x1 = 300, int y1 = 200, int x2 = 500, int y2 = 350)
        {
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)worksheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = xlCharts.Add(x1, y1, x2, y2);
            Excel.Chart chartPage = myChart.Chart;
            // 图表数据来源
            //chartRange = rng;
            chartPage.SetSourceData(rng, Excel.XlRowCol.xlColumns);

            // 图表类型
            chartPage.ChartType = chartType;
        }

        /// <summary>
        /// 更换图标类型
        /// </summary>
        /// <param name="oriChart">必须是Active的chart</param>
        /// <param name="newType">新的图类型</param>
        public static void ModifyChart(Excel.Chart oriChart, Excel.XlChartType newType)
        {
            try
            {
                oriChart.ChartType = newType;
                oriChart.Refresh();
            }
            catch
            {
                MessageBox.Show("请选择你要修改的图表！");
            }

        }
    }
}
