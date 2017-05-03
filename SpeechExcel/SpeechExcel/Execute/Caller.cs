using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using Microsoft.Cognitive.LUIS;

namespace SpeechExcel.Execute
{

    static class Caller
    {
        public static Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        /// <summary>
        /// Property intent: Pivot
        /// </summary>
        public static Dictionary<string, Action<LuisResult, List<Parser.ReplaceNode>>> intentExe =
            new Dictionary<string, Action<LuisResult, List<Parser.ReplaceNode>>>()
            {
                //TODO: 添加你的意图函数集映射(intent => function set)
                { "PivotCreate", Pivot.CreatePivot },   // 创建透视图，可指定选中列
                { "AddColumnToPivot", Pivot.AddColumn },    // 向透视表中添加列，可指定区域
                { "ModiFunc", Pivot.ChangeFunc },   // 修改透视表统计函数
                { "Find_Min_Max", SheetOpe.find_min_max },  // 最值
                { "Get_Value", SheetOpe.get_value },    // 查找
                { "Sort", SheetOpe.sort },  // 排序
                { "Filter",SheetOpe.filter },   // 过滤，筛选数据
                { "CancelFilter",SheetOpe.cancelFilter }    // 取消筛选
            };
        
        /// <summary>
        /// 根据意图调用相应的函数
        /// </summary>
        /// <param name="res">Luis的解析结果</param>
        /// <param name="replace_list"></param>
        public static void CallFunc(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            try
            {
                intentExe[res.Intents[0].Name](res, replace_list);
            }
            catch
            {
                MessageBox.Show("Cannot Parse this Intent: " + res.Intents[0].Name);
            }

        }
    }
}
