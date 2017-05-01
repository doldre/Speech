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
            { "PivotCreate", Pivot.CreatePivot }, { "AddColumnToPivot", Pivot.AddColumn }, { "ModiFunc", Pivot.ChangeFunc },
                { "Find_Min_Max", SheetOpe.find_min_max }, { "Get_Value", SheetOpe.get_value }, { "Sort", SheetOpe.sort },
                {"Filter",SheetOpe.filter }, {"CancelFilter",SheetOpe.cancelFilter }
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
                MessageBox.Show("haha");
                intentExe[res.Intents[0].Name](res, replace_list);
            }
            catch
            {
                MessageBox.Show("Cannot Parse this Intent: " + res.Intents[0].Name);
            }

        }
    }
}
