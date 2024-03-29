﻿using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Cognitive.LUIS;

namespace SpeechExcel.Execute
{

    static class Caller
    {
        public static Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        /// <summary>
        /// Property intent: Pivot
        /// </summary>
        public static Dictionary<string, Func<LuisResult, List<Parser.ReplaceNode>, string>> intentExe =
            new Dictionary<string, Func<LuisResult, List<Parser.ReplaceNode>, string>>()
            {
                //TODO: 添加你的意图函数集映射(intent => function set)
                { "PivotCreate", Pivot.CreatePivot },
                { "AddColumnToPivot", Pivot.AddColumn },
                { "ModiFunc", Pivot.ChangeFunc },
                { "Find_Min_Max", SheetOpe.find_min_max },
                { "Get_Value", SheetOpe.get_value },
                { "Sort", SheetOpe.sort },
                { "Filter", Filter.filter },
                { "CancelFilter", Filter.cancelFilter },
                { "OriChart", OrdinaryChart.CreateChart },
                { "Analysis", Other.UseTemplate },
                { "Get_Sum", SheetOpe.get_sum }
            };
        
        /// <summary>
        /// 根据意图调用相应的函数
        /// </summary>
        /// <param name="res">Luis的解析结果</param>
        /// <param name="replace_list"></param>
        public static string CallFunc(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            try
            {
                if (res.Intents[0].Name == "None" || res.Intents[0].Score <= 0.5)
                {
                    return Properties.Resources.unkown;
                }
                string mss = intentExe[res.Intents[0].Name](res, replace_list);
                //return "你的意图：" + res.Intents[0].Name + "\n" + mss;
                return mss;
            }
            catch
            {
                return Properties.Resources.bug;
            }
        }
    }
}
