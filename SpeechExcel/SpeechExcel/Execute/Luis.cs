using System.Collections.Generic;
using Microsoft.Cognitive.LUIS;
using System.Windows;
using System;

namespace SpeechExcel.Execute
{
    class Luis
    {
        private string LuisAppId
        {
            //get { return "00562bbb-2a3a-4e36-afa9-ccb398c7a103"; } // word
            get { return "1284fc06-1d2a-4dad-a49d-0aa0086af56c"; } // yang
            //get { return "37b0dffd-5511-4975-bd9d-b4d2530b5bc0"; } // my excel
        }

        /// <summary>
        /// Gets the LUIS subscription identifier.
        /// </summary>
        /// <value>
        /// The LUIS subscription identifier.
        /// </value>
        private string LuisSubscriptionID
        {
            get { return "2f8b5ad36e6e44a2889702101e5e26bf"; }
        }

        public async void predict(string text, SpeechUserControl UI)
        {
            List<Parser.ReplaceNode> replace_list = null;
            string replaced_text = text;
            try
            {
                replaced_text = Parser.replace(text, out replace_list);
            }
            finally
            {
                try
                {
                    LuisClient client = new LuisClient(LuisAppId, LuisSubscriptionID);
                    LuisResult res = await client.Predict(replaced_text);
                    UI.MessageShow = processRes(res, replace_list);
                }
                catch(Exception e)
                {
                    MessageBox.Show("Some Error has happend at luis.cs: " + e.ToString());
                }
                finally
                {
                    UI.done = true;
                    UI.ProcessStatus = Visibility.Hidden;
                }
            }
        }

        public string processRes(LuisResult res, List<Parser.ReplaceNode> replace_list)
        {
            return Caller.CallFunc(res, replace_list);
        }
    }
}
