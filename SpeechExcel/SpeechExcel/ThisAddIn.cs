using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace SpeechExcel
{
    public partial class ThisAddIn
    {
        private SpeechUserControl userControl;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            userControl = new SpeechUserControl();
            System.Windows.Forms.UserControl baseCtr = new System.Windows.Forms.UserControl();
            baseCtr.Width = 600;
            System.Windows.Forms.Integration.ElementHost host = new System.Windows.Forms.Integration.ElementHost();
            host.Dock = System.Windows.Forms.DockStyle.Fill;
            host.Child = userControl;

            baseCtr.Controls.Add(host);
            Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane = this.CustomTaskPanes.Add(baseCtr, "Speech");
            myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            myCustomTaskPane.Width = 350;
            myCustomTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
