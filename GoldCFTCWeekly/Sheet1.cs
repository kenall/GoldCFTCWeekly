using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace GoldCFTCWeekly
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            //IsNeedUpdate();
            eo = new ExcelOperator();
            //eo.IsNeedUpdate();
            //var task = new Task(() => { eo.CheckAndUpdate(); });
            button1.Visible = eo.IsNeedUpdate();
            //button1.Enabled = true;
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        private ExcelOperator eo;
        private bool IsNeedUpdate()
        {
            //check date in A6
            DateTime today = DateTime.Now.Date;
            DateTime lastestTuesday = GetLastTuesday(today);
            var dateInRecord = Globals.Sheet1.Range["A6"].Value2;
            return false;
        }

        private DateTime GetLastTuesday(DateTime iDate)
        {
            DateTime lastTuesday = new DateTime();
            int iweek = Convert.ToInt32(iDate.DayOfWeek);
            int delta;
            iweek = (iweek == 0 )?7:iweek;
            if (iweek >5)
                lastTuesday = iDate.AddDays(-5+7-iweek);
            else
                lastTuesday = iDate.AddDays(-5-iweek);

            return lastTuesday;
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.Update_Click);
            this.表8.Change += new Microsoft.Office.Tools.Excel.ListObjectChangeHandler(this.表8_Change);
            this.BeforeDoubleClick += new Microsoft.Office.Interop.Excel.DocEvents_BeforeDoubleClickEventHandler(this.Sheet1_BeforeDoubleClick);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private static object _LockObj = new object();
        private void Sheet1_BeforeDoubleClick(Excel.Range Target, ref bool Cancel)
        {

        }

        private void Update_Click(object sender, EventArgs e)
        {
            
            var task = new Task(() => { eo.CheckAndUpdate(); });
            task.Start();
            button1.Enabled = false;
            button1.Visible = false;
        }

        private void 表8_Change(Excel.Range targetRange, ListRanges changedRanges)
        {

        }
    }
}
