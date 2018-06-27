using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GoldCFTCWeekly
{
    using Microsoft.Office.Interop.Excel;
    using System.Runtime.InteropServices;
    
    public class ExcelOperator
    {
        private DataFetch _DFetch;
        public ExcelOperator()
        {
            FilePath = "CFTC.xlsx";
            InitExcelApplication();
            _DFetch = new DataFetch();
        }
        public string FilePath { get; }
        private Application excelApp;
        private Workbook xlsWorkbook;
        private readonly string DataSheet = "CFTC";
        private readonly string AnaySheet = "Analysis";
        private readonly int firstRow = 6;
        private int firstCol = 12;
        private Worksheet xlsWorkSheet;
        private Worksheet anayWorkSheet;
        
        private int _rowsCount;
        private int _colsCount;
        private void InitExcelApplication()
        {
            //excelApp = new Application();
            //excelApp.Visible = false;
            //excelApp.DisplayAlerts = false;
            //xlsWorkbook = excelApp.Workbooks.Open(FilePath, System.Type.Missing, System.Type.Missing, System.Type.Missing,
            //        System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
            //        System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
            //        System.Type.Missing);
            excelApp = Globals.Sheet1.Application;
            xlsWorkbook = excelApp.ActiveWorkbook;
            xlsWorkSheet = xlsWorkbook.Worksheets[DataSheet];
           
           // xlsWorkSheet = (Worksheet)xlsWorkbook.Worksheets[DataSheet];
            _rowsCount = xlsWorkSheet.UsedRange.Rows.Count;
            _colsCount = xlsWorkSheet.UsedRange.Columns.Count;
        }

        public bool IsNeedUpdate()
        {
            DateTime dt = getWeekUpOfDate(DateTime.Today, DayOfWeek.Tuesday, -1);
            var temp = ((Range)xlsWorkSheet.Cells[6, 1]).Value;
            bool ret = DateTime.Equals(dt, temp);
            return !ret;
        }

        public DateTime getWeekUpOfDate(DateTime dt, DayOfWeek weekday, int Number)
        {
            int wd1 = (int)weekday;
            int wd2 = (int)dt.DayOfWeek;
            return wd2 == wd1 ? dt.AddDays(7 * Number) : dt.AddDays(7 * Number - wd2 + wd1);
        }
        private static object lockObj = new object();

        public void CheckAndUpdate()
        {

            while (IsNeedUpdate())
            {
                DateTime latest = getWeekUpOfDate(DateTime.Today, DayOfWeek.Tuesday, -1);

                lock (lockObj)
                {
                    DateTime dt;
                    dt = ((Range) xlsWorkSheet.Cells[6, 1]).Value;
                    dt = dt.AddDays(7);
                    if (dt >=latest)
                        break;
                    _DFetch.FetchData(out List<int> lst, ref dt);
                    if (lst.Count == 0)
                    {
                        //MessageBox.Show("Fetch website data error.");
                        int offset = -1;
                        _DFetch.FetchData(out lst,ref dt, offset);
                        if (lst.Count == 0)
                        {
                            offset = 0;
                            _DFetch.FetchData(out lst,ref dt, offset);
                            if (lst.Count == 0)
                            {
                                offset = 1;
                                _DFetch.FetchData(out lst, ref dt, offset);
                                if (lst.Count == 0)
                                {
                                    MessageBox.Show("Fetch website data error.");
                                    break;
                                }
                            }
                        }
                    }

                    //lst.ForEach(x => Debug.Write(x.ToString() + " "));
                    //Debug.WriteLine("");

                    UpdateData(ref lst,dt);
                }

            }
            xlsWorkbook.Save();
        }


        public bool UpdateData(ref List<int> lst, DateTime dt)
        {
            var rang = (Range) xlsWorkSheet.Rows[firstRow, Type.Missing];
        
            rang.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            int colDiff = 0;
            foreach (int data in lst)
            {
                xlsWorkSheet.Cells[firstRow, firstCol + colDiff] = data;
                colDiff++;
            }
            xlsWorkSheet.Cells[firstRow, 1] = dt;
            colDiff = 0;
            for (int index = 2; index < 10; index++)
            {
                string formula = xlsWorkSheet.Rows[firstRow + 1].Columns[index].Formula;
                xlsWorkSheet.Cells[index][firstRow] = excelApp.ConvertFormula(formula,
                    XlReferenceStyle.xlA1, XlReferenceStyle.xlR1C1,
                    XlReferenceType.xlRelative, xlsWorkSheet.Cells[index][firstRow + 1]);
            }
            AddAnaySheet();
            return true;
            

        }

        private void AddAnaySheet()
        {
            anayWorkSheet = Globals.ThisWorkbook.Worksheets[AnaySheet];

            var rang = (Range)anayWorkSheet.Rows[firstRow, Type.Missing];
            rang.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            for (int index = 2; index < 11; index++)
            {
                string formula = anayWorkSheet.Rows[firstRow + 1].Columns[index].Formula;
                anayWorkSheet.Cells[index][firstRow] = excelApp.ConvertFormula(formula,
                    XlReferenceStyle.xlA1, XlReferenceStyle.xlR1C1,
                    XlReferenceType.xlRelative, anayWorkSheet.Cells[index][firstRow + 1]);
            }
        }

        ///// <summary>
        ///// 关闭Excel进程
        ///// </summary>
        //public class KeyMyExcelProcess
        //{
        //    [DllImport("User32.dll", CharSet = CharSet.Auto)]
        //    public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        //    public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        //    {
        //        try
        //        {
        //            IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口
        //            int k = 0;
        //            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
        //            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
        //            p.Kill();     //关闭进程k
        //        }
        //        catch (System.Exception ex)
        //        {
        //            throw ex;
        //        }
        //    }
        //}


        ////关闭打开的Excel方法
        //public void CloseExcel(Application ExcelApplication, Workbook ExcelWorkbook)
        //{
        //    ExcelWorkbook.Close(false, Type.Missing, Type.Missing);
        //    ExcelWorkbook = null;
        //    ExcelApplication.Quit();
        //    GC.Collect();
        //    KeyMyExcelProcess.Kill(ExcelApplication);
        //}
    }
}
