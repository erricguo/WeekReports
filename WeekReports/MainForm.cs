using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Skins;
using DevExpress.XtraGrid.Columns;
using Nini.Config;
using SharpUpdate;
using System.Reflection;
using System.Diagnostics;
using DevExpress;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using OfficeOpenXml.DataValidation;

namespace WeekReports
{
    public partial class WeekReportF : DevExpress.XtraEditors.XtraForm, ISharpUpdatable
    {
        public class TMatchDay 
        {
            public int mday;
            public string mtime;
            public int mindex;
            public TMatchDay(int xday,string xtime,int xindex)
            {
                mday = xday;
                mtime = xtime;
                mindex = xindex;
            }       
        }
        string FclassDefaultPath = "C:\\";
        string FreportDefaultPath = "C:\\";
        string FsaveasDefaultPath = "C:\\";
        string Fteacher = "";
        private BackgroundWorker bgWorker;
        DateTime FDate = DateTime.Now;
        private SharpUpdater update;
        int FDaysinMonth = 30;
        List<int> FModiDays = new List<int>();
        List<TMatchDay> FMatchDays = new List<TMatchDay>();
        List<int> DaysinWeek = new List<int>();
        List<string> FClassType = new List<string>()
        {
            "",
            "1大班課程(含雲端)",
            "2學校包班-上課",
            "3學校包班-交通",
            "4驗收與被驗收 ",
            "5教材開發",
            "6訓練\\新課程備課",
            "7專案\\派工",
            "8WEB案件",
            "9加值課程",
            "10SM自主上線",
            "11其他行政",
            "12請假"
        };
        DataTable FDT = new DataTable();

        public WeekReportF()
        {
            
            InitializeComponent();
            FDT.Columns.Add("日期", typeof(System.Int32));
            FDT.Columns.Add("星期");
            FDT.Columns.Add("時間");
            FDT.Columns.Add("內容");
            FDT.Columns.Add("講師");
            FDT.Columns.Add("類別");
            GVMain.BestFitColumns();
            //GVMain.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void btnSaveToReport_Click(object sender, EventArgs e)
        {
            

        }

        public bool IsOpenedFile(string file)
        {
            bool result = false;
            try
            {
                FileStream fs = File.OpenWrite(file);
                fs.Close();
            }
            catch(Exception ex)
            {
                result = true;
            }
    
            return result;
        }

        private void cboTeacher_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboTeacher.Text == "ALL")
                FDT.DefaultView.RowFilter = "";
            else
                FDT.DefaultView.RowFilter = string.Format("講師 LIKE '%{0}%'", cboTeacher.Text);
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw = Stopwatch.StartNew();
            int errn = 0;
            int errm = 0;
            try
            {
                splashScreenManager1.ShowWaitForm();
                FDT.Rows.Clear();
                cboTeacher.Properties.Items.Clear();
                cboTeacher.Properties.Items.Add("ALL");

                FDaysinMonth = DateTime.DaysInMonth(DEMonth.DateTime.Year, DEMonth.DateTime.Month);

                //DayOfWeek 列舉的常數，表示一週天數。
                //這個屬性值的範圍從 0 開始 (表示星期日) 到 6 (表示星期六)
                DaysinWeek.Clear();
                int mdays = 0;
                for (int i = 1; i <= FDaysinMonth; i++)
                {
                    int mdow = (int)(new DateTime(2018, 05, i).DayOfWeek);
                    mdays++;
                    if (mdow == 0 || i == FDaysinMonth)
                    {
                        DaysinWeek.Add(mdays);
                        mdays = 0;
                    }
                }

                ////建立Excel 2007檔案
                IWorkbook etbook = new XSSFWorkbook(tbFrom.Text);
                ISheet etsheet = null;
                IFormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator(etbook);
               
                //FileStream fs = new FileStream(tbFrom.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                //ExcelPackage ep = new ExcelPackage(fs);
                //ExcelWorkbook etbook = ep.Workbook;
                //ExcelWorksheet etsheet = null;
                //ExcelRange etrange = null;
                /*ET.Application etApp = null;
                ET.Workbook etbook = null;
                ET.Worksheet etsheet = null;
                ET.Range etrange = null;
                //获取工作表表格
                etApp = new ET.Application();
                etbook = (ET.Workbook)etApp.Workbooks.Open(tbFrom.Text);*/

                string mYear = (DEMonth.DateTime.Year - 1911).ToString();
                string mMonth = DEMonth.DateTime.Month.ToString().PadLeft(2,'0');
                //获取数据区域
                for (int m = 0; m < etbook.NumberOfSheets; m++)
                {

                    if (etbook.GetSheetAt(m).SheetName.Trim() == mYear + "." + mMonth)
                    {
                        etsheet = etbook.GetSheetAt(m);
                        //break;
                    }
                }

                /*int startRowNumber = etsheet.Dimension.Start.Row;//起始列編號，從1算起
                int endRowNumber = etsheet.Dimension.End.Row;//結束列編號，從1算起
                int startColumn = etsheet.Dimension.Start.Column;//開始欄編號，從1算起
                int endColumn = etsheet.Dimension.End.Column;//結束欄編號，從1算起*/
                int endColumn = etsheet.GetRow(1).Cells.Count;

                //etrange = CellRangeAddress.ValueOf(range)[startColumn, startRowNumber, endColumn, endRowNumber];
                //4. 读取某单元格的数据内容：
                int MaxColumns = 0;
                for (int m = 0; m < endColumn; m++)
                {
                    if (etsheet.GetRow(1).GetCell(m).StringCellValue.Trim() == "")
                    {
                        MaxColumns = m;
                        break;
                    }
                }

                for (int n = 2; n < (FDaysinMonth) * 2 + 2; n++)
                {
                    errn = n;

                    string sDate = GetValue(formulaEvaluator, etsheet.GetRow(n).GetCell(0));
                    string sWeek = GetValue(formulaEvaluator, etsheet.GetRow(n).GetCell(1));
                    string sTime = GetValue(formulaEvaluator, etsheet.GetRow(n).GetCell(2));
                    for (int m = 3; m <= MaxColumns - 2; m = m + 2)
                    {
                        errm = m;
                        string mTitle = GetValue(formulaEvaluator,etsheet.GetRow(1).GetCell(m)).ToUpper().Trim();
                        string mClass = GetValue(formulaEvaluator,etsheet.GetRow(n).GetCell(m)).Trim();
                        string mTeacher = GetValue(formulaEvaluator,etsheet.GetRow(n).GetCell(m+1)).Trim();

                        if (mClass == "" && mTeacher == "")
                        {
                            continue;
                        }

                        if (mTeacher.ToUpper() == "X")
                        {
                            continue;
                        }

                        DataRow row = FDT.NewRow();
                        row["日期"] = sDate;
                        row["星期"] = sWeek;
                        row["時間"] = sTime;
                        row["內容"] = mClass;
                        row["講師"] = mTeacher;

                        if (!cboTeacher.Properties.Items.Contains(mTeacher) && !(mTeacher.Trim() == ""))
                        {
                            cboTeacher.Properties.Items.Add(mTeacher);
                        }

                        if (mTitle.Contains("教室") || mTitle.Contains("遠距"))
                        {
                            row["類別"] = "1";
                        }
                        else if (mTitle.Contains("就業") || mTitle.Contains("包班"))
                        {
                            row["類別"] = "2";
                        }
                        else if (mTitle.Contains("驗收"))
                        {
                            row["類別"] = "4";
                        }
                        else if (mTitle.Contains("聽課"))
                        {
                            row["類別"] = "6";
                        }
                        else if (mTitle.Contains("專案"))
                        {
                            row["類別"] = "7";
                        }
                        else if (mTitle.Contains("WEB"))
                        {
                            row["類別"] = "8";
                        }
                        else if (mTitle.Contains("加值"))
                        {
                            row["類別"] = "9";
                        }
                        else if (mTitle.Contains("教室維護") || mTitle.Contains("教育訓練"))
                        {
                            row["類別"] = "11";
                        }
                        else if (mTitle.Contains("休假"))
                        {
                            row["類別"] = "12";
                        }
                        else
                        {
                            row["類別"] = "";
                        }

                        FDT.Rows.Add(row);
                    }

                }
                
                GCMain.DataSource = FDT;
                GVMain.Columns[0].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                GVMain.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                GVMain.Columns[2].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                GVMain.Columns[3].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                GVMain.Columns[4].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                GVMain.Columns[5].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;

                foreach (GridColumn col in GVMain.Columns)
                {
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                }

                GVMain.BestFitColumns();
                GVMain.BestFitColumns();
                
                try
                {
                    cboTeacher.SelectedIndex = 0;
                    IConfigSource source = new IniConfigSource(fc.INIPath);
                    IConfig config = source.Configs["SYSTEM"];

                    string mteacher = config.GetString("TEACHER");
                    if (cboTeacher.Properties.Items.Contains(mteacher))
                    {
                        cboTeacher.EditValue = mteacher;
                    }
                }
                catch (System.Exception ex)
                {
                    ErrorLog("載入變數TEACHER發生錯誤.." + ex.Message);
                }


                //建立檔案
                

                /*etbook.Close(ET.XlSaveAction.xlDoNotSaveChanges, Type.Missing, Type.Missing);
                etApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(etrange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(etsheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(etbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(etApp);*/
                etbook.Close();               
                splashScreenManager1.CloseWaitForm();
                sw.Stop();
                Msg(DateTime.Now.ToString() + " 匯出資料完成\r\n耗時：" + sw.Elapsed.Seconds + "秒", "提示",false);
                
            }
            catch (System.Exception ex)
            {
                fc.ErrorLog("n=" + errn.ToString() + ",m=" + errm.ToString());
                splashScreenManager1.CloseWaitForm();
                sw.Stop();
                Msg(DateTime.Now.ToString() + ex.Message + "\r\n耗時：" + sw.Elapsed.Seconds + "秒", "錯誤");
            }
            finally
            {
                
            }
        }

        private string GetValue(IFormulaEvaluator xformulaEvaluator, ICell xCell)
        {
            string columnStr = string.Empty;

            if (xCell == null)
            {
                return columnStr;
            }

            switch (xCell.CellType)
            {
                case CellType.Numeric:  // 數值格式
                    if (DateUtil.IsCellDateFormatted(xCell))
                    {   // 日期格式
                        columnStr = xCell.DateCellValue.ToString();
                    }
                    else
                    {   // 數值格式
                        columnStr = xCell.NumericCellValue.ToString();
                    }
                    break;
                case CellType.String:   // 字串格式
                    columnStr = xCell.StringCellValue;
                    break;
                case CellType.Formula:  // 公式格式
                    var formulaValue = xformulaEvaluator.Evaluate(xCell);
                    if (formulaValue.CellType == CellType.String) columnStr = formulaValue.StringValue.ToString();          // 執行公式後的值為字串型態
                    else if (formulaValue.CellType == CellType.Numeric) columnStr = formulaValue.NumberValue.ToString();    // 執行公式後的值為數字型態
                    break;
                default:
                    break;
            }
            return columnStr;
        }



        private void btnSave_Click(object sender, EventArgs e)
        {
            /*if (File.Exists(tbSaveAs.Text))
            {
                File.Delete(tbSaveAs.Text);
            }*/
            if (GVMain.RowCount <= 0)
            {
                Msg("請先載入課表！", "錯誤");
                return;
            }

            if (cboTeacher.EditValue.ToString() == "ALL")
            {
                Msg("請選擇一個講師！", "錯誤");
                return;
            }

            if (IsOpenedFile(tbTo.Text))
            {
                Msg("欲儲存的目標檔案目前開啟中，請先關閉！", "錯誤");
                return;
            }

            if (chkSaveAs.CheckState == CheckState.Checked)
            {
                if (tbSaveAs.Text.Trim() == "")
                {
                    Msg("另存新檔路徑不可空白", "錯誤");
                    return;
                }
            }

            try
            {
                if (rgDate.SelectedIndex == 1) //0.全月 1.日期
                {
                    if (tbDateRange.Text.Trim() == "")
                    {
                        Msg("使用區間日期，日期不可空白！", "錯誤");
                        return;
                    }
                    int mdayB = 0;
                    int mdayE = 0;
                    string[] daysarray = tbDateRange.Text.Split(',');
                    for (int i = 0; i <= daysarray.Length - 1; i++)
                    {
                        mdayB = -1;
                        mdayE = -1;
                        if (daysarray[i].Contains('~'))
                        {
                            string[] daysrange = daysarray[i].Split('~');

                            if (Int32.TryParse(daysrange[0], out mdayB))
                            {
                                if (Int32.TryParse(daysrange[1], out mdayE))
                                {
                                    for (int p = mdayB; p <= mdayE; p++)
                                    {
                                        if (!FModiDays.Contains(p) && (p <= FDaysinMonth))
                                        {
                                            FModiDays.Add(p);
                                        }
                                    }
                                }
                            }

                        }
                        else
                        {
                            if (Int32.TryParse(daysarray[i], out mdayB))
                            {
                                if (!FModiDays.Contains(mdayB) && (mdayB <= FDaysinMonth))
                                {
                                    FModiDays.Add(mdayB);
                                }
                            }
                        }

                    }
                }
            }
            catch (System.Exception ex)
            {
                Msg("日期格式有誤！結束匯出。  " + ex.Message , "錯誤");
                return;
            }


            FDaysinMonth = DateTime.DaysInMonth(DEMonth.DateTime.Year, DEMonth.DateTime.Month);
            FMatchDays.Clear();
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw = Stopwatch.StartNew();

            splashScreenManager1.ShowWaitForm();
            ErrorLog("========================================開始匯出========================================");
            /*ET.Application etApp = null;
            ET.Workbook etbook = null;
            ET.Worksheet etsheet = null;*/
            //ET.Range etrange = null;

            try
            {
                //DayOfWeek 列舉的常數，表示一週天數。
                //這個屬性值的範圍從 0 開始 (表示星期日) 到 6 (表示星期六)
                DaysinWeek.Clear();
                int mdays = 0;
                for (int i = 1; i <= FDaysinMonth; i++)
                {
                    int mdow = (int)(new DateTime(DEMonth.DateTime.Year, DEMonth.DateTime.Month, i).DayOfWeek);
                    mdays++;
                    if (mdow == 0 || i == FDaysinMonth)
                    {
                        DaysinWeek.Add(mdays);
                        mdays = 0;
                    }
                }

                //2018週報-EDDIE1.xlsx

                //获取工作表表格
                /*etApp = new ET.Application();
                etbook = etApp.Workbooks.Open(tbTo.Text);*/

                using (FileStream fs = new FileStream(tbTo.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    ExcelPackage ep = new ExcelPackage(fs);
                    ExcelWorkbook etbook = ep.Workbook;
                    ExcelWorksheet etsheet = null;

                    //获取数据区域
                    for (int m = 1; m <= etbook.Worksheets.Count; m++)
                    {
                        if (etbook.Worksheets[m].Name.ToString().Trim().Contains(DEMonth.DateTime.Month.ToString() + "月"))
                        {
                            etsheet = etbook.Worksheets[m];
                        }
                    }

                    string mTitleH = "";

                    if (etsheet.Cells[1, 9].Value != null)
                    {
                        mTitleH = etsheet.Cells[1, 9].Value.ToString();
                    }

                    //
                    if (mTitleH != "修改日期")
                    {
                        etsheet.InsertColumn(2, 1);
                        //5. 写入某单元格的数据内容
                        for (int i = 2; i <= FDaysinMonth; i++)
                        {
                            var b3address = GetMergedRangeAddress(etsheet.Cells[i, 6]);
                            etsheet.Cells[b3address].Merge = false;
                            etsheet.Cells[i, 6].Clear();
                            etsheet.Cells[i, 6].Style.Font.Color.SetColor(Color.Red);
                            etsheet.Cells[i, 6].Style.Font.Bold = true;
                            etsheet.Cells[i, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            etsheet.Cells[i, 7].Style.Fill.BackgroundColor.SetColor(255, 255, 255, 206);

                        }

                        for (int i = 3; i < FDaysinMonth * 2 + 3; i = i + 2)
                        {
                            etsheet.InsertRow(i, 1);
                            DateTime mdt = DateTime.FromOADate(float.Parse(etsheet.Cells[i - 1, 1].Value.ToString()));

                            etsheet.Cells[i - 1, 1].Value = new DateTime(DateTime.Now.Year, mdt.Month, mdt.Day);
                            etsheet.Cells[i - 1, 1].Style.Numberformat.Format = "m月d日";
                            mdt = (DateTime)(etsheet.Cells[i - 1, 1].Value);
                            if ((int)mdt.DayOfWeek == 0 || (int)mdt.DayOfWeek == 6)
                            {
                                etsheet.Cells[i - 1, 1].Style.Font.Color.SetColor(Color.Red);
                                etsheet.Cells[i - 1, 2].Style.Font.Color.SetColor(Color.Red);
                            }
                            else
                            {
                                etsheet.Cells[i - 1, 1].Style.Font.Color.SetColor(Color.Black);
                                etsheet.Cells[i - 1, 2].Style.Font.Color.SetColor(Color.Black);
                            }
                            etsheet.Cells[i - 1, 1].Copy(etsheet.Cells[i, 1]);
                            etsheet.Cells[i - 1, 7].Copy(etsheet.Cells[i, 7]);
                        }

                        etsheet.Cells[1, 1, FDaysinMonth * 2 + 2, 1].Copy(etsheet.Cells[1, 2, FDaysinMonth * 2 + 2, 2]);
                        etsheet.Cells[1, 2, FDaysinMonth * 2 + 2, 2].Style.Numberformat.Format = "";

                        for (int i = 0; i < GVMain.RowCount; i++)
                        {
                            int daya = Int32.Parse(GVMain.GetRowCellValue(i, "日期").ToString());
                            string timea = GVMain.GetRowCellValue(i, "時間").ToString().ToUpper().Trim();
                            for (int j = 2; j < FDaysinMonth * 2 + 2; j++)
                            {
                                int dayb = ((DateTime)etsheet.Cells[j, 1].Value).Day;

                                if (daya == dayb)
                                {
                                    bool mbool = false;
                                    if (timea == "AM" && (j % 2 == 0)) //AM在偶數位置
                                    {
                                        mbool = true;
                                        etsheet.Cells[j, 4].Value = GVMain.GetRowCellValue(i, "內容").ToString();

                                    }
                                    else if (timea == "PM" && (j % 2 == 1)) //AM在偶數位置
                                    {
                                        mbool = true;
                                        etsheet.Cells[j, 4].Value = GVMain.GetRowCellValue(i, "內容").ToString();
                                    }
                                    if (mbool)
                                    {
                                        etsheet.Cells[j, 5].Value = 4;
                                        int mtype = 0;
                                        Int32.TryParse(GVMain.GetRowCellValue(i, "類別").ToString(), out mtype);
                                        etsheet.Cells[j, 3].Value = FClassType[mtype];
                                    }
                                }
                                if (j % 2 == 0)
                                {
                                    etsheet.Cells[j, 2].Value = "AM";

                                }
                                else
                                {
                                    etsheet.Cells[j, 2].Value = "PM";
                                }

                            }


                        }

                        int mStart = 2;
                        int mEnd = 0;
                        for (int i = 0; i < DaysinWeek.Count; i++)  //FDaysinMonth可能是 28,29,30,31
                        {
                            mEnd = mStart + DaysinWeek[i] * 2 - 1;
                            string mCellA = mStart.ToString();
                            string mCellB = mEnd.ToString();
                            var etr = etsheet.Cells[mStart, 6, mEnd,6];
                            etr.Value = 0;
                            etr.Merge = true;
                            etsheet.Cells[mStart, 6].Formula = "SUM(E" + mCellA + ":E" + mCellB + ")";
                            etr.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            mStart = mEnd + 1;
                        }

                        int CharGridIndex = 0;
                        for (int i = FDaysinMonth * 2 + 2; i < 200; i++)
                        {
                            if (etsheet.Cells[i, 4].Value != null)
                            {
                                if (etsheet.Cells[i, 4].Value.ToString().Trim() == "月總工時")
                                {
                                    CharGridIndex = i;
                                    break;
                                }
                            }
                        }

                        int mEndDayCol = FDaysinMonth * 2 + 1;
                        etsheet.Cells[CharGridIndex, 1, CharGridIndex, 2].Merge = true;
                        etsheet.Cells[CharGridIndex + 1,4].Formula = "SUM($C$" + (CharGridIndex + 1).ToString() + ":$C$" + (CharGridIndex + 12).ToString() + ")";

                        for (int i = CharGridIndex + 1; i <= CharGridIndex + 12; i++)
                        {
                            etsheet.Cells[i, 3].Formula = "SUMIF($C$2:$C$" + mEndDayCol.ToString() + ",$A" + i.ToString() + ",$E$2:$E$" + mEndDayCol.ToString() +")";
                            etsheet.Cells[i, 5].Formula = "C" + i.ToString() + "/$D$"+(CharGridIndex+1).ToString();
                            etsheet.Cells[i, 1, i, 2].Merge = true;
                        }

                        etsheet.Cells[1, 2].Value = "時間";
                        etsheet.Cells[FDaysinMonth * 2 + 2, 2].Value = "";
                        etsheet.Cells[FDaysinMonth * 2 + 2, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;                        
                        etsheet.Cells[FDaysinMonth * 2 + 2, 9].Style.Fill.BackgroundColor.SetColor(255, 255, 154, 206); 

                        etsheet.Cells[1, 9].Value = "修改日期";
                        etsheet.Cells[1, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        etsheet.Cells[2, 9].Value = DateTime.Now.ToString();
                        
                        etsheet.Cells["A2:I" + (FDaysinMonth * 2 + 2).ToString()].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        etsheet.Cells["A2:I" + (FDaysinMonth * 2 + 2).ToString()].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        etsheet.Cells["A2:I" + (FDaysinMonth * 2 + 2).ToString()].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        etsheet.Cells["A2:I" + (FDaysinMonth * 2 + 2).ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        for (int i = 1; i <= 9; i++)
                        {
                            etsheet.Cells[1, i].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        }

                        etsheet.Cells["A" + (CharGridIndex - 1).ToString() + ":E" + (CharGridIndex + 12).ToString()].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        etsheet.Cells["A" + (CharGridIndex - 1).ToString() + ":E" + (CharGridIndex + 12).ToString()].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        etsheet.Cells["A" + (CharGridIndex - 1).ToString() + ":E" + (CharGridIndex + 12).ToString()].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        etsheet.Cells["A" + (CharGridIndex - 1).ToString() + ":E" + (CharGridIndex + 12).ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        etsheet.Cells["C2:C" + (FDaysinMonth * 2 + 1).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        
                        //etsheet.Cells["C2:C" + (FDaysinMonth * 2 + 1).ToString()].DataValidation;
                        for (int i=0;i<=etsheet.DataValidations.Count-1;i++)
                        {
                            if (etsheet.DataValidations[i].Address.Start.Address == "C2")
                            {
                                etsheet.DataValidations.Remove(etsheet.DataValidations[i]);
                                break;
                            }
                        }

                        // add a validation and set values
                        var validation = etsheet.DataValidations.AddListValidation("C2:C" + (FDaysinMonth * 2 + 1).ToString());
                        validation.ShowErrorMessage = true;
                        validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                        validation.ErrorTitle = "無效的值";
                        validation.Error = "請從下拉選單選取值";
                        for (int j = CharGridIndex + 1; j <= CharGridIndex + 12; j++)
                        {
                            validation.Formula.Values.Add(etsheet.Cells[j, 1].Text);
                        }
                        
                        
                    }
                    else
                    {
                        int mi = 0;
                        List<int> FDaysinMonthTmp = new List<int>();
                        for (int j = 2; j < FDaysinMonth * 2 + 2; j++)
                        {
                            //DateTime.FromOADate(float.Parse(etsheet.Cells[i - 1, 1].Value.ToString()));
                            int dayb = (DateTime.FromOADate(float.Parse(etsheet.Cells[j, 1].Value.ToString()))).Day; //週報
                            string timeb = etsheet.Cells[j, 2].Value.ToString().ToUpper().Trim();

                            bool mbool = false;

                            for (int i = mi; i < GVMain.RowCount; i++)
                            {
                                int daya = Int32.Parse(GVMain.GetRowCellValue(i, "日期").ToString()); //課表
                                string timea = GVMain.GetRowCellValue(i, "時間").ToString().ToUpper().Trim();

                                if (rgDate.SelectedIndex == 1)
                                {
                                    if (!FModiDays.Contains(daya))
                                    {
                                        mi = i + 1;
                                        break;
                                    }
                                }

                                if (daya == dayb && timea == timeb)
                                {
                                    TMatchDay tmd = new TMatchDay(dayb, timeb, j);
                                    FMatchDays.Add(tmd);


                                    if (etsheet.Cells[j, 4].Value != null)
                                    {
                                        if (GVMain.GetRowCellValue(i, "內容") != null)
                                        {
                                            if (etsheet.Cells[j, 4].Value.ToString() != GVMain.GetRowCellValue(i, "內容").ToString())
                                            {

                                                ErrorLog("資料修改：日期＝" + DEMonth.DateTime.Month.ToString() + "/" + GVMain.GetRowCellValue(i, "日期").ToString() +
                                                            ",時間＝" + GVMain.GetRowCellValue(i, "時間").ToString() +
                                                            ",原資料＝" + etsheet.Cells[j, 4].Value.ToString() +
                                                            ",修改為＝" + GVMain.GetRowCellValue(i, "內容").ToString());
                                                mbool = true;
                                                etsheet.Cells[j, 4].Value = GVMain.GetRowCellValue(i, "內容").ToString();
                                            }
                                            else
                                            {
                                                mi = i + 1;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (GVMain.GetRowCellValue(i, "內容") != null)
                                        {
                                            ErrorLog("資料修改：日期＝" + DEMonth.DateTime.Month.ToString() + "/" + GVMain.GetRowCellValue(i, "日期").ToString() +
                                                    ",時間＝" + GVMain.GetRowCellValue(i, "時間").ToString() +
                                                    ",原資料＝" +
                                                    ",修改為＝" + GVMain.GetRowCellValue(i, "內容").ToString());
                                            mbool = true;
                                            etsheet.Cells[j, 4].Value = GVMain.GetRowCellValue(i, "內容").ToString();
                                        }
                                    }


                                    if (mbool)
                                    {
                                        etsheet.Cells[j, 5].Value = 4;
                                        int mtype = 0;
                                        Int32.TryParse(GVMain.GetRowCellValue(i, "類別").ToString(), out mtype);
                                        etsheet.Cells[j, 3].Value = FClassType[mtype];
                                        mi = i + 1;
                                        break;
                                    }
                                }
                                else
                                {
                                    if (daya > dayb)
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        bool match = false;
                        for (int j = 2; j < FDaysinMonth * 2 + 2; j++)
                        {
                            match = false;
                            for (int i = 0; i <= FMatchDays.Count - 1; i++)
                            {
                                if (j == FMatchDays[i].mindex)
                                {
                                    match = true;
                                    break;
                                }
                            }
                            if (!match)
                            {
                                if (etsheet.Cells[j, 4].Value != null)
                                {
                                    ErrorLog("資料修改：日期＝" + DEMonth.DateTime.Month.ToString() + "/" +
                                                                (DateTime.FromOADate(float.Parse(etsheet.Cells[j, 1].Value.ToString()))).Day.ToString() +
                                                   ",時間＝" + etsheet.Cells[j, 2].Value.ToString() +
                                                   ",原資料＝" + etsheet.Cells[j, 4].Value.ToString() +
                                                   ",修改為＝");

                                    etsheet.Cells[j, 3].Value = "";
                                    etsheet.Cells[j, 4].Value = "";
                                    etsheet.Cells[j, 5].Value = "";
                                }
                            }
                        }



                        int p = 3;
                        while (true)
                        {
                            if (etsheet.Cells[p, 9].Value != null)
                            {
                                p++;
                            }
                            else
                            {
                                etsheet.Cells[p, 9].Value = DateTime.Now.ToString();
                                break;
                            }
                        }

                        for (int i = 0; i <= etsheet.DataValidations.Count - 1; i++)
                        {
                            if (etsheet.DataValidations[i].Address.Start.Address == "C2")
                            {
                                etsheet.DataValidations.Remove(etsheet.DataValidations[i]);
                                break;
                            }
                        }

                        // add a validation and set values
                        var validation = etsheet.DataValidations.AddListValidation("C2:C" + (FDaysinMonth * 2 + 1).ToString());
                        validation.ShowErrorMessage = true;
                        validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                        validation.ErrorTitle = "無效的值";
                        validation.Error = "請從下拉選單選取值";

                        int CharGridIndex = 0;
                        for (int i = FDaysinMonth * 2 + 2; i < 200; i++)
                        {
                            if (etsheet.Cells[i, 4].Value != null)
                            {
                                if (etsheet.Cells[i, 4].Value.ToString().Trim() == "月總工時")
                                {
                                    CharGridIndex = i;
                                    break;
                                }
                            }
                        }

                        for (int j = CharGridIndex + 1; j <= CharGridIndex + 12; j++)
                        {
                            validation.Formula.Values.Add(etsheet.Cells[j, 1].Text);
                        }
                    } //這邊已修改過EXCEL



                    for (int i = 1; i <= 9; i++)
                    {
                        etsheet.Column(i).AutoFit();
                    }

                    //建立檔案
                    if (chkSaveAs.CheckState == CheckState.Checked)
                    {
                        using (FileStream createStream = new FileStream(tbSaveAs.Text, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                        {
                            ep.SaveAs(createStream);//存檔
                        }
                    }
                    else
                    {
                        using (FileStream createStream = new FileStream(tbTo.Text, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                        {
                            ep.SaveAs(createStream);//存檔
                        }
                    }

                    /*if (chkSaveAs.CheckState == CheckState.Checked)
                    {
                        etbook.SaveAs(tbSaveAs.Text, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, ET.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    else
                    {
                        etbook.Save();
                    }*/

                    /*etbook.Close(ET.XlSaveAction.xlDoNotSaveChanges, Type.Missing, Type.Missing);
                    etApp.Quit();
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(etrange);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(etsheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(etbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(etApp);*/
                    sw.Stop();
                    splashScreenManager1.CloseWaitForm();
                    Msg(DateTime.Now.ToString() + " 匯出資料完成！\r\n" + "耗時：" + sw.Elapsed.Seconds + "秒", "提示");
                }
            }
            catch (System.Exception ex)
            {
                
                /*etbook.Close(ET.XlSaveAction.xlDoNotSaveChanges, Type.Missing, Type.Missing);
                etApp.Quit();
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(etrange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(etsheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(etbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(etApp);*/
                
                sw.Stop();
                Msg(DateTime.Now.ToString() + ex.Message + "\r\n" + "耗時：" + sw.Elapsed.Seconds + "秒", "錯誤");
            }
            finally
            {
                if (splashScreenManager1.IsSplashFormVisible)
                {
                    splashScreenManager1.CloseWaitForm();
                }
                ErrorLog("=====================================匯出資料完成！=====================================");
                GC.Collect();
            }
        }
        public static string GetMergedRangeAddress(ExcelRange xrange)
        {
            if (xrange.Merge)
            {
                var idx = xrange.Worksheet.GetMergeCellId(xrange.Start.Row, xrange.Start.Column);
                return xrange.Worksheet.MergedCells[idx - 1]; //the array is 0-indexed but the mergeId is 1-indexed...
            }
            else
            {
                return xrange.Address;
            }
        }

        private void tbFrom_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            //建立OpenFileDialog
            OpenFileDialog dialog = new OpenFileDialog();

            //設定Filter，過濾檔案 
            dialog.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";

            //設定起始目錄為C:\
            dialog.InitialDirectory = FclassDefaultPath;

            //設定dialog的Title
            dialog.Title = "選擇來源的課表";

            //假如使用者按下OK鈕，則將檔案名稱顯示於TextBox1上
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                FclassDefaultPath = Path.GetDirectoryName(dialog.FileName);
                tbFrom.Text = dialog.FileName;
            }
        }

        private void tbTo_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            //建立OpenFileDialog
            OpenFileDialog dialog = new OpenFileDialog();

            //設定Filter，過濾檔案 "Excel Files|*.xls;*.xlsx;*.xlsm";
            dialog.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";

            //設定起始目錄為C:\
            dialog.InitialDirectory = FreportDefaultPath;

            //設定dialog的Title
            dialog.Title = "選擇要編輯的週報";

            //假如使用者按下OK鈕，則將檔案名稱顯示於TextBox1上
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                FreportDefaultPath = Path.GetDirectoryName(dialog.FileName);
                tbTo.Text = dialog.FileName;
            }
        }

        private void tbSaveAs_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            //建立OpenFileDialog
            SaveFileDialog dialog = new SaveFileDialog();

            //設定Filter，過濾檔案 "Excel Files|*.xls;*.xlsx;*.xlsm";
            dialog.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";

            //設定起始目錄為C:\
            dialog.InitialDirectory = FsaveasDefaultPath;

            //設定dialog的Title
            dialog.Title = "輸入要儲存的檔案名稱";

            //假如使用者按下OK鈕，則將檔案名稱顯示於TextBox1上
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                FsaveasDefaultPath = Path.GetDirectoryName(dialog.FileName);
                tbSaveAs.Text = dialog.FileName;
            }
        }

        private void WeekReportF_Load(object sender, EventArgs e)
        {
            SkinContainerCollection skins = SkinManager.Default.Skins;
            for (int i = 0; i < skins.Count - 1;i++ )
            {
                if (skins[i].SkinName != "Metropolis Dark")
                {
                    repositoryItemComboBox1.Items.Add(skins[i].SkinName);
                }
            }

            if (repositoryItemComboBox1.Items.Count > 0)
            {
                object val = repositoryItemComboBox1.Items[0];
                barEditItem1.EditValue = val;
            }

            DEMonth.DateTime = DateTime.Now;

            update = new SharpUpdater(this);
            barButtonItem3.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

            InitiniFiles();
            LoadConfig();

            bgWorker = new BackgroundWorker();
            bgWorker.DoWork += new DoWorkEventHandler(bgWorker_DoWork);
            bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);

            DoCheckUpdate();

            Text = "WeekReports Ver "+ Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        private void InitiniFiles()
        {
            ErrorLog("初始化INI檔..");
            if (!File.Exists(fc.INIPath))
            {
                try
                {
                    ErrorLog("INI檔案不存在，建立INI檔..");
                    File.Create(fc.INIPath).Close();
                    IConfigSource source = new IniConfigSource(fc.INIPath);
                    source.AddConfig("SYSTEM");
                    IConfig config = source.Configs["SYSTEM"];
                    config.Set("SKINNAME", "Dark Side");
                    config.Set("CLASSPATH", "");
                    config.Set("CLASSDEFAULTPATH", FclassDefaultPath);
                    config.Set("DATE", DateTime.Now.ToOADate());
                    config.Set("DATETYPE", 0);
                    config.Set("REPORTPATH", "");
                    config.Set("REPORTDEFAULTPATH", FreportDefaultPath);
                    config.Set("DATERANGE", "");
                    config.Set("TEACHER", Fteacher);
                    config.Set("SAVEAS", (int)chkSaveAs.CheckState);
                    config.Set("SAVEASPATH", "");
                    config.Set("SAVEASDEFAULTPATH", FsaveasDefaultPath);
                    config.Set("UPDATEPATH", "http://kupoautos.com/DSC/WeekReports/update.xml");
                    
                    source.Save();
                }
                catch (Exception error)
                {
                    ErrorLog("建立INI檔發生錯誤.." + error.Message);
                }
                finally
                {
                    ErrorLog("初始化INI檔..完成");
                }
            }
            else
            {
                try
                {
                    ErrorLog("檢查INI檔..");
                    IConfigSource source = new IniConfigSource(fc.INIPath);
                    IConfig config = source.Configs["SYSTEM"];
                    ErrorLog("檢查SYSTEM標籤..");
                    if (config == null)
                    {
                        source.AddConfig("SYSTEM");
                        config.Set("SKINNAME", "Dark Side");
                        config.Set("CLASSPATH", "");
                        config.Set("CLASSDEFAULTPATH", FclassDefaultPath);
                        config.Set("DATE", DateTime.Now.ToOADate());
                        config.Set("DATETYPE", 0);
                        config.Set("REPORTPATH", "");
                        config.Set("REPORTDEFAULTPATH", FreportDefaultPath);
                        config.Set("DATERANGE", "");
                        config.Set("TEACHER", Fteacher);
                        config.Set("SAVEAS", (int)chkSaveAs.CheckState);
                        config.Set("SAVEASPATH", "");
                        config.Set("SAVEASDEFAULTPATH", FsaveasDefaultPath);
                        config.Set("UPDATEPATH", "http://kupoautos.com/DSC/WeekReports/update.xml");
                    }
                    else
                    {
                        ErrorLog("檢查節點..");
                        string[] mkeys = config.GetKeys();
                        CheckiniNode(config, mkeys, "SKINNAME", "Dark Side");
                        CheckiniNode(config, mkeys, "CLASSPATH", "");
                        CheckiniNode(config, mkeys, "CLASSDEFAULTPATH", FclassDefaultPath);
                        CheckiniNode(config, mkeys, "DATE", DateTime.Now.ToOADate());
                        CheckiniNode(config, mkeys, "DATETYPE", 0);
                        CheckiniNode(config, mkeys, "REPORTPATH", "");
                        CheckiniNode(config, mkeys, "REPORTDEFAULTPATH", FreportDefaultPath);
                        CheckiniNode(config, mkeys, "DATERANGE", "");
                        CheckiniNode(config, mkeys, "TEACHER", Fteacher);
                        CheckiniNode(config, mkeys, "SAVEAS", (int)chkSaveAs.CheckState);
                        CheckiniNode(config, mkeys, "SAVEASPATH", "");
                        CheckiniNode(config, mkeys, "SAVEASDEFAULTPATH", FsaveasDefaultPath);
                        CheckiniNode(config, mkeys, "UPDATEPATH", "http://kupoautos.com/DSC/WeekReports/update.xml");
                    }
                    source.Save();
                }
                catch (Exception error)
                {
                    ErrorLog("檢查INI檔發生錯誤.." + error.Message);
                }
                finally
                {
                    ErrorLog("初始化INI檔..完成");
                }
            }
        }

        private void CheckiniNode(IConfig xconfig ,string[] xkeys,string xkey,object xdefault)
        {
            try
            {
                if (!xkeys.Contains(xkey))
                {
                    ErrorLog("建立節點.." + xkey);
                    xconfig.Set(xkey, xdefault);
                }
            }
            catch (System.Exception ex)
            {
            	
            }
        }

        private void Checkinifiles()
        {
           
            if (!File.Exists(fc.INIPath))
            {
                try
                {
                    ErrorLog("初始化INI檔..");
                    File.Create(fc.INIPath).Close();
                    IConfigSource source = new IniConfigSource(fc.INIPath);
                    source.AddConfig("SYSTEM");
                    IConfig config = source.Configs["SYSTEM"];
                    config.Set("SKINNAME", "Dark Side");
                    config.Set("CLASSPATH", "");
                    config.Set("CLASSDEFAULTPATH", FclassDefaultPath);
                    config.Set("DATE", DateTime.Now.ToOADate());
                    config.Set("DATETYPE", 0);
                    config.Set("REPORTPATH", "");
                    config.Set("REPORTDEFAULTPATH", FreportDefaultPath);
                    config.Set("DATERANGE", "");
                    config.Set("TEACHER", Fteacher);
                    config.Set("SAVEAS", (int)chkSaveAs.CheckState);
                    config.Set("SAVEASPATH", "");
                    config.Set("SAVEASDEFAULTPATH", FsaveasDefaultPath);

                    source.Save();
                }
                catch (Exception error)
                {
                    ErrorLog("初始化INI檔發生錯誤.." + error.Message);
                }
                finally
                {
                    ErrorLog("初始化INI檔..完成");
                }
            }

        }

        private void barEditItem1_EditValueChanged(object sender, EventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SkinName = barEditItem1.EditValue.ToString();
        }

        private void LoadConfig()
        {
            ErrorLog("載入INI檔..");
            try
            {
                IConfigSource source = new IniConfigSource(fc.WRDIR + "\\Config.ini");
                IConfig config = source.Configs["SYSTEM"];
                barEditItem1.EditValue = config.GetString("SKINNAME");
                
                tbFrom.Text = config.GetString("CLASSPATH");
                FclassDefaultPath = config.GetString("CLASSDEFAULTPATH");
                DEMonth.DateTime = DateTime.FromOADate(config.GetDouble("DATE"));
                rgDate.SelectedIndex = config.GetInt("DATETYPE");
                tbTo.Text = config.GetString("REPORTPATH");
                FreportDefaultPath = config.GetString("REPORTDEFAULTPATH");
                tbDateRange.Text = config.GetString("DATERANGE");
                cboTeacher.EditValue = config.GetString("TEACHER");
                chkSaveAs.CheckState = (CheckState)config.GetInt("SAVEAS");
                tbSaveAs.Text = config.GetString("SAVEASPATH");
                FsaveasDefaultPath = config.GetString("SAVEASDEFAULTPATH");
            }
            catch (Exception error)
            {
                ErrorLog("載入INI檔發生錯誤.."+error.Message);
            }
            finally
            {
                ErrorLog("載入INI檔..完成");
            }
        }
        

        private void SaveConfig()
        {
            ErrorLog("儲存INI檔..");
            try
            {
                IConfigSource source = new IniConfigSource(fc.INIPath);
                IConfig config = source.Configs["SYSTEM"];
                config.Set("SKINNAME", barEditItem1.EditValue.ToString());
                config.Set("CLASSPATH", tbFrom.Text);
                try
                {
                    config.Set("CLASSDEFAULTPATH", Path.GetDirectoryName(tbFrom.Text));
                }
                catch (System.Exception ex)
                {
                    config.Set("CLASSDEFAULTPATH", "");
                }
                config.Set("DATE", DEMonth.DateTime.ToOADate());
                config.Set("DATETYPE", rgDate.SelectedIndex);
                config.Set("REPORTPATH", tbTo.Text);
                try
                {
                    config.Set("REPORTDEFAULTPATH", Path.GetDirectoryName(tbTo.Text));
                }
                catch (System.Exception ex)
                {
                    config.Set("REPORTDEFAULTPATH", "");
                }
                
                config.Set("DATERANGE", tbDateRange.Text);
                config.Set("TEACHER", cboTeacher.EditValue==null ? "" : cboTeacher.EditValue);
                config.Set("SAVEAS", (int)chkSaveAs.CheckState);
                config.Set("SAVEASPATH", tbSaveAs.Text);
                try
                {
                    config.Set("SAVEASDEFAULTPATH", Path.GetDirectoryName(tbSaveAs.Text));
                }
                catch (System.Exception ex)
                {
                    config.Set("SAVEASDEFAULTPATH", "");
                }
                source.Save();
            }
            catch (Exception error)
            {
                ErrorLog("儲存INI檔發生錯誤.." + error.Message);
            }
            finally
            {
                ErrorLog("儲存INI檔..完成");
            }
        }

        private void WeekReportF_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
        }

        #region ISharpUpdatable 成員
        public string ApplicationName
        {
            get { return "WeekReports"; }
        }

        public string ApplicationID
        {
            get { return "WeekReports"; }
        }

        public Assembly ApplicationAssembly
        {
          
            get { return Assembly.GetExecutingAssembly(); }
        }

        public Icon ApplicationIcon
        {
            get { return this.Icon; }
        }

        public Uri UpdateXmlLocation
        {
            get
            {
                IConfigSource source = new IniConfigSource(fc.INIPath);
                string mPath = source.Configs["SYSTEM"].Get("UPDATEPATH", "http://kupoautos.com/DSC/WeekReports/update.xml");
                return new Uri(mPath);  
            }
        }

        public Form Context
        {
            get { return this; }
        }

        #endregion

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            update.DoUpdate(false);
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SaveConfig();
        }

        private void ErrorLog(string xmsg)
        {
            fc.ErrorLog(xmsg);
            string mStr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\t\t" + xmsg;
            memoLog.Text += mStr + "\r\n";
                
        }

        private void Msg(string xmsg,string xinfo,bool xshow = true)
        {
            if (xshow)
            {
                fc.Msg(xmsg, xinfo);
            }            
            string mStr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\t\t" + xmsg;
            memoLog.Text += mStr + "\r\n";
        }

        private void btnClearLog_Click(object sender, EventArgs e)
        {
            memoLog.Text = "";
        }

        private void btnFromOpen_Click(object sender, EventArgs e)
        {
            if (File.Exists(tbFrom.Text))
            {
                Process.Start(tbFrom.Text);
            }
        }

        private void btnToOpen_Click(object sender, EventArgs e)
        {
            if (File.Exists(tbTo.Text))
            {
                Process.Start(tbTo.Text);
            }
        }

        private void btnSaveAsOpen_Click(object sender, EventArgs e)
        {
            if (File.Exists(tbSaveAs.Text))
            {
                Process.Start(tbSaveAs.Text);
            }
        }

        private void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            IConfigSource source = new IniConfigSource(fc.INIPath);
            string mPath = source.Configs["SYSTEM"].Get("UPDATEPATH", "http://kupoautos.com/DSC/WeekReports/update.xml");
            var mUri = new Uri(mPath);  

            if (!SharpUpdateXml.ExistOnServer(mUri))
            {
                e.Cancel = true;
            }
            else
            {
                e.Result = SharpUpdateXml.Parse(mUri, Assembly.GetExecutingAssembly().GetName().Name);
            }
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                SharpUpdateXml update = (SharpUpdateXml)e.Result;
                var applicationInfo = Assembly.GetExecutingAssembly();

                if (update != null && update.IsNewerThan(applicationInfo.GetName().Version))
                {
                    barButtonItem3.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                }
                else
                {
                    barButtonItem3.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                }
            }
        }

        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            barButtonItem3.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            timer1.Enabled = true;
            update.DoUpdate(false);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            /*if (barButtonItem3.Visibility == DevExpress.XtraBars.BarItemVisibility.Never)
            {
                DoCheckUpdate();
            }*/
        }

        private void DoCheckUpdate()
        {
            try
            {
                if (!bgWorker.IsBusy)
                    bgWorker.RunWorkerAsync();
            }
            catch (System.Exception ex)
            {
                fc.ErrorLog("檢查更新發生錯誤！  " + ex.Message);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }
    }


}