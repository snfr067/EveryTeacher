﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;
using Application = System.Windows.Forms.Application;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using DataTable = System.Data.DataTable;
using Label = System.Windows.Forms.Label;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace EveryTeacher
{
    internal static class Program
    {
        public static string APP_VERSION = "Free.22.03.17.01";
        
        public static int EXAMPLE_HEADER_ROW = 4;
        public static int ORIGIN_HEADER_ROW = 1;

        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //以下是整合DLL檔案的設定
            string rsNPOI = "EveryTeacher.NPOI.dll";
            string rsOOXML = "EveryTeacher.NPOI.OOXML.dll";
            string rs4Net = "EveryTeacher.NPOI.OpenXml4Net.dll";
            string rsFormats = "EveryTeacher.NPOI.OpenXmlFormats.dll";
            string rsICSharp = "EveryTeacher.ICSharpCode.SharpZipLib.dll";
            
            EmbeddedAssembly.Load(rsNPOI, "NPOI.dll");
            EmbeddedAssembly.Load(rsOOXML, "NPOI.OOXML.dll");
            EmbeddedAssembly.Load(rs4Net, "NPOI.OpenXml4Net.dll");
            EmbeddedAssembly.Load(rsFormats, "NPOI.OpenXmlFormats.dll");
            EmbeddedAssembly.Load(rsICSharp, "ICSharpCode.SharpZipLib.dll");

            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
            //

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ImportPath());
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name);
        }


        public static void KillExcelApp()
        {
            try
            {
                System.Diagnostics.Process[] procs =
                    System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process p in procs)
                {
                    p.Kill();
                }
            }
            catch
            {

            }
        }

        static string isExcelFormat(string path, string[] checkHeaders, int rowIndex)
        {
            string val = "";
            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(path, 0, true, 5, "", "", true,
                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                  , "\t", false, false, 0, true, 1, 0);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(path);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[1];
            Excel.Range row = Wsheet.Rows[rowIndex];

            foreach (Excel.Range r in row.Cells) //range1.Cells represents all the columns/rows
            {
                bool isNull = (r == null) || (r.Value == null);
                if (!isNull)
                {
                    val += r.Value.ToString();
                }
                else if (!val.Equals(""))
                    break;
            }

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
            KillExcelApp(App);

            foreach (string check in checkHeaders)
            {
                if (!val.Contains(check))
                    return "檔案標頭未包含" + check;
            }

            return "";
        }



        static void writeToExcel(string path, int indexOfSheet)
        {
            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(path);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(path);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[indexOfSheet];
            //Wbook.Worksheets.Add()

            //取得工作表的單元格
            //列(左至右)ABCDE, 行(上至下)12345
            Excel.Range aRangeChange = Wsheet.get_Range("A1");

            //在工作表的特定儲存格，設定內容
            aRangeChange.Value2 = "測試";

            //設置禁止彈出保存和覆蓋的詢問提示框
            Wsheet.Application.DisplayAlerts = false;
            Wsheet.Application.AlertBeforeOverwriting = false;

            //保存工作表，因為禁止彈出儲存提示框，所以需在此儲存，否則寫入的資料會無法儲存
            Wbook.Save();

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
            KillExcelApp(App);
        }

        public static DataSet ExcelToDataSet(string filePath, bool isFirstLineColumnName)
        {
            DataSet dataSet = new DataSet();
            int startRow = 0;
            try
            {
                using (FileStream fs = File.OpenRead(filePath))
                {
                    IWorkbook workbook = null;
                    // 如果是2007+的Excel版本
                    if (filePath.IndexOf(".xlsx") > 0)
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    // 如果是2003-的Excel版本
                    else if (filePath.IndexOf(".xls") > 0)
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    if (workbook != null)
                    {
                        //回圈讀取Excel的每個sheet，每個sheet頁都轉換為一個DataTable，并放在DataSet中
                        for (int p = 0; p < workbook.NumberOfSheets; p++)
                        {
                            ISheet sheet = workbook.GetSheetAt(p);
                            DataTable dataTable = new DataTable();
                            dataTable.TableName = sheet.SheetName;
                            if (sheet != null)
                            {
                                int rowCount = sheet.LastRowNum;//獲取總行數
                                if (rowCount > 0)
                                {
                                    IRow firstRow = sheet.GetRow(0);//獲取第一行
                                    int cellCount = firstRow.LastCellNum;//獲取總列數

                                    //構建datatable的列
                                    if (isFirstLineColumnName)
                                    {
                                        startRow = 1;//如果第一行是列名，則從第二行開始讀取
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                        {
                                            ICell cell = firstRow.GetCell(i);
                                            if (cell != null)
                                            {
                                                if (cell.StringCellValue != null)
                                                {
                                                    DataColumn column = new DataColumn(cell.StringCellValue);
                                                    dataTable.Columns.Add(column);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                        {
                                            DataColumn column = new DataColumn("column" + (i + 1));
                                            dataTable.Columns.Add(column);
                                        }
                                    }

                                    //填充行
                                    for (int i = startRow; i <= rowCount; ++i)
                                    {
                                        IRow row = sheet.GetRow(i);
                                        if (row == null) continue;

                                        DataRow dataRow = dataTable.NewRow();
                                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                                        {
                                            ICell cell = row.GetCell(j);
                                            if (cell == null)
                                            {
                                                dataRow[j] = "";
                                            }
                                            else
                                            {
                                                //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        dataRow[j] = "";
                                                        break;
                                                    case CellType.Numeric:
                                                        short format = cell.CellStyle.DataFormat;
                                                        //對時間格式（2015.12.5、2015/12/5、2015-12-5等）的處理
                                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                                            dataRow[j] = cell.DateCellValue;
                                                        else
                                                            dataRow[j] = cell.NumericCellValue;
                                                        break;
                                                    case CellType.String:
                                                        dataRow[j] = cell.StringCellValue;
                                                        break;
                                                }
                                            }
                                        }
                                        dataTable.Rows.Add(dataRow);
                                    }
                                }
                            }
                            dataSet.Tables.Add(dataTable);
                        }

                    }
                }
                return dataSet;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }


        public static bool DataTableToExcel(DataTable dt, string path)
        {
            if (dt.Rows.Count == 0)
                return false;
            //建立Excel物件    
            Excel.Application excel = new Excel.Application();
            Excel.Workbook Wbook = excel.Application.Workbooks.Add(true);
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[1];
            DataRow row;

            //生成欄位名稱    
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                Wsheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
            }
            //填充資料    
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                row = dt.Rows[i];
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    Wsheet.Cells[i + 2, j + 1] = "'" + row[j].ToString();
                    Wsheet.Columns.ColumnWidth = 20;
                }
            }

            //設置禁止彈出保存和覆蓋的詢問提示框
            excel.Application.DisplayAlerts = false;
            excel.Application.AlertBeforeOverwriting = false;

            //保存工作表，因為禁止彈出儲存提示框，所以需在此儲存，否則寫入的資料會無法儲存
            Wbook.SaveAs(path);

            Wbook.Close();

            //離開應用程式
            excel.Quit();
            Program.KillExcelApp(excel);

            return true;
        }

        public static void genSendMailFile(SendMail[] smobjs, string file)
        {
            DataTable sendTable = new DataTable();
            DataRow row;
            DataColumn column;
            string[] headers = SendMail.getAllHeadersText();
            string[] smobjsText;

            foreach (string header in headers)
            {
                column = new DataColumn(header);
                sendTable.Columns.Add(column);
            }

            foreach (SendMail sendMail in smobjs)
            {
                if (sendMail != null)
                {
                    row = sendTable.NewRow();
                    smobjsText = sendMail.getAllItemsText();
                    for (int i = 0; i < smobjsText.Length; i++)
                    {
                        row[i] = smobjsText[i];
                    }
                    sendTable.Rows.Add(row);
                }
                else
                    break;
            }
            

            try
            {
                /*StreamWriter sw = new StreamWriter(file, false, Encoding.UTF8);     //亂碼的話用UTF8
                sw.Write(csvData);
                sw.Close();*/
                DataTableToExcel(sendTable, file);
            }
            catch (Exception ex)
            {
                MessageBox.Show("CSV檔案錯誤: " + ex.Message);
            }
        }

        // release excel resource
        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId
        (IntPtr hWnd, out int ProcessId);
        public static void KillExcelApp(Excel.Application app)
        {
            if (app != null)
            {
                try
                {
                    app.Quit();
                    IntPtr intptr = new IntPtr(app.Hwnd);
                    var ps = Process.GetProcessesByName("EXCEL").ToList();
                    int id;
                    GetWindowThreadProcessId(intptr, out id);
                    var p = Process.GetProcessById(id);
                    //if (p != null)
                    p.Kill();
                }
                catch (Exception)
                {
                }
            }
        }

        public static string readStrExcelCellinRow(string path, int rowIndex)
        {
            string val = "";
            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(path, 0, true, 5, "", "", true,
                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                  , "\t", false, false, 0, true, 1, 0);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(path);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[1];
            Excel.Range row = Wsheet.Rows[rowIndex];

            foreach (Excel.Range r in row.Cells) //range1.Cells represents all the columns/rows
            {
                bool isNull = (r == null) || (r.Value == null);
                if (!isNull)
                {
                    val += r.Value.ToString();
                }
                else if (!val.Equals(""))
                    break;
            }

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
            KillExcelApp(App);

            return val;
        }
        public static string[] readStrArrExcelCellinRow(string path, int rowIndex)
        {
            string[] vals;
            int cellCount = 0, i = 0;

            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(path, 0, true, 5, "", "", true,
                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                  , "\t", false, false, 0, true, 1, 0);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(path);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[1];
            Excel.Range row = Wsheet.Rows[rowIndex];


            foreach (Excel.Range r in row.Cells) //range1.Cells represents all the columns/rows
            {
                if (!((r == null) || (r.Value == null)))
                {
                    cellCount++;
                }
                else
                    break;
            }

            vals = new string[cellCount];


            foreach (Excel.Range r in row.Cells) //range1.Cells represents all the columns/rows
            {
                bool isNull = (r == null) || (r.Value == null);
                if (!isNull)
                {
                    vals[i] = r.Value.ToString();
                    i++;
                }
                else
                    break;
            }

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
            KillExcelApp(App);

            return vals;
        }
        public static void genLogFile(string orgPath, string splitHeader, string examplePath, string exportPath, 
            bool isSendMail, string sendName, string sendTo, int countExFiles)
        {
            string log = "";

            log += "Export " + countExFiles + " Files" + "\n";
            log += "orgPath: " + orgPath + "\n";
            log += "splitHeader: " + splitHeader + "\n";
            log += "examplePath: " + examplePath + "\n";
            log += "exportPath: " + exportPath + "\n";
            log += "isSendMail: " + isSendMail + "\n";
            log += "sendName: " + sendName + "\n";
            log += "sendTo: " + sendTo + "\n";
            log += "Date: " + DateTime.Now.ToString() + "\n";
            log += "Version: " + APP_VERSION + "\n"; 

            try
            {
                StreamWriter sw = new StreamWriter(exportPath + "\\" + 
                    "輸出結果("+ DateTime.Now.ToString("yyyyMMddHHmmss")+").txt", false, Encoding.UTF8);     
                sw.Write(log);
                sw.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Log檔案錯誤: " + ex.Message);
            }
        }
    }

    
}
