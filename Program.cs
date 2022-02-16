using System;
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
using DataTable = System.Data.DataTable;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Label = System.Windows.Forms.Label;
using System.Text;

namespace EveryTeacher
{
    internal static class Program
    {
        public static string HEADER_DEPERTMENT = "系所";
        public static string HEADER_COLLEGE = "學院";
        public static string HEADER_TEACHERS = "導師姓名";
        public static string HEADER_CLASS = "班級";
        public static string HEADER_STUDENT_NUM = "學號";
        public static string HEADER_STUDENT_NAME = "姓名";
        public static string HEADER_STUDENT_PHONE = "學生手機";
        public static string HEADER_RELIEF = "減免類別";
        public static string HEADER_TCH_EMAIL = "導師Email"; 


        public static int INDEX_TCH_CLASS = 1;
        public static int INDEX_TCH_STUDENT_NUM = 2;
        public static int INDEX_TCH_STUDENT_NAME = 3;
        public static int INDEX_TCH_STUDENT_PHONE = 4;
        public static int INDEX_TCH_RELIEF = 5;

        public static int INDEX_DEP_TCH_NAME = 1;
        public static int INDEX_DEP_CLASS = 2;
        public static int INDEX_DEP_STUDENT_NUM = 3;
        public static int INDEX_DEP_STUDENT_NAME = 4;
        public static int INDEX_DEP_STUDENT_PHONE = 5;
        public static int INDEX_DEP_RELIEF = 6;

        public static int INDEX_COL_DEP = 1;
        public static int INDEX_COL_CLASS = 2;
        public static int INDEX_COL_STUDENT_NUM = 3;
        public static int INDEX_COL_STUDENT_NAME = 4;
        public static int INDEX_COL_STUDENT_PHONE = 5;
        public static int INDEX_COL_RELIEF = 6;

        public static string DIR_NAME_TEACHERS = "寄給導師的\\";
        public static string DIR_NAME_DEPARTMENT = "寄給系主任的\\";
        public static string DIR_NAME_COLLEGE = "寄給院長的\\";

        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ImportPath());
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

            foreach (string check in checkHeaders)
            {
                if (!val.Contains(check))
                    return "檔案標頭未包含" + check;
            }

            return "";
        }

        public static void writeTchExcel(string importPath, string exportPath, string tchFile, 
            ProgressBar bar, Label prgText)
        {
            string val = "";
            string dstFile = "";

            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(tchFile, 0, true, 5, "", "", true,
                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                  , "\t", false, false, 0, true, 1, 0);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(tchFile);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[1];
            Excel.Range rng;// = Wsheet.Rows[5];
            Excel.Range cellClass;
            Excel.Range cellStNum;
            Excel.Range cellStName;
            Excel.Range cellStPhone;
            Excel.Range cellRelief;
            int tchWriteIndex = 5;
            int tchDataRowIndex = tchWriteIndex;
            string teacherName = "";
            

            /*for(int i = 5; i < 300; i++)
            {
                for(int j = 5; j < 300; j++)
                {
                    Wsheet.Cells[i, j].Value = "";
                }
            }*/

            DataSet ds = ExcelToDataSet(importPath, true);
            if (ds != null)
            {
                DataTable dt = ds.Tables[0];
                DataRowCollection readRows = dt.Rows.Copy();

                bar.Minimum = 0;
                bar.Maximum = dt.Rows.Count;

                prgText.Text = bar.Value + "/"+ bar.Maximum;

                while (dt.Rows.Count > 0)
                {
                    teacherName = dt.Rows[0][HEADER_TEACHERS].ToString();
                    foreach (DataRow dataRow in readRows) 
                    {
                        bool isNull = (dataRow == null);
                        if (dataRow != null)
                        {
                            if (teacherName.Equals(dataRow[HEADER_TEACHERS].ToString()))        //過濾老師
                            {
                                val += dataRow[HEADER_CLASS].ToString();
                                rng = Wsheet.Rows[tchDataRowIndex];

                                cellClass = Wsheet.Cells[tchDataRowIndex, INDEX_TCH_CLASS];
                                cellStNum = Wsheet.Cells[tchDataRowIndex, INDEX_TCH_STUDENT_NUM];
                                cellStName = Wsheet.Cells[tchDataRowIndex, INDEX_TCH_STUDENT_NAME];
                                cellStPhone = Wsheet.Cells[tchDataRowIndex, INDEX_TCH_STUDENT_PHONE];
                                cellRelief = Wsheet.Cells[tchDataRowIndex, INDEX_TCH_RELIEF];

                                cellClass.Value2 = dataRow[HEADER_CLASS].ToString();
                                cellStNum.Value2 = dataRow[HEADER_STUDENT_NUM].ToString();
                                cellStName.Value2 = dataRow[HEADER_STUDENT_NAME].ToString();
                                cellStPhone.NumberFormat = "@";
                                if (dataRow[HEADER_STUDENT_PHONE].ToString().Length == 9
                                    && dataRow[HEADER_STUDENT_PHONE].ToString().StartsWith("9"))
                                    cellStPhone.Value = "0" + dataRow[HEADER_STUDENT_PHONE].ToString();
                                else
                                    cellStPhone.Value = dataRow[HEADER_STUDENT_PHONE].ToString();

                                cellRelief.Value2 = dataRow[HEADER_RELIEF].ToString();


                                System.Diagnostics.Debug.WriteLine("write:"+dataRow[HEADER_STUDENT_NAME].ToString());

                                if (dstFile.Equals(""))
                                {
                                    dstFile = exportPath + dataRow[HEADER_CLASS].ToString() + ".xlsx";

                                    if (!File.Exists(dstFile))
                                        File.Copy(tchFile, dstFile);

                                }

                                dt.Rows.RemoveAt(0);    //由於刪掉，所以下一個要被刪的，會變第0個

                                tchDataRowIndex++;
                            }
                        }
                        else
                            break;
                    }

                    //設置禁止彈出保存和覆蓋的詢問提示框
                    Wsheet.Application.DisplayAlerts = false;
                    Wsheet.Application.AlertBeforeOverwriting = false;

                    //保存工作表，因為禁止彈出儲存提示框，所以需在此儲存，否則寫入的資料會無法儲存
                    Wbook.SaveCopyAs(dstFile);

                    tchDataRowIndex = tchWriteIndex;
                    dstFile = "";

                    Wbook = App.Workbooks.Open(tchFile, 0, true, 5, "", "", true,
                     Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                      , "\t", false, false, 0, true, 1, 0);                    
                    xlsAttribute.Attributes = FileAttributes.Normal;
                    Wsheet = (Excel.Worksheet)Wbook.Sheets[1];

                    if (readRows.Count == dt.Rows.Count)
                    {
                        MessageBox.Show("原始資料有異常!");
                        break;
                    }
                    else
                    {
                        bar.Step = readRows.Count - dt.Rows.Count;
                        bar.PerformStep();

                        prgText.Text = bar.Value + "/" + bar.Maximum;

                        readRows = dt.Rows.Copy();
                    }


                    System.Diagnostics.Debug.WriteLine("------------");
                }
            }
            
            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
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
            Excel.Range aRangeChange = Wsheet.get_Range("B7");

            //在工作表的特定儲存格，設定內容
            aRangeChange.Value2 = "施argaza";

            //設置禁止彈出保存和覆蓋的詢問提示框
            Wsheet.Application.DisplayAlerts = false;
            Wsheet.Application.AlertBeforeOverwriting = false;

            //保存工作表，因為禁止彈出儲存提示框，所以需在此儲存，否則寫入的資料會無法儲存
            Wbook.Save();

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
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

        public static void genSendMailFile(SendMail[] smobjs, string file)
        {
            string csvData =
                nameof(SendMail.SendName) + "," +
                nameof(SendMail.Sendto) + "," +
                nameof(SendMail.CC) + "," +
                nameof(SendMail.Attach) + "," +
                nameof(SendMail.Title) + "," +
                nameof(SendMail.Subject) + "\n";

            foreach (SendMail sendMail in smobjs)
            {
                if (sendMail != null)
                {
                    csvData += sendMail.SendName + ",";
                    csvData += sendMail.Sendto + ",";
                    csvData += sendMail.CC + ",";
                    csvData += sendMail.Attach + ",";
                    csvData += sendMail.Title + ",";
                    csvData += sendMail.Subject + "\n";
                }
                else
                    break;
            }

            try
            {
                StreamWriter sw = new StreamWriter(file, false, Encoding.UTF8);
                sw.Write(csvData);
                sw.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("CSV檔案錯誤: "+ex.Message);
            }
        }

    }
}
