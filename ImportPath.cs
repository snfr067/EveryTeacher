using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using static EveryTeacher.Program;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using DataTable = System.Data.DataTable;
using Label = System.Windows.Forms.Label;
using Application = System.Windows.Forms.Application;

namespace EveryTeacher
{
    public partial class ImportPath : Form
    {
        static string ORIGIN_FILE_NAME = "原始檔.xlsx";
        static string TEACHER_FILE_NAME = "範例檔導師.xlsx";
        static string DEPARTMENT_FILE_NAME = "範例檔系主任.xlsx";
        static string COLLEGE_FILE_NAME = "範例檔院長.xlsx";
        static string EXPORT_PATH_NAME = "輸出檔案";
        

        string fileName = "";
        string pathName = "";

        string orgFilePath = "";
        string tchFilePath = "";
        string depFilePath = "";
        string colFilePath = "";
        string exportPath = "";

        public ImportPath()
        {
            InitializeComponent();
            this.FormClosed += new FormClosedEventHandler(this.ImportPathClosed);
        }

        private void ImportPath_Load(object sender, EventArgs e)
        {
            initUI();
            importOrgPath_txtbx.Text =
                System.Windows.Forms.Application.StartupPath + "\\" + ORIGIN_FILE_NAME;
            importTchPath_txtbx.Text =
                System.Windows.Forms.Application.StartupPath + "\\" + TEACHER_FILE_NAME;
            importDepPath_txtbx.Text =
                System.Windows.Forms.Application.StartupPath + "\\" + DEPARTMENT_FILE_NAME;
            importColPath_txtbx.Text =
                System.Windows.Forms.Application.StartupPath + "\\" + COLLEGE_FILE_NAME;
            exportPath_txtbx.Text =
                System.Windows.Forms.Application.StartupPath + "\\" + EXPORT_PATH_NAME + "\\";

        }

        private void initUI()
        {
            ckOrg_txt.Text = "確認中...";
            ckTch_txt.Text = "確認中...";
            ckDep_txt.Text = "確認中...";
            ckCol_txt.Text = "確認中...";

            ckOrg_txt.ForeColor = System.Drawing.Color.Orange;
            ckTch_txt.ForeColor = System.Drawing.Color.Orange;
            ckDep_txt.ForeColor = System.Drawing.Color.Orange;
            ckCol_txt.ForeColor = System.Drawing.Color.Orange;

            ckOrg_txt.Visible = false;
            ckTch_txt.Visible = false;
            ckDep_txt.Visible = false;
            ckCol_txt.Visible = false;
        }

        private void ImportPathClosed(object sender, FormClosedEventArgs e)
        {
            //Program.KillExcelApp();

            Application.Exit(); //這樣程式才會完全關閉並釋放資源
        }

        private void importOrgPath_btn_Click(object sender, EventArgs e)
        {
            fileName = ImportExcelFile();

            if(!fileName.Equals(""))
                importOrgPath_txtbx.Text = fileName;            
        }
        
        private void importTchPath_btn_Click(object sender, EventArgs e)
        {
            fileName = ImportExcelFile();

            if (!fileName.Equals(""))
                importTchPath_txtbx.Text = fileName;
        }

        private void importDepPath_btn_Click(object sender, EventArgs e)
        {
            fileName = ImportExcelFile();

            if (!fileName.Equals(""))
                importDepPath_txtbx.Text = fileName;
        }
        private void importColPath_btn_Click(object sender, EventArgs e)
        {
            fileName = ImportExcelFile();

            if (!fileName.Equals(""))
                importColPath_txtbx.Text = fileName;
        }
        private void exportPath_btn_Click(object sender, EventArgs e)
        {
            pathName = getSelectedFolderPath();

            if (!pathName.Equals(""))
                exportPath_txtbx.Text = pathName;
        }

        private void next_btn_Click(object sender, EventArgs e)
        {
            //匯入檔案路徑
            //匯出檔案路徑
            //開始讀取
            //開始複製, 寫入
            //產生寄信設定
            string errorResult = "";

            next_btn.Text = "小等一下...";
            next_btn.Enabled = false;

            errorResult = checkAnyError();
            if (!errorResult.Equals(""))
            {
                MessageBox.Show(errorResult);
                initUI();
            }
            else
            {
                orgFilePath = importOrgPath_txtbx.Text;
                tchFilePath = importTchPath_txtbx.Text;
                depFilePath = importDepPath_txtbx.Text;
                colFilePath = importColPath_txtbx.Text;
                exportPath = exportPath_txtbx.Text;

                SplitExcel split = new SplitExcel(orgFilePath, tchFilePath, 
                    depFilePath, colFilePath, exportPath);
                split.Show();
                split.Visible = true;

                this.Visible = false;
            }
            
        }

        public void setCheckString(Label text, bool isRight)
        {
            if(isRight)
            {
                text.Text = "檔案正確";
                text.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                text.Text = "檔案錯誤";
                text.ForeColor = System.Drawing.Color.Red;
            }
        }

        public string checkAnyError()
        {
            string result = "";
            string formatRet = "";
            string nextLine = "\n";

            string[] checkOrgHeaders = { HEADER_DEPERTMENT, HEADER_COLLEGE, HEADER_TEACHERS,
                HEADER_CLASS, HEADER_STUDENT_NUM, HEADER_STUDENT_NAME,
                HEADER_STUDENT_PHONE, HEADER_RELIEF, HEADER_TCH_EMAIL };
            int orgIndex = 1;

            string[] checkTchHeaders = { HEADER_CLASS, HEADER_STUDENT_NUM, HEADER_STUDENT_NAME,
                HEADER_STUDENT_PHONE, HEADER_RELIEF };
            int tchIndex = 4;

            string[] checkDepHeaders = { HEADER_TEACHERS,
                HEADER_CLASS, HEADER_STUDENT_NUM, HEADER_STUDENT_NAME,
                HEADER_STUDENT_PHONE, HEADER_RELIEF };
            int depIndex = 4;

            string[] checkColHeaders = { HEADER_DEPERTMENT,
                HEADER_CLASS, HEADER_STUDENT_NUM, HEADER_STUDENT_NAME,
                HEADER_STUDENT_PHONE, HEADER_RELIEF };
            int colIndex = 4;



            //路徑
            try
            {
                ckOrg_txt.Visible = true;
                if (!File.Exists(importOrgPath_txtbx.Text))
                {
                    result += ORIGIN_FILE_NAME + "不存在!" + nextLine;
                    setCheckString(ckOrg_txt, false);
                }
                //格式
                else
                {
                    formatRet = isExcelFormat(importOrgPath_txtbx.Text,
                        checkOrgHeaders, orgIndex);
                    if (!formatRet.Equals(""))
                    {
                        result += ORIGIN_FILE_NAME + "格式錯誤: " + formatRet + nextLine;
                        setCheckString(ckOrg_txt, false);
                    }
                    else
                        setCheckString(ckOrg_txt, true);
                }

                ckTch_txt.Visible = true;
                if (!File.Exists(importTchPath_txtbx.Text))
                {
                    result += TEACHER_FILE_NAME + "不存在!" + nextLine;
                    setCheckString(ckTch_txt, false);
                }
                //格式
                else
                {
                    formatRet = isExcelFormat(importTchPath_txtbx.Text,
                        checkTchHeaders, tchIndex);
                    if (!formatRet.Equals(""))
                    {
                        result += TEACHER_FILE_NAME + "格式錯誤: " + formatRet + nextLine;

                        setCheckString(ckTch_txt, false);
                    }
                    else
                        setCheckString(ckTch_txt, true);
                }

                ckDep_txt.Visible = true;
                if (!File.Exists(importDepPath_txtbx.Text))
                {
                    result += DEPARTMENT_FILE_NAME + "不存在!" + nextLine;
                    setCheckString(ckDep_txt, false);
                }
                //格式
                else
                {
                    formatRet = isExcelFormat(importDepPath_txtbx.Text,
                        checkDepHeaders, depIndex);
                    if (!formatRet.Equals(""))
                    {
                        result += DEPARTMENT_FILE_NAME + "格式錯誤: " + formatRet + nextLine;

                        setCheckString(ckDep_txt, false);

                    }
                    else
                        setCheckString(ckDep_txt, true);
                }

                ckCol_txt.Visible = true;
                if (!File.Exists(importColPath_txtbx.Text))
                {
                    result += COLLEGE_FILE_NAME + "不存在!" + nextLine;
                    setCheckString(ckCol_txt, false);
                }
                //格式
                else
                {
                    formatRet = isExcelFormat(importColPath_txtbx.Text,
                        checkColHeaders, colIndex);
                    if (!formatRet.Equals(""))
                    { 
                        result += COLLEGE_FILE_NAME + "格式錯誤: " + formatRet + nextLine;

                        setCheckString(ckCol_txt, false);
                    }
                    else
                        setCheckString(ckCol_txt, true);
                }

                if (result.Equals("") && !Directory.Exists(exportPath_txtbx.Text))
                {
                    Directory.CreateDirectory(exportPath_txtbx.Text);
                }
            }
            catch (Exception fileEx)
            {
                result += "檢查路徑錯誤: "+ fileEx.Message + nextLine;
            }
            
            return result;
        }

        public bool checkPathExist(string path)
        {
            return Directory.Exists(path);
        }


        public String ImportExcelFile()
        {
            string windowFilter = "Excel files|*.xlsx";
            string windowTitle = "匯入Excel資料";

            OpenFileDialog openFileDialogFunction = new OpenFileDialog();
            openFileDialogFunction.Filter = windowFilter; //開窗搜尋副檔名
            openFileDialogFunction.Title = windowTitle; //開窗標題

            DataTable dataTable = new DataTable();

            if (openFileDialogFunction.ShowDialog() == DialogResult.OK)
            {
                string FilePath = openFileDialogFunction.FileName;
                return FilePath;
            }

            return "";
        }

        public string getSelectedFolderPath()
        {
            FolderBrowserDialog openFolderDialog = new FolderBrowserDialog();

            if (openFolderDialog.ShowDialog() == DialogResult.OK)
            {
                return openFolderDialog.SelectedPath;
            }
            else
                return "";
        }

        /*public static bool DataTableToExcel(DataSet dataSet, string Outpath)
        {
            bool result = false;
            try
            {
                if (dataSet == null || dataSet.Tables == null || dataSet.Tables.Count == 0 || string.IsNullOrEmpty(Outpath))
                    throw new Exception("輸入的DataSet或路徑例外");
                int sheetIndex = 0;
                //根據輸出路徑的擴展名判斷workbook的實體型別
                IWorkbook workbook = null;
                string pathExtensionName = Outpath.Trim().Substring(Outpath.Length - 5);
                if (pathExtensionName.Contains(".xlsx"))
                {
                    workbook = new XSSFWorkbook();
                }
                else if (pathExtensionName.Contains(".xls"))
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    Outpath = Outpath.Trim() + ".xls";
                    workbook = new HSSFWorkbook();
                }
                //將DataSet匯出為Excel
                foreach (DataTable dt in dataSet.Tables)
                {
                    sheetIndex++;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        ISheet sheet = workbook.CreateSheet(string.IsNullOrEmpty(dt.TableName) ? ("sheet" + sheetIndex) : dt.TableName);//創建一個名稱為Sheet0的表
                        int rowCount = dt.Rows.Count;//行數
                        int columnCount = dt.Columns.Count;//列數

                        //設定列頭
                        IRow row = sheet.CreateRow(0);//excel第一行設為列頭
                        for (int c = 0; c < columnCount; c++)
                        {
                            ICell cell = row.CreateCell(c);
                            cell.SetCellValue(dt.Columns[c].ColumnName);
                        }

                        //設定每行每列的單元格,
                        for (int i = 0; i < rowCount; i++)
                        {
                            row = sheet.CreateRow(i + 1);
                            for (int j = 0; j < columnCount; j++)
                            {
                                ICell cell = row.CreateCell(j);//excel第二行開始寫入資料
                                cell.SetCellValue(dt.Rows[i][j].ToString());
                            }
                        }
                    }
                }
                //向outPath輸出資料
                using (FileStream fs = File.OpenWrite(Outpath))
                {
                    workbook.Write(fs);//向打開的這個xls檔案中寫入資料
                    result = true;
                }
                return result;
            }
            catch (Exception ex)
            {
                return false;
            }
        }*/

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
                else if(!val.Equals(""))
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
                    return "檔案標頭未包含"+check;
            }

            return "";
        }

        static void readExcel(string path, int indexOfSheet)
        {
            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(path, 0, true, 5, "", "", true,
                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                  , "\t", false, false, 0, true, 1, 0);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(path);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[indexOfSheet];

            Excel.Range row = Wsheet.Rows[1];
            Excel.Range column = Wsheet.Columns[5];

            Excel.Range cell = (Excel.Range)Wsheet.Cells[3, 4];
            Excel.Range rng = (Excel.Range)Wsheet.Cells[3, 4];
            string val = rng.Value.ToString();

            

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
            KillExcelApp(App);
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
            KillExcelApp(App);
        }

    }
    
}
