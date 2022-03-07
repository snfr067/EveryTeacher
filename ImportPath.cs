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

        string[] headers;
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

            version_txt.Text = APP_VERSION;
        }

        private void ImportPath_Load(object sender, EventArgs e)
        {
            initUI();
            importOrgPath_txtbx.Text =
                System.Windows.Forms.Application.StartupPath + "\\" + ORIGIN_FILE_NAME;
            importTchPath_txtbx.Text =
                System.Windows.Forms.Application.StartupPath + "\\" + TEACHER_FILE_NAME;
            exportPath_txtbx.Text =
                System.Windows.Forms.Application.StartupPath + "\\" + EXPORT_PATH_NAME + "\\";

            //reloadHeaders();
        }

        private void initUI()
        {
            ckOrg_txt.Text = "確認中...";
            ckTch_txt.Text = "確認中...";

            ckOrg_txt.ForeColor = System.Drawing.Color.Orange;
            ckTch_txt.ForeColor = System.Drawing.Color.Orange;

            ckOrg_txt.Visible = false;
            ckTch_txt.Visible = false;

            next_btn.Text = "下一頁";
            next_btn.Enabled = true;

            header_combox.Enabled = true;

            if(need_mail_cbx.Checked)
            {
                sendName_combox.Enabled = true;
                sendTo_combox.Enabled = true;
            }
        }
        

        private void reloadsendToData()
        {
            sendName_combox.Items.Clear();
            sendName_combox.Text = "讀取中...";
            sendName_combox.Enabled = false;

            sendTo_combox.Items.Clear();
            sendTo_combox.Text = "讀取中...";
            sendTo_combox.Enabled = false;

            if (headers.Length != 0)
            {
                sendName_combox.Items.Add("");
                sendTo_combox.Items.Add("");
                foreach (string header in headers)
                {
                    sendName_combox.Items.Add(header);
                    sendTo_combox.Items.Add(header);
                }
                sendName_combox.SelectedIndex = 0;
                sendTo_combox.SelectedIndex = 0;
            }
            else
            {
                sendName_combox.Text = "";
                sendTo_combox.Text = "";
            }
            
            sendName_combox.Enabled = true;
            sendTo_combox.Enabled = true;
        }


        private void ImportPathClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit(); //這樣程式才會完全關閉並釋放資源
        }

        private void importOrgPath_btn_Click(object sender, EventArgs e)
        {
            fileName = ImportExcelFile();

            if (!fileName.Equals(""))
            {
                importOrgPath_txtbx.Text = fileName;
            }
        }
        
        private void importTchPath_btn_Click(object sender, EventArgs e)
        {
            fileName = ImportExcelFile();

            if (!fileName.Equals(""))
                importTchPath_txtbx.Text = fileName;
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

            header_combox.Enabled = false;
            sendName_combox.Enabled = false;
            sendTo_combox.Enabled = false;

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
                exportPath = exportPath_txtbx.Text;

                SplitExcel split;

                if(need_mail_cbx.Checked)
                {
                    split = new SplitExcel(orgFilePath, header_combox.Text, tchFilePath, exportPath, 
                        sendName_combox.Text, sendTo_combox.Text);
                }
                else
                {
                    split = new SplitExcel(orgFilePath, header_combox.Text, tchFilePath, exportPath);
                }
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
            string nextLine = "\n";
            string orgHeadersStr;
            string[] tchHeaders;
            string[] mailData = { MAIL_HEADER_TEACHERS, MAIL_HEADER_TCH_EMAIL }; 

            int orgIndex = 1;            
            int tchIndex = 4;


            //路徑
            try
            {
                ckOrg_txt.Visible = true;
                if (!File.Exists(importOrgPath_txtbx.Text))
                {
                    result += ORIGIN_FILE_NAME + "不存在!" + nextLine;
                    setCheckString(ckOrg_txt, false);
                }

                ckTch_txt.Visible = true;
                if (!File.Exists(importTchPath_txtbx.Text))
                {
                    result += TEACHER_FILE_NAME + "不存在!" + nextLine;
                    setCheckString(ckTch_txt, false);
                }
                else
                    setCheckString(ckTch_txt, true);
                


                //格式
                if (result.Equals(""))       //路徑檢查無誤
                {
                    orgHeadersStr = readStrExcelCellinRow(importOrgPath_txtbx.Text, orgIndex);
                    tchHeaders = readStrArrExcelCellinRow(importTchPath_txtbx.Text, tchIndex);

                    foreach(string tchHeader in tchHeaders)
                    {
                        if(!orgHeadersStr.Contains(tchHeader))      //未包含，格式錯誤
                        {
                            result += ORIGIN_FILE_NAME + "格式錯誤: 檔案標頭未包含[" + tchHeader + "]" + nextLine;
                            setCheckString(ckOrg_txt, false);
                        }
                    }

                    if (!orgHeadersStr.Contains(header_combox.SelectedItem.ToString()))      //未包含，格式錯誤
                    {
                        result += ORIGIN_FILE_NAME + "格式錯誤: 檔案標頭未包含[" + header_combox.SelectedItem.ToString() + "]" + nextLine;
                        setCheckString(ckOrg_txt, false);
                    }

                    foreach (string data in mailData)
                    {
                        if (!orgHeadersStr.Contains(data))      //未包含，格式錯誤
                        {
                            result += ORIGIN_FILE_NAME + "格式錯誤: 檔案標頭未包含[" + data + "]" + nextLine;
                            setCheckString(ckOrg_txt, false);
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine(result.Equals(""));
                    if (result.Equals(""))
                    {
                        System.Diagnostics.Debug.WriteLine("result:"+result);
                        setCheckString(ckOrg_txt, true);
                    }
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

        private void ImportPathTextChanged(object sender, EventArgs e)
        {
            header_combox.Items.Clear();
            sendName_combox.Items.Clear();
            sendTo_combox.Items.Clear();

            header_combox.Text = "讀取中...";
            header_combox.Enabled = false;
            if (need_mail_cbx.Checked)
            {
                sendName_combox.Text = "讀取中...";
                sendName_combox.Enabled = false;
                sendTo_combox.Text = "讀取中...";
                sendTo_combox.Enabled = false;
            }

            if (File.Exists(importOrgPath_txtbx.Text))
            {
                headers = readStrArrExcelCellinRow(importOrgPath_txtbx.Text, ORIGIN_HEADER_ROW);

                foreach (string header in headers)
                {
                    header_combox.Items.Add(header);
                }
                header_combox.SelectedIndex = 0;

                if (need_mail_cbx.Checked)
                {
                    sendName_combox.Items.Add("");
                    sendTo_combox.Items.Add("");
                    foreach (string header in headers)
                    {
                        sendName_combox.Items.Add(header);
                        sendTo_combox.Items.Add(header);
                    }
                    sendName_combox.SelectedIndex = 0;
                    sendTo_combox.SelectedIndex = 0;
                }

            }
            else
            {
                header_combox.Text = "";
                sendName_combox.Text = "";
                sendTo_combox.Text = "";
            }

            header_combox.Enabled = true;
            if (need_mail_cbx.Checked)
            {
                sendName_combox.Enabled = true;
                sendTo_combox.Enabled = true;
            }
        }

        private void needMailChanged(object sender, EventArgs e)
        {
            if(need_mail_cbx.Checked)
            {
                reloadsendToData();
            }
            else
            {
                sendName_combox.Text = "";
                sendTo_combox.Text = "";

                sendName_combox.Enabled = false;
                sendTo_combox.Enabled = false;
            }
        }
    }
    
}
