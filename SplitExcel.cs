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
using Excel = Microsoft.Office.Interop.Excel;
using static EveryTeacher.Program;
using System.Threading;
using System.Text.RegularExpressions;

namespace EveryTeacher
{
    public partial class SplitExcel : Form
    {
        string orgFilePath = "";
        string splitHeader = "";
        string tchFilePath = "";
        string exportPath = "";
        bool isSendMail = false;
        string sendNameHeader = "";
        string sendToHeader = "";

        SendMail[] sendMail = new SendMail[1024];
        int sendMailIndex = 0;

        public SplitExcel(string postorgFilePath, string postsplitHeader, string posttchFilePath, string postexportPath)
        {
            InitializeComponent();

            orgFilePath = postorgFilePath;
            splitHeader = postsplitHeader;
            tchFilePath = posttchFilePath;
            exportPath = postexportPath;

            this.FormClosed += new FormClosedEventHandler(this.SplitExcelClosed);

        }

        public SplitExcel(string postorgFilePath, string postsplitHeader, string posttchFilePath, string postexportPath, 
                            string postsendName, string postSendTo)
        {
            InitializeComponent();

            orgFilePath = postorgFilePath;
            splitHeader = postsplitHeader;
            tchFilePath = posttchFilePath;
            exportPath = postexportPath;
            isSendMail = true;
            sendNameHeader = postsendName;
            sendToHeader = postSendTo;

            this.FormClosed += new FormClosedEventHandler(this.SplitExcelClosed);

        }

        private void SplitExcelLoad(object sender, EventArgs e)
        {
            tchFileP_txt.Text = "test";
        }

        private void SplitExcelClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit(); //這樣程式才會完全關閉並釋放資源
        }

        private void SplitExcelShown(object sender, EventArgs e)
        {
            tchFileP_txt.Text = "";

            tchFileP_txt.Text = "計算中...";
            int countExFiles = writeTchExcel(orgFilePath, exportPath, tchFilePath);        //產出給老師的檔案、寄信設定

            if(isSendMail)
                genSendMailFile(sendMail, exportPath + "寄給所有人的.csv");

            genLogFile(exportPath, countExFiles);

            Over_btn.Text = "完成";
        }
        
        public int writeTchExcel(string importPath, string exportPath, string tchFile)
        {
            string dstFile = "";

            Excel.Application App = new Excel.Application();

            Excel.Workbook Wbook;
            System.IO.FileInfo xlsAttribute;
            Excel.Worksheet Wsheet;
            Excel.Range row;
            Excel.Range[] cellHeaders;
            int tchWriteIndex = 5;
            int tchDataRowIndex = tchWriteIndex;
            string teacherName = "";
            int otherTchIndex = 0;
            int fileCount = 0;
            string[] headerStrArr = readStrArrExcelCellinRow(tchFile, EXAMPLE_HEADER_ROW);
            cellHeaders = new Excel.Range[headerStrArr.Length];

            //創資料夾
            if (!Directory.Exists(exportPath))
            {
                Directory.CreateDirectory(exportPath);
            }

            tchFileP_txt.Text = "計算中...";

            DataSet ds = Program.ExcelToDataSet(importPath, true);
            if (ds != null)
            {
                DataTable dt = ds.Tables[0];
                DataRowCollection readRows = dt.Rows.Copy();

                tchFile_pbar.Minimum = 0;
                tchFile_pbar.Maximum = dt.Rows.Count;
                tchFile_pbar.Value = 0;
                
                tchFileP_txt.Text = "已產出"+ fileCount + "個檔案";

                while (dt.Rows.Count > 0)
                {
                    teacherName = dt.Rows[0][splitHeader].ToString();
                    otherTchIndex = 0;       //換老師時歸零

                    if(isSendMail)
                    {
                        sendMail[sendMailIndex] = new SendMail();
                        if(!sendNameHeader.Equals(""))
                            sendMail[sendMailIndex].SendName = dt.Rows[0][sendNameHeader].ToString();
                        if(!sendToHeader.Equals(""))
                            sendMail[sendMailIndex].Sendto = dt.Rows[0][sendToHeader].ToString();
                        sendMail[sendMailIndex].Attach = dt.Rows[0][splitHeader].ToString();

                        if(splitHeader.Contains("導師") || splitHeader.Contains("老師"))
                            sendMail[sendMailIndex].Attach += "老師.pdf";
                        else
                            sendMail[sendMailIndex].Attach += ".pdf";
                    }

                    dstFile = exportPath + dt.Rows[0][splitHeader].ToString();
                    if (splitHeader.Contains("導師") || splitHeader.Contains("老師"))
                        dstFile += "老師.xlsx";
                    else
                        dstFile += ".xlsx";

                    Regex.Escape(dstFile);

                    if (!File.Exists(dstFile))
                        File.Copy(tchFile, dstFile);

                    //取得欲寫入的檔案路徑
                    Wbook = App.Workbooks.Open(dstFile, 0, true, 5, "", "", true,
                         Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                          , "\t", false, false, 0, true, 1, 0);

                    //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
                    xlsAttribute = new FileInfo(dstFile);
                    xlsAttribute.Attributes = FileAttributes.Normal;

                    //取得工作表
                    Wsheet = (Excel.Worksheet)Wbook.Sheets[1];

                    foreach (DataRow dataRow in readRows)
                    {
                        bool isNull = (dataRow == null);
                        if (dataRow != null)
                        {
                            if (teacherName.Equals(dataRow[splitHeader].ToString()))        //過濾老師
                            {
                                row = Wsheet.Rows[tchDataRowIndex];

                                /*cellClass = Wsheet.Cells[tchDataRowIndex, Program.INDEX_TCH_CLASS];
                                cellStNum = Wsheet.Cells[tchDataRowIndex, Program.INDEX_TCH_STUDENT_NUM];
                                cellStName = Wsheet.Cells[tchDataRowIndex, Program.INDEX_TCH_STUDENT_NAME];
                                cellStPhone = Wsheet.Cells[tchDataRowIndex, Program.INDEX_TCH_STUDENT_PHONE];
                                cellRelief = Wsheet.Cells[tchDataRowIndex, Program.INDEX_TCH_RELIEF];

                                cellClass.Value2 = dataRow[Program.HEADER_CLASS].ToString();
                                cellStNum.Value2 = dataRow[Program.HEADER_STUDENT_NUM].ToString();
                                cellStName.Value2 = dataRow[Program.HEADER_STUDENT_NAME].ToString();
                                cellStPhone.NumberFormat = "@";
                                if (dataRow[Program.HEADER_STUDENT_PHONE].ToString().Length == 9
                                    && dataRow[Program.HEADER_STUDENT_PHONE].ToString().StartsWith("9"))
                                    cellStPhone.Value = "0" + dataRow[Program.HEADER_STUDENT_PHONE].ToString();
                                else
                                    cellStPhone.Value = dataRow[Program.HEADER_STUDENT_PHONE].ToString();

                                cellRelief.Value2 = dataRow[Program.HEADER_RELIEF].ToString();    */     
                                
                                for(int i = 0; i < cellHeaders.Length; i++)
                                {
                                    cellHeaders[i] = Wsheet.Cells[tchDataRowIndex, i+1];
                                    cellHeaders[i].NumberFormat = "@";
                                    cellHeaders[i].Value2 = dataRow[headerStrArr[i]].ToString();
                                }

                                dt.Rows.RemoveAt(0 + otherTchIndex);
                                //這邊的index代表要刪掉的資料
                                //王(0)、王、陳、陳、王
                                //王(0)、陳、陳、王
                                //陳、陳、王(2)

                                tchDataRowIndex++;
                            }
                            else
                            {
                                otherTchIndex++;        //比對不同時+1
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
                    sendMailIndex++;
                    fileCount++;

                    //關閉EXCEL
                    Wbook.Close();

                    if (readRows.Count == dt.Rows.Count)
                    {
                        MessageBox.Show("原始資料有異常!");
                        break;
                    }
                    else
                    {
                        tchFile_pbar.Invoke((MethodInvoker)delegate
                        {
                            tchFile_pbar.Step = readRows.Count - dt.Rows.Count;
                            tchFile_pbar.PerformStep();
                        });
                        
                        tchFileP_txt.Text = "已產出" + fileCount + "個檔案";

                        readRows = dt.Rows.Copy();
                    }


                    System.Diagnostics.Debug.WriteLine("------------");
                }
            }


            //離開應用程式
            App.Quit();
            KillExcelApp(App);

            return fileCount;
        }

        
        private void Over_btnClick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
