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

namespace EveryTeacher
{
    public partial class SplitExcel : Form
    {
        string orgFilePath = "";
        string tchFilePath = "";
        string depFilePath = "";
        string colFilePath = "";
        string exportPath = "";

        SendMail[] sendMail = new SendMail[1024];
        int sendMailIndex = 0;

        public SplitExcel(string postorgFilePath, string posttchFilePath, 
            string postdepFilePath, string postcolFilePath, string postexportPath)
        {
            InitializeComponent();

            orgFilePath = postorgFilePath;
            tchFilePath = posttchFilePath;
            depFilePath = postdepFilePath;
            colFilePath = postcolFilePath;
            exportPath = postexportPath;

            this.FormClosed += new FormClosedEventHandler(this.SplitExcelClosed);

        }

        private void SplitExcelLoad(object sender, EventArgs e)
        {
            tchFileP_txt.Text = "test";
        }

        private void SplitExcelClosed(object sender, FormClosedEventArgs e)
        {
            //Program.KillExcelApp();

            Application.Exit(); //這樣程式才會完全關閉並釋放資源
        }

        private void SplitExcelShown(object sender, EventArgs e)
        {
            tchFileP_txt.Text = "";
            depFileP_txt.Text = "";
            colFileP_txt.Text = "";

            tchFileP_txt.Text = "計算中...";
            writeTchExcel(orgFilePath, exportPath, tchFilePath);        //產出給老師的檔案、寄信設定
            genSendMailFile(sendMail, exportPath + DIR_NAME_TEACHERS + "寄給所有人的.csv");

            sendMail = new SendMail[1024];
            sendMailIndex = 0;

            depFileP_txt.Text = "計算中...";
            Thread.Sleep(500);
            Application.DoEvents();
            writeDepExcel(orgFilePath, exportPath, depFilePath);        //產出給系主任的檔案
            
            genSendMailFile(sendMail, exportPath + DIR_NAME_DEPARTMENT + "寄給所有人的.csv");

            sendMail = new SendMail[1024];
            sendMailIndex = 0;

            colFileP_txt.Text = "計算中...";
            Thread.Sleep(500);
            Application.DoEvents();
            writeColExcel(orgFilePath, exportPath, colFilePath);        //產出給院長的檔案

            genSendMailFile(sendMail, exportPath + DIR_NAME_COLLEGE + "寄給所有人的.csv");

            Over_btn.Text = "完成";
        }
        
        public void writeTchExcel(string importPath, string exportPath, string tchFile)
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
            Excel.Range firstDataRow;// = Wsheet.Rows[5];
            Excel.Range row;
            Excel.Range cellClass;
            Excel.Range cellStNum;
            Excel.Range cellStName;
            Excel.Range cellStPhone;
            Excel.Range cellRelief;
            int tchWriteIndex = 5;
            int tchDataRowIndex = tchWriteIndex;
            string teacherName = "";
            int otherTchIndex = 0;
            int fileCount = 0;



            //創資料夾
            if (!Directory.Exists(exportPath + DIR_NAME_TEACHERS))
            {
                Directory.CreateDirectory(exportPath + DIR_NAME_TEACHERS);
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

                //tchFileP_txt.Text = tchFile_pbar.Value + "/" + tchFile_pbar.Maximum;
                tchFileP_txt.Text = "已產出"+ fileCount + "個檔案";

                while (dt.Rows.Count > 0)
                {
                    teacherName = dt.Rows[0][Program.HEADER_TEACHERS].ToString();
                    otherTchIndex = 0;       //換老師時歸零

                    sendMail[sendMailIndex] = new SendMail();
                    sendMail[sendMailIndex].SendName = teacherName;
                    sendMail[sendMailIndex].Sendto = dt.Rows[0][Program.HEADER_TCH_EMAIL].ToString();


                    foreach (DataRow dataRow in readRows)
                    {
                        bool isNull = (dataRow == null);
                        if (dataRow != null)
                        {
                            if (teacherName.Equals(dataRow[Program.HEADER_TEACHERS].ToString()))        //過濾老師
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    dataRow[Program.HEADER_STUDENT_NAME].ToString() + "," +
                                    dataRow[Program.HEADER_TEACHERS].ToString() + "," +
                                    dataRow[Program.HEADER_TCH_EMAIL].ToString());

                                val += dataRow[Program.HEADER_CLASS].ToString();
                                row = Wsheet.Rows[tchDataRowIndex];

                                cellClass = Wsheet.Cells[tchDataRowIndex, Program.INDEX_TCH_CLASS];
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

                                cellRelief.Value2 = dataRow[Program.HEADER_RELIEF].ToString();


                                
                                if (dstFile.Equals(""))
                                {
                                    dstFile = exportPath + DIR_NAME_TEACHERS
                                        + dataRow[Program.HEADER_TEACHERS].ToString() 
                                        + "老師.xlsx";

                                    sendMail[sendMailIndex].Attach = dataRow[Program.HEADER_TEACHERS].ToString()
                                        + "老師.pdf";

                                    if (!File.Exists(dstFile))
                                        File.Copy(tchFile, dstFile);

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
                        tchFile_pbar.Invoke((MethodInvoker)delegate
                        {
                            tchFile_pbar.Step = readRows.Count - dt.Rows.Count;
                            tchFile_pbar.PerformStep();
                        });

                        //tchFileP_txt.Text = tchFile_pbar.Value + "/" + tchFile_pbar.Maximum;
                        tchFileP_txt.Text = "已產出" + fileCount + "個檔案";

                        readRows = dt.Rows.Copy();
                    }


                    System.Diagnostics.Debug.WriteLine("------------");
                }
            }

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
            KillExcelApp(App);
        }


        public void writeDepExcel(string importPath, string exportPath, string depFile)
        {
            string val = "";
            string dstFile = "";

            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(depFile, 0, true, 5, "", "", true,
                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                  , "\t", false, false, 0, true, 1, 0);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(depFile);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[1];
            Excel.Range firstDataRow;// = Wsheet.Rows[5];
            Excel.Range row;
            Excel.Range cellTchName;
            Excel.Range cellClass;
            Excel.Range cellStNum;
            Excel.Range cellStName;
            Excel.Range cellStPhone;
            Excel.Range cellRelief;
            int depWriteIndex = 5;
            int depDataRowIndex = depWriteIndex;
            string departmentName = "";
            int fileCount = 0;
            int otherDepIndex = 0;

            if (!Directory.Exists(exportPath + DIR_NAME_DEPARTMENT))
            {
                Directory.CreateDirectory(exportPath + DIR_NAME_DEPARTMENT);
            }
            depFileP_txt.Text = "計算中...";

            DataSet ds = Program.ExcelToDataSet(importPath, true);
            if (ds != null)
            {
                DataTable dt = ds.Tables[0];
                DataRowCollection readRows = dt.Rows.Copy();

                depFile_pbar.Minimum = 0;
                depFile_pbar.Maximum = dt.Rows.Count;
                depFile_pbar.Value = 0;

                depFileP_txt.Text = "已產出" + fileCount + "個檔案";

                while (dt.Rows.Count > 0)
                {
                    departmentName = dt.Rows[0][Program.HEADER_DEPERTMENT].ToString();
                    otherDepIndex = 0;       //換系時歸零

                    foreach (DataRow dataRow in readRows)
                    {
                        bool isNull = (dataRow == null);
                        if (dataRow != null)
                        {
                            if (departmentName.Equals(dataRow[Program.HEADER_DEPERTMENT].ToString()))        //過濾老師
                            {
                                val += dataRow[Program.HEADER_CLASS].ToString();
                                row = Wsheet.Rows[depDataRowIndex];

                                cellTchName = Wsheet.Cells[depDataRowIndex, Program.INDEX_DEP_TCH_NAME];
                                cellClass = Wsheet.Cells[depDataRowIndex, Program.INDEX_DEP_CLASS];
                                cellStNum = Wsheet.Cells[depDataRowIndex, Program.INDEX_DEP_STUDENT_NUM];
                                cellStName = Wsheet.Cells[depDataRowIndex, Program.INDEX_DEP_STUDENT_NAME];
                                cellStPhone = Wsheet.Cells[depDataRowIndex, Program.INDEX_DEP_STUDENT_PHONE];
                                cellRelief = Wsheet.Cells[depDataRowIndex, Program.INDEX_DEP_RELIEF];

                                cellTchName.Value2 = dataRow[Program.HEADER_TEACHERS].ToString();
                                cellClass.Value2 = dataRow[Program.HEADER_CLASS].ToString();
                                cellStNum.Value2 = dataRow[Program.HEADER_STUDENT_NUM].ToString();
                                cellStName.Value2 = dataRow[Program.HEADER_STUDENT_NAME].ToString();
                                cellStPhone.NumberFormat = "@";
                                if (dataRow[Program.HEADER_STUDENT_PHONE].ToString().Length == 9
                                    && dataRow[Program.HEADER_STUDENT_PHONE].ToString().StartsWith("9"))
                                    cellStPhone.Value = "0" + dataRow[Program.HEADER_STUDENT_PHONE].ToString();
                                else
                                    cellStPhone.Value = dataRow[Program.HEADER_STUDENT_PHONE].ToString();

                                cellRelief.Value2 = dataRow[Program.HEADER_RELIEF].ToString();


                                System.Diagnostics.Debug.WriteLine("write:" + dataRow[Program.HEADER_STUDENT_NAME].ToString());

                                if (dstFile.Equals(""))
                                {
                                    dstFile = exportPath + DIR_NAME_DEPARTMENT
                                        + dataRow[Program.HEADER_DEPERTMENT].ToString() + ".xlsx";

                                    sendMail[sendMailIndex] = new SendMail();
                                    sendMail[sendMailIndex].Attach = dataRow[Program.HEADER_DEPERTMENT].ToString()
                                        + ".pdf";

                                    if (!File.Exists(dstFile))
                                        File.Copy(depFile, dstFile);

                                }

                                dt.Rows.RemoveAt(0 + otherDepIndex);    

                                depDataRowIndex++;
                            }
                            else
                            {
                                otherDepIndex++;
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

                    depDataRowIndex = depWriteIndex;
                    dstFile = "";
                    sendMailIndex++;
                    fileCount++;

                    Wbook = App.Workbooks.Open(depFile, 0, true, 5, "", "", true,
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
                        depFile_pbar.Invoke((MethodInvoker)delegate
                        {
                            depFile_pbar.Step = readRows.Count - dt.Rows.Count;
                            depFile_pbar.PerformStep();
                        });

                        depFileP_txt.Text = "已產出" + fileCount + "個檔案";

                        readRows = dt.Rows.Copy();
                    }


                    System.Diagnostics.Debug.WriteLine("------------");
                }
            }

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
            KillExcelApp(App);
        }


        public void writeColExcel(string importPath, string exportPath, string colFile)
        {
            string val = "";
            string dstFile = "";

            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(colFile, 0, true, 5, "", "", true,
                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                  , "\t", false, false, 0, true, 1, 0);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(colFile);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets[1];
            Excel.Range firstDataRow;// = Wsheet.Rows[5];
            Excel.Range row;// = Wsheet.Rows[5];
            Excel.Range cellDep;
            Excel.Range cellClass;
            Excel.Range cellStNum;
            Excel.Range cellStName;
            Excel.Range cellStPhone;
            Excel.Range cellRelief;
            int colWriteIndex = 5;
            int colDataRowIndex = colWriteIndex;
            string collegeName = "";
            int fileCount = 0;
            int otherColIndex = 0;

            if (!Directory.Exists(exportPath + DIR_NAME_COLLEGE))
            {
                Directory.CreateDirectory(exportPath + DIR_NAME_COLLEGE);
            }
            colFileP_txt.Text = "計算中...";

            DataSet ds = Program.ExcelToDataSet(importPath, true);
            if (ds != null)
            {
                DataTable dt = ds.Tables[0];
                DataRowCollection readRows = dt.Rows.Copy();

                colFile_pbar.Minimum = 0;
                colFile_pbar.Maximum = dt.Rows.Count;
                colFile_pbar.Value = 0;

                colFileP_txt.Text = "已產出" + fileCount + "個檔案";

                while (dt.Rows.Count > 0)
                {
                    collegeName = dt.Rows[0][Program.HEADER_COLLEGE].ToString();
                    otherColIndex = 0;
                    foreach (DataRow dataRow in readRows)
                    {
                        bool isNull = (dataRow == null);
                        if (dataRow != null)
                        {
                            if (collegeName.Equals(dataRow[Program.HEADER_COLLEGE].ToString()))        //過濾老師
                            {
                                val += dataRow[Program.HEADER_CLASS].ToString();
                                row = Wsheet.Rows[colDataRowIndex];

                                cellDep = Wsheet.Cells[colDataRowIndex, Program.INDEX_COL_DEP];
                                cellClass = Wsheet.Cells[colDataRowIndex, Program.INDEX_COL_CLASS];
                                cellStNum = Wsheet.Cells[colDataRowIndex, Program.INDEX_COL_STUDENT_NUM];
                                cellStName = Wsheet.Cells[colDataRowIndex, Program.INDEX_COL_STUDENT_NAME];
                                cellStPhone = Wsheet.Cells[colDataRowIndex, Program.INDEX_COL_STUDENT_PHONE];
                                cellRelief = Wsheet.Cells[colDataRowIndex, Program.INDEX_COL_RELIEF];

                                cellDep.Value2 = dataRow[Program.HEADER_DEPERTMENT].ToString();
                                cellClass.Value2 = dataRow[Program.HEADER_CLASS].ToString();
                                cellStNum.Value2 = dataRow[Program.HEADER_STUDENT_NUM].ToString();
                                cellStName.Value2 = dataRow[Program.HEADER_STUDENT_NAME].ToString();
                                cellStPhone.NumberFormat = "@";
                                if (dataRow[Program.HEADER_STUDENT_PHONE].ToString().Length == 9
                                    && dataRow[Program.HEADER_STUDENT_PHONE].ToString().StartsWith("9"))
                                    cellStPhone.Value = "0" + dataRow[Program.HEADER_STUDENT_PHONE].ToString();
                                else
                                    cellStPhone.Value = dataRow[Program.HEADER_STUDENT_PHONE].ToString();

                                cellRelief.Value2 = dataRow[Program.HEADER_RELIEF].ToString();


                                System.Diagnostics.Debug.WriteLine("write:" + dataRow[Program.HEADER_STUDENT_NAME].ToString());

                                if (dstFile.Equals(""))
                                {
                                    dstFile = exportPath + DIR_NAME_COLLEGE
                                        + dataRow[Program.HEADER_COLLEGE].ToString() + ".xlsx";

                                    sendMail[sendMailIndex] = new SendMail();
                                    sendMail[sendMailIndex].Attach = dataRow[Program.HEADER_COLLEGE].ToString()
                                          + ".pdf";

                                    if (!File.Exists(dstFile))
                                        File.Copy(colFile, dstFile);

                                }

                                dt.Rows.RemoveAt(0 + otherColIndex);    

                                colDataRowIndex++;
                            }
                            else
                            {
                                otherColIndex++;
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

                    colDataRowIndex = colWriteIndex;
                    dstFile = "";
                    sendMailIndex++;
                    fileCount++;

                    Wbook = App.Workbooks.Open(colFile, 0, true, 5, "", "", true,
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
                        colFile_pbar.Invoke((MethodInvoker)delegate
                        {
                            colFile_pbar.Step = readRows.Count - dt.Rows.Count;
                            colFile_pbar.PerformStep();
                        });


                        colFileP_txt.Text = "已產出" + fileCount + "個檔案";

                        readRows = dt.Rows.Copy();
                    }


                    System.Diagnostics.Debug.WriteLine("------------");
                }
            }

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
            KillExcelApp(App);
        }

        private void Over_btnClick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
