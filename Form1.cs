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

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using DataTable = System.Data.DataTable;

namespace EveryTeacher
{
    public partial class Form1 : Form
    {

        Excel.Application _xlApp;
        Excel.Workbook _xlWorkBook;
        Excel.Worksheet _xlWorkSheet;
        Excel.Range _range;
        String _file_name = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //指定Excel檔案路徑
            string str = System.Windows.Forms.Application.StartupPath + "\\input_file";
            //-----------------------下拉選單製作
            
        }
        private void btn_select_file_Click(object sender, EventArgs e)
        {
            _file_name = ImportExcelFile();
            //開啟Excel
            _xlApp = new Excel.Application();
            _xlWorkBook = _xlApp.Workbooks.Open(_file_name, 0, true, 5, "", "", true,
              Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
             , "\t", false, false, 0, true, 1, 0);
            //-----------------------worksheet下拉選單
            Excel.Worksheet xlWorkSheet2;
            lb_sys_info.Text = "Excel工作表載入中";
            for (int i = 1; i <= _xlWorkBook.Sheets.Count; i++) //計算總共有幾個工作表
            {
                //取得該工作表
                xlWorkSheet2 = (Excel.Worksheet)_xlWorkBook.Worksheets[i];
                //將該工作表名稱加入下拉選單
               
            }
            lb_sys_info.Text = "Excel工作表載入完成";
            //使用完EXCEL資源釋放
            _xlApp.Quit();
            Form1.KillExcelApp(_xlApp);
        }
        private System.Data.DataTable _tb = new System.Data.DataTable("table");
        DataRow _NewRow;
        Thread _read_execel_data;
        private void btn_select_sheet_Click(object sender, EventArgs e)
        {
            lb_load_data_count.Visible = true;
            _read_execel_data = new Thread(read_execel_data);
            _read_execel_data.Start();
        }
        private void read_execel_data()
        {
            string[] teachers;
            int teacherIndex = 0;
            int teacherCount = 0;
            int wantToReadRows = 200;
            int select_sheet = 0;  //select_sheet第幾個工作表
            string str = System.Windows.Forms.Application.StartupPath;
            //開啟Excel
            _xlApp = new Excel.Application();
            this.Invoke((MethodInvoker)delegate
            {
                dataGridView1.DataSource = null;
                _tb.Clear(); //清空表格
                _tb.Rows.Clear();//清空資料
                _tb.Columns.Clear();
            });
            try
            {
                _xlWorkBook = _xlApp.Workbooks.Open(_file_name, 0, true, 5, "", "", true,
                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                  , "\t", false, false, 0, true, 1, 0);
            }
            catch
            {
                MessageBox.Show("未選擇檔案");
                return;
            }
            //取得下拉選單數值讀取指定工作表
            this.Invoke((MethodInvoker)delegate
            {                
                select_sheet = 1;
                _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets[select_sheet];
                _range = _xlWorkSheet.UsedRange;//讀取Excel列與行資訊
            });
            int rw = _range.Rows.Count;     //直 (總行數)
            int cl = _range.Columns.Count;  //橫 (總列數)
            int sum = rw;
            DataColumn[] colItem = new DataColumn[cl + 1];
            String[] title = new String[cl + 1];
            Excel.Worksheet excelSheet =
           (Excel.Worksheet)_xlWorkBook.Worksheets[select_sheet];
            //-----------------------------表格欄位製作
            for (int i = 1; i <= cl; i++)
            {
                //將Table Titile塞入Excel第一列作為標題 [1,n]
                Excel.Range rng = (Excel.Range)excelSheet.Cells[1, i];
                try
                {
                    colItem[i] = new DataColumn(rng.Value, Type.GetType("System.String"));
                    title[i] = rng.Value;
                }
                catch
                {
                    colItem[i] = new DataColumn(" ", Type.GetType("System.String"));
                }
                _tb.Columns.Add(colItem[i]);
            }
            sum = rw - 1;
            lb_sys_info.Invoke((MethodInvoker)delegate
            {
                lb_sys_info.Text = "Excel工作表內容載入中";
            });

            teachers = new string[wantToReadRows];

            for (int i = 2; i <= wantToReadRows; i++)
            {
                lb_load_data_count.Invoke((MethodInvoker)delegate
                {
                    lb_load_data_count.Text =
                    "尚有 " + (wantToReadRows - i - 1).ToString() + " 筆資料未加入";
                });

                
                _NewRow = _tb.NewRow(); //開新的一列
                for (int j = 1; j <= cl; j++)
                {
                    try
                    {
                        Excel.Range rng = (Excel.Range)excelSheet.Cells[i, j];
                        //讀取 Excel每格內容
                        _NewRow[title[j]] = rng.Value.ToString(); //將取得的值存入陣列
                    }
                    catch
                    {
                        _NewRow[title[j]] = "";
                    }
                }

               text.Invoke((MethodInvoker)delegate
                {
                    try
                    {
                        Excel.Range rng = (Excel.Range)excelSheet.Cells[i, 4];
                        text.Text = "("+ i + ")" +rng.Value.ToString(); //將取得的值存入陣列

                        for(teacherIndex = 0; teacherIndex < teacherCount; teacherIndex++)
                        {
                            if (teachers[teacherIndex] == rng.Value.ToString())
                            {
                                //addNewTeacher = false;
                                break;
                            }
                            /*else
                            {
                                addNewTeacher = true;
                            }*/
                        }
                        if(teacherIndex == teacherCount)
                        {
                            teachers[teacherCount] = rng.Value.ToString();
                            teacherCount++; 
                        }

                      
                    }
                    catch (Exception ex)
                    {
                        text.Text = ex.Message.ToString();
                    }
                });
                


                this.Invoke((MethodInvoker)delegate
                {
                    _tb.Rows.Add(_NewRow);
                });
            }
            this.Invoke((MethodInvoker)delegate
            {
                _tb.AcceptChanges();
                dataGridView1.AutoSizeColumnsMode
                = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                dataGridView1.DataSource = _tb;
                lb_load_data_count.Visible = false;
            });

            text.Invoke((MethodInvoker)delegate
            {
                string allTeachers = "";
                foreach(string teacher in teachers)
                {
                    allTeachers += teacher + "\n";
                }
                text.Text = allTeachers;
            });

            //使用完EXCEL資源釋放
            _xlApp.Quit();
            Form1.KillExcelApp(_xlApp);
            lb_sys_info.Invoke((MethodInvoker)delegate
            {
                lb_sys_info.Text = "Excel工作表內容載入完成";
            });
            _read_execel_data.Abort();
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
    }
}
