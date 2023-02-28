using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace GetKeywords
{
    public partial class Form1 : Form
    {

        static System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();
        static int alarmCounter = 0;
        static bool exitFlag = false;
        static Point pt;
        private int[] StepTimer = new int[100];
        private int[] NextStepDelay = new int[100];
        static int v_speed = 1;

        private ExcelConnect f;
        private string CurrentKeywords = null;

        private int KeyIndex = 1; // y nghia bien

        string h_ID = "";
        public Form1()
        {
            InitializeComponent();
            // Lấy các dữ liệu setting

            // Delay Time after event;
            NextStepDelay[0] = 2; //focus text search
            NextStepDelay[1] = 2; //input text search
            NextStepDelay[2] = 5; // click search
            NextStepDelay[3] = 3; // click download button
            NextStepDelay[4] = 3; // click excel
            NextStepDelay[5] = 2; // click save as
            NextStepDelay[6] = 2; // click save
            NextStepDelay[7] = 2;
            NextStepDelay[8] = 2;
            NextStepDelay[9] = 2;
            NextStepDelay[10] = 2;
            NextStepDelay[11] = 2;
            NextStepDelay[12] = 2;

            StepTimer[0] = 2;
            for (int i = 0; i < 20; i++)
            {
                StepTimer[i + 1] = StepTimer[i] + NextStepDelay[i];
            }
        }

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;

        [DllImport("user32.dll")]

        // Định nghĩa hàm mouse_event() từ thư viện user32.dll
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        // ...
        private void btnStart_Click(object sender, EventArgs e)
        {


            if (cboPlan.SelectedIndex == 0) // Lựa chọn Login
            {
                progressBar1.Maximum = 13; // số lượng các thao tác trong kế hoạch.
                progressBar1.Value = 0;
                tmrPlan02.Interval = Convert.ToInt16(txtSpeed.Text);
                tmrPlan02.Start();
            }
            if (cboPlan.SelectedIndex == 1) // Lựa chọn Get Keywords
            {

                progressBar1.Maximum = 6; // số lượng các thao tác trong kế hoạch.
                progressBar1.Value = 0;
                tmrPlan01.Interval = Convert.ToInt16(txtSpeed.Text);
                tmrPlan01.Start();
            }
            if (cboPlan.SelectedIndex == 2) // Lựa chọn Get Keywords
            {

                progressBar1.Maximum = 8; // số lượng các thao tác trong kế hoạch.
                progressBar1.Value = 0;
                tmrPlan03.Interval = Convert.ToInt16(txtSpeed.Text);
                tmrPlan03.Start();
            }
        }

        private void AddList(string str)
        {
            lstStatus.Items.Add(str);
        }


        private void timer1_Tick(object sender, EventArgs e)
        {

        }





        private void Form1_Load(object sender, EventArgs e)
        {

            // Lấy các dữ liệu setting
            v_speed = Convert.ToInt32(txtSpeed.Text);
            h_ID = txtmatkhau.Text;

            // Thêm kịch bản
            cboPlan.Items.Clear();
            cboPlan.Items.Add("Login");
            cboPlan.Items.Add("Get keywords");
            cboPlan.Items.Add("Dowload Keyword tiep theo");

            // Mở kết nối file excel
            //f.fileName = "Keyword Tool Export -Keyword Suggestions - " + CurrentKeywords;
        }
        /// <summary>
        /// Day la doan nhap file excel thu 1
        /// </summary>
        /// <param name="path"></param>
        private void Importexcel(string path)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
                int i = 1;
                while ((excelWorksheet.Cells[i+1,1].Value != null))
                {
                    dgrListKeywords.Rows.Add(excelWorksheet.Cells[i+1, 1].Value, excelWorksheet.Cells[i+1, 2].Value) ;
                    i++;
                }
            }
        }
        /// <summary>
        /// doan nhap file cho vong lay keywords tiep theo
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void ImportExcelCircle(string path)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
                int i = 1;
                while ((excelWorksheet.Cells[i + 1, 1].Value != null) && (Convert.ToInt16(excelWorksheet.Cells[i + 1, 2].Value) > 1000))
                {
                    dgrListKeywords.Rows.Add(excelWorksheet.Cells[i + 1, 1].Value, excelWorksheet.Cells[i + 1, 2].Value);
                    i++;
                }
            }
        }

        private void openExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Import Excel";
            openFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
            //if (openFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    try
            //    {
            //        Importexcel(openFileDialog.FileName);
            //        MessageBox.Show("Nhap file thanh cong");
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Nhap file khong thanh cong \n" + ex.Message);
            //    }
            //}
        }
        private void ExportExcel(string path)
        { 
            // Sau se toi uu toc do doc ghi bang su dung DataTable, DataSource.

            Excel.Application application = new Excel.Application();
            application.Application.Workbooks.Add(Type.Missing);
            for (int i = 0; i < dgrListKeywords.Columns.Count; i++)
            {
                application.Cells[1, i + 1] = dgrListKeywords.Columns[i].HeaderText;
            }
            for (int i = 0; i < dgrListKeywords.Rows.Count; i++)
            {
                for (int j = 0; j < dgrListKeywords.Columns.Count-1; j++)
                {
                    application.Cells[i + 2, j + 1] = dgrListKeywords.Rows[i].Cells[j].Value;
                }
            }
            application.Columns.AutoFit();
            application.ActiveWorkbook.SaveCopyAs(path);
            application.ActiveWorkbook.Saved = true;
        }
        private void saveExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel";
            saveFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //try
                {
                    ExportExcel(saveFileDialog.FileName);
                    MessageBox.Show("Xuat file thanh cong");
                }
                //catch (Exception ex)
                //{
                //    MessageBox.Show("Xuat file khong thanh cong \n" + ex.Message);
                //}
            }
        }

        private void cboPlan_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            alarmCounter++;
            if (alarmCounter == StepTimer[0]) //CLICK TO danh sach
            {
                pt.X = 1313;
                pt.Y = 129;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Focus to danh sach");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[1]) //CLICK TO LOGIN
            {
                pt.X = 389;
                pt.Y = 433;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Focus to login");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[2]) //focus text taikhoan
            {
                pt.X = 1087;
                pt.Y = 399;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Focus to TaiKhoan");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[3]) //input to tai khoan
            {
                SendKeys.Send(txttaikhoan.Text);
                AddList("input Text TaiKhoan");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[4]) //focus text mat khau
            {
                pt.X = 1074;
                pt.Y = 480;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Focus to mat khau");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[5]) //input text mat khau
            {
                SendKeys.Send(txtmatkhau.Text);
                AddList("input Text MatKhau");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[6]) // click login
            {
                pt.X = 588;
                pt.Y = 580;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click login");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[7]) //focus text search
            {
                pt.X = 823;
                pt.Y = 498;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Focus to Search");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[8]) //input text search
            {
                SendKeys.Send(txtKeywords.Text);
                AddList("input Text search");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[9]) // click search
            {
                pt.X = 1272;
                pt.Y = 506;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click Search");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[10]) // click download button
            {
                pt.X = 1276;
                pt.Y = 949;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click Download");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[11]) // click export to excel
            {
                pt.X = 1285;
                pt.Y = 758;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click to excel File");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[12])
            {
                tmrPlan02.Stop();
                alarmCounter = 0;

                this.WindowState = FormWindowState.Normal;
                AddList("Finished");

                progressBar1.Value += 1;
            }
        }

        private void tmrPlan03_Tick(object sender, EventArgs e)
        {
            alarmCounter++;

            if (alarmCounter == StepTimer[0]) //focus text search
            {
                pt.X = 811;
                pt.Y = 262;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Focus to Search");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[1]) //input text search
            {
                //SendKeys.SendWait("+(CTRL)");
                //SendKeys.SendWait("+(A)");
                SendKeys.Send("^(A)");
                AddList("input Text Search");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[2]) //input text search
            {
                txtKeywords.Text = dgrListKeywords.Rows[KeyIndex].Cells[2].Value.ToString();
                SendKeys.Send(txtKeywords.Text);
                AddList("input Text Search");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[3]) // click search
            {
                pt.X = 1287;
                pt.Y = 259;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click Search");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[4]) // click download button
            {
                pt.X = 1275;
                pt.Y = 946;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click Download");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[5]) // click export to excel
            {
                pt.X = 1292;
                pt.Y = 757;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click to excel File");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[6])
            {
                // Kiem tra file excel co ton tai khong
                // - Co ton tai: Import file excel them vao grid
                // - Khong ton tai: ????

                //tmrPlan03.Stop()
                string filePath = @"C:\Users\Admin\Desktop\Keyword Tool Export - Keyword Suggestions - "+ txtKeywords.Text +".xlsx";
                if (System.IO.File.Exists(filePath) == true)
                {
                    MessageBox.Show("Tồn tại tập tin ");
                }
                else
                {
                    MessageBox.Show("Không tồn tại Tập tin ");
                }
                alarmCounter = 0;

                this.WindowState = FormWindowState.Normal;
                AddList("Finished");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[7]) // click download button
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Import Excel";
                openFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ImportExcelCircle(openFileDialog.FileName);
                        MessageBox.Show("Nhap file thanh cong");
                        KeyIndex++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Nhap file khong thanh cong \n" + ex.Message);
                    }
                }
                alarmCounter = 0;

                this.WindowState = FormWindowState.Normal;
                AddList("Finished");

                progressBar1.Value = 0;
            }
        }

        private void txtSpeed_TextChanged(object sender, EventArgs e)
        {

        }

        private void txttaikhoan_TextChanged(object sender, EventArgs e)
        {
            h_ID = txtmatkhau.Text;
        }

        private void tmrPlan01_Tick(object sender, EventArgs e)
        {

        }
    }
}