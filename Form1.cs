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

            StepTimer[0] = 2;
            for (int i = 0; i < 10; i++)
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


            //if (cboPlan.SelectedIndex == 0) // Lựa chọn Login
            //{

            //    tmrPlan01.Interval = Convert.ToInt16(txtSpeed.Text);
            //    tmrPlan01.Start();
            //}
            if (cboPlan.SelectedIndex == 1) // Lựa chọn Get Keywords
            {

                progressBar1.Maximum = 6; // số lượng các thao tác trong kế hoạch.
                progressBar1.Value = 0;
                tmrPlan01.Interval = Convert.ToInt16(txtSpeed.Text);
                tmrPlan01.Start();
            }

        }

        private void AddList(string str)
        {
            lstStatus.Items.Add(str);
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            alarmCounter++;

            if (alarmCounter == StepTimer[0]) //focus text search
            {
                pt.X = 772;
                pt.Y = 519;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Focus to Search");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[1]) //input text search
            {
                SendKeys.Send(txtKeywords.Text);
                AddList("input Text Search");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[2]) // click search
            {
                pt.X = 1271;
                pt.Y = 516;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click Search");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[3]) // click download button
            {
                pt.X = 1590;
                pt.Y = 988;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click Download");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[4]) // click export to excel
            {
                pt.X = 1563;
                pt.Y = 794;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AddList("Click to excel File");

                progressBar1.Value += 1;
            }

            //if (alarmCounter == StepTimer[5]) // click save as
            //{
            //    pt.X = 1231;
            //    pt.Y = 159;
            //    Cursor.Position = pt;
            //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
            //    AddList("Click to save as");

            //    progressBar1.Value += 1;
            //}

            //if (alarmCounter == StepTimer[6]) // click save
            //{
            //    pt.X = 1011;
            //    pt.Y = 567;
            //    Cursor.Position = pt;
            //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
            //    AddList("Click to save into Computer");

            //    progressBar1.Value += 1;
            //}

            if (alarmCounter == StepTimer[5])
            {
                tmrPlan01.Stop();
                alarmCounter = 0;

                this.WindowState = FormWindowState.Normal;
                AddList("Finished");

                progressBar1.Value += 1;
            }
        }





        private void Form1_Load(object sender, EventArgs e)
        {

            // Lấy các dữ liệu setting
            v_speed = Convert.ToInt32(txtSpeed.Text);

            // Thêm kịch bản
            cboPlan.Items.Clear();
            cboPlan.Items.Add("Login");
            cboPlan.Items.Add("Get keywords");

            // Mở kết nối file excel
            //f.fileName = "Keyword Tool Export -Keyword Suggestions - " + CurrentKeywords;

        }
        private void Importexcel(string path)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
                //excelWorksheet.Cells[3, 3].Value = 34;
                //DataTable dataTable = new DataTable();
                //for (int i = excelWorksheet.Dimension.Start.Column; i <= excelWorksheet.Dimension.End.Column; i++)
                //{
                // dataTable.Columns.Add(excelWorksheet.Cells[1, i].Value.ToString());
                //}
                //for (int i = excelWorksheet.Dimension.Start.Row + 1; i <= excelWorksheet.Dimension.End.Row; i++)
                //{
                //    List<string> listRows = new List<string>();
                //    for (int j = excelWorksheet.Dimension.Start.Column; j <= excelWorksheet.Dimension.End.Column; j++)
                //    {
                //        listRows.Add(excelWorksheet.Cells[i,j].Value.ToString());
                //    }
                //    dataTable.Rows.Add(listRows.ToArray());
                //}



                //dgrListKeywords.DataSource = dataTable;
                //dgrListKeywords.Columns.Add( ("clnKey", "Keywords");
                int i = 1;
                while (excelWorksheet.Cells[i,1].Value != null)
                {
                    dgrListKeywords.Rows.Add(excelWorksheet.Cells[i,1].Value);
                    i++; 
                }
            }
        }
   

        private void openExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Import Excel";
            openFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Importexcel(openFileDialog.FileName);
                    MessageBox.Show("Nhap file thanh cong");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Nhap file khong thanh cong \n" + ex.Message);
                }
            }
        }
        private void ExportExcel(string path)
        {
            Excel.Application application = new Excel.Application();
            application.Application.Workbooks.Add(Type.Missing);
            for (int i = 0; i <= dgrListKeywords.Columns.Count; i++)
            {
                application.Cells[1, i + 1] = dgrListKeywords.Columns[i].HeaderText;
            }
            for (int i = 0; i <= dgrListKeywords.Rows.Count; i++)
            {
                for (int j = 0; j < dgrListKeywords.Columns.Count; j++)
                {
                    application.Cells[i+2, j+1] = dgrListKeywords.Rows[i].Cells[j].Value;
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
                try
                {
                    ExportExcel(saveFileDialog.FileName);
                    MessageBox.Show("Xuat file thanh cong");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xuat file khong thanh cong \n" + ex.Message);
                }
            }
        }

        private void cboPlan_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {

        }

        private void timeSpanChartRangeControlClient1_CustomizeSeries(object sender, DevExpress.XtraEditors.ClientDataSourceProviderCustomizeSeriesEventArgs e)
        {

        }
    }
}
