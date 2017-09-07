using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Windows.Forms.DataVisualization.Charting;
using System.Threading;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Acupuncture_Points
{
    public partial class Form1 : Form
    {
        private const int Btn_Size = 10;
        private const int File_Field = 2;
        private const int PortData_Size = 150;
        private const int ConboBox_Item_Size = 32;

        private String[] File = new String[File_Field];
        private List<int> PortData = new List<int>();

        private System.Windows.Forms.ComboBox[] Acupuncture_Points_ComboBox;
        private System.Windows.Forms.Button[] Btn;
        private System.Windows.Forms.DataVisualization.Charting.Chart[] Chart;

        private SerialPort port;

        delegate void UpdateChart(List<int> list);
        delegate void UpdateChart_single(int data);

        private Thread PrintCahrt = null;

        private Excel.Application ExcelApp;
        private Excel.Workbook WBook;
        private Excel.Worksheet WSheet;
        //private Excel.Range WRange;

        private String[] ComboBox_Item = new String[ConboBox_Item_Size];

        private int index;

        private int This_Width;
        private int This_Height;

        private int Excel_Row = 1;
        private int Excel_Col = 1;

        public Form1()
        {
            InitializeComponent();

            this.Text = "傅耳電針";

            this.saveToolStripMenuItem.Enabled = false;

            this.FormClosed += Form1_FormClosed;

            String[] All_Port = SerialPort.GetPortNames();

            foreach (String item in All_Port)
                ComPort_Combobox.Items.Add(item);
            
            Acupuncture_Points_ComboBox = new System.Windows.Forms.ComboBox[Btn_Size]
            {Acupuncture_Points_1_Combobox,Acupuncture_Points_2_Combobox,Acupuncture_Points_3_Combobox,Acupuncture_Points_4_Combobox,Acupuncture_Points_5_Combobox,
             Acupuncture_Points_6_Combobox,Acupuncture_Points_7_Combobox,Acupuncture_Points_8_Combobox,Acupuncture_Points_9_Combobox,Acupuncture_Points_10_Combobox};

            Btn = new System.Windows.Forms.Button[Btn_Size]
                { Acupuncture_Points_1_Btn, Acupuncture_Points_2_Btn, Acupuncture_Points_3_Btn, Acupuncture_Points_4_Btn, Acupuncture_Points_5_Btn,
                  Acupuncture_Points_6_Btn, Acupuncture_Points_7_Btn, Acupuncture_Points_8_Btn, Acupuncture_Points_9_Btn, Acupuncture_Points_10_Btn};

            Chart = new System.Windows.Forms.DataVisualization.Charting.Chart[Btn_Size]
                { Acupuncture_Points_1_Chart, Acupuncture_Points_2_Chart, Acupuncture_Points_3_Chart, Acupuncture_Points_4_Chart, Acupuncture_Points_5_Chart,
                  Acupuncture_Points_6_Chart, Acupuncture_Points_7_Chart, Acupuncture_Points_8_Chart, Acupuncture_Points_9_Chart, Acupuncture_Points_10_Chart};

//================================================================ Add ComboBox Items ============================================================================
            for (int i = 1; i < 7; i++)
            {
                for (int j = 0; j < Btn_Size; j++)
                {
                    Acupuncture_Points_ComboBox[j].Items.Add("LH0" + i);
                    Acupuncture_Points_ComboBox[j].Items.Add("LF0" + i);
                    Acupuncture_Points_ComboBox[j].Items.Add("RH0" + i);
                    Acupuncture_Points_ComboBox[j].Items.Add("RF0" + i);

                    if( i < 5 )
                    {
                        Acupuncture_Points_ComboBox[j].Items.Add("LS0" + i);
                        Acupuncture_Points_ComboBox[j].Items.Add("RS0" + i);
                    }
                }
            }
//================================================================================================================================================================
            for (int i = 0; i < Btn_Size; i++)
                Btn[i].Enabled = false;

            ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Add(true);
            WBook = ExcelApp.Workbooks[1];
            WBook.Activate();

            WSheet = (Excel.Worksheet)WBook.Worksheets[1];
            WSheet.Name = "Acupuncture_Points_Measure";

            WSheet.Activate();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if(e.CloseReason == e.CloseReason)
            {
                if (port != null)
                    port.Close();
            }
        }

        void producer()
        {
            port.BaseStream.Flush();
            PortData.Clear();

            for (int i = 0; i < PortData_Size; i++)
                UpdateChartsingle(port.ReadByte());
            Excel_Col++;
            Excel_Row = 1;
        }
        private void UpdateChartsingle(int data)
        {
            if (this.Chart[index].InvokeRequired)     // 檢查是不是同一個執行緒
            {
                UpdateChart_single d = new UpdateChart_single(UpdateChartsingle);    // 不同的執行緒要用到delegate
                this.Invoke(d, new object[] { data });
            }
            else
            {
                ExcelApp.Cells[Excel_Row++, Excel_Col] = data;
                PortData.Add(data);
                Chart[index].Series[0].Points.AddY(data);
            }
        }
        private void Ok_Btn_Click(object sender, EventArgs e)
        {
            if ( ComPort_Combobox.Text != "")
            {

                port = new SerialPort(ComPort_Combobox.GetItemText(ComPort_Combobox.SelectedItem), 115200, Parity.None, 8, StopBits.One);

                try {
                    port.Open();
                }catch(Exception io)
                {
                    MessageBox.Show(io.ToString());
                }
                for (int i = 0; i < Btn_Size; i++)
                    Btn[i].Enabled = true;

                Ok_Btn.Enabled = false;
                this.saveToolStripMenuItem.Enabled = true;
            }
            else MessageBox.Show("未輸入檔名及穴位數、者欄位數超過10個或者有欄位未輸入");
        }
        private void SetupChart()
        {
            Chart[index].Series.Clear();

            ChartArea area = Chart[index].ChartAreas[0];

            Series series1 = new Series() ;
            Chart[index].Series.Add(series1);
                        
            //折線圖
            series1.ChartType = SeriesChartType.Line;
            series1.Color = Color.Blue;
            series1.Font = new System.Drawing.Font("標楷體", 12);
            series1.ChartType = SeriesChartType.Line;
            area.AxisX.Maximum = 150;
            area.AxisX.Minimum = 0;
            area.AxisY.Maximum = 260;
            area.AxisY.Minimum = 0;
            area.AxisX.Interval = 10;

            int StartOffset = 0;
            int EndOffset = 15;

            string[] Label_Name = { "1" , "2" ,"3", "4", "5", "6" , "7", "8", "9" , "10", "11", "12" , "13", "14", "15" };

            foreach (string item in Label_Name)
            {
                CustomLabel Label = new CustomLabel(StartOffset, EndOffset, item, 0, LabelMarkStyle.None);
                area.AxisX.CustomLabels.Add(Label);
                StartOffset += 10;
                EndOffset += 10;
            }
        }
        private void Get_Btn_Index(Button button)
        {
            for(int i = 0; i < Btn.Length; i++)
            {
                if (button.Name.Equals(Btn[i].Name))
                    index = i;
            }
        }
        private void Acupuncture_Points_Btn_Click(object sender, EventArgs e)
        {
            Button Button = (Button)sender;
            Get_Btn_Index(Button);
            if (Acupuncture_Points_ComboBox[index].Text != "")
            {
                ExcelApp.Cells[Excel_Row++, Excel_Col] = Acupuncture_Points_ComboBox[index].GetItemText(Acupuncture_Points_ComboBox[index].Text);
                SetupChart();
                port.Write("S");
                PrintCahrt = new Thread(new ThreadStart(this.producer));
                PrintCahrt.Start();
            }
            else MessageBox.Show("未輸入穴位名");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            This_Width = Screen.PrimaryScreen.Bounds.Width;
            This_Height = Screen.PrimaryScreen.Bounds.Height;
            
            this.Width = This_Width - This_Width / 4;
            this.Height = This_Height - This_Height / 4;
        }
        private void Get_ComboBox_Index(ComboBox combobox)
        {
            for (int i = 0; i < Btn.Length; i++)
            {
                if (combobox.Name.Equals(Acupuncture_Points_ComboBox[i].Name))
                    index = i;
            }
        }
        private void Combobox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                ComboBox combobox = (ComboBox)sender;
                //foreach (ComboBox i in Acupuncture_Points_ComboBox)
                //    i.Items.Add(combobox.Text);
                Get_ComboBox_Index(combobox);
                combobox.Items.Add(combobox.Text);
            }
        }

        private void saveToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            saveFileDialog1.DefaultExt = ".xlsx";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                /*
                if (Path.EndsWith("\\"))
                    WBook.SaveCopyAs(Path + FileName_Tb.Text + ".xlsx");
                else
                    WBook.SaveCopyAs(Path + "\\" + FileName_Tb.Text + ".xlsx");
                */
                WBook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                MessageBox.Show("Excel檔案新增成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}