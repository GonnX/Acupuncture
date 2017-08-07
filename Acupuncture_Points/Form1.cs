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

namespace Acupuncture_Points
{
    public partial class Form1 : Form
    {
        private const int Btn_Size = 10;
        private const int File_Field = 2;
        private const int PortData_Size = 150;

        private String[] File = new String[File_Field];
        private List<int> PortData = new List<int>();

        private System.Windows.Forms.TextBox[] Tb;
        private System.Windows.Forms.Button[] Btn;
        private System.Windows.Forms.DataVisualization.Charting.Chart[] Chart;

        private SerialPort port;
        
        delegate void UpdateChart(List<int> list);
        delegate void UpdateChart_single(int data);

        private Thread PrintCahrt = null;

        private int index;

        private bool Flag = false;

        StreamWriter File_Stream_Writer;

        public Form1()
        {
            InitializeComponent();

            this.FormClosing += Form1_FormClosing;

            String[] All_Port = SerialPort.GetPortNames();

            foreach (String item in All_Port)
                ComPort_Combobox.Items.Add(item);
            
            Tb = new System.Windows.Forms.TextBox[Btn_Size]
            {Acupuncture_Points_1_Tb, Acupuncture_Points_2_Tb, Acupuncture_Points_3_Tb, Acupuncture_Points_4_Tb, Acupuncture_Points_5_Tb,
             Acupuncture_Points_6_Tb, Acupuncture_Points_7_Tb, Acupuncture_Points_8_Tb, Acupuncture_Points_9_Tb, Acupuncture_Points_10_Tb};

            Btn = new System.Windows.Forms.Button[Btn_Size]
                { Acupuncture_Points_1_Btn, Acupuncture_Points_2_Btn, Acupuncture_Points_3_Btn, Acupuncture_Points_4_Btn, Acupuncture_Points_5_Btn,
                  Acupuncture_Points_6_Btn, Acupuncture_Points_7_Btn, Acupuncture_Points_8_Btn, Acupuncture_Points_9_Btn, Acupuncture_Points_10_Btn};

            Chart = new System.Windows.Forms.DataVisualization.Charting.Chart[Btn_Size]
                { Acupuncture_Points_1_Chart, Acupuncture_Points_2_Chart, Acupuncture_Points_3_Chart, Acupuncture_Points_4_Chart, Acupuncture_Points_5_Chart,
                  Acupuncture_Points_6_Chart, Acupuncture_Points_7_Chart, Acupuncture_Points_8_Chart, Acupuncture_Points_9_Chart, Acupuncture_Points_10_Chart};

            for (int i = 0; i < Btn_Size; i++)
            {
                Btn[i].Enabled = false;
                Tb[i].Enabled = false;
            }
            //port = new SerialPort("COM14", 9600, Parity.None, 8, StopBits.One);
            //port.Open();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (File_Stream_Writer != null)
                File_Stream_Writer.Close();
            Environment.Exit(Environment.ExitCode);
        }

        void producer()
        {
            port.BaseStream.Flush();
            PortData.Clear();

            File_Stream_Writer.WriteLine("穴位名 : " + Tb[index].Text);

            File_Stream_Writer.Write("\r\n");

            for (int i = 0; i < PortData_Size; i++)
                UpdateChartsingle(port.ReadByte());

            /*
            int counter = 0;

            do
            {
                if(counter != 0)
                {
                    SetupChart();
                    port.Write("R");
                    MessageBox.Show("資料傳輸有誤，正在重新擷取中，請等候...");
                }

                for (int i = 0; i < PortData_Size; i++)
                    UpdateChartsingle(port.ReadByte());
                counter++;
            } while (Flag == true);

            Flag = false;
            */
            File_Stream_Writer.Write("\r\n");
            File_Stream_Writer.Write("\r\n");
            
            /*
            if (PortData.Count != PortData_Size)
            {
                do
                {
                    MessageBox.Show("資料尚未收齊，目前正重新擷取");
                    PortData.Clear();
                    port.Write("R");
                    for (int i = 0; i < PortData_Size; i++)
                        UpdateChartsingle(port.ReadByte());
                } while (PortData.Count != PortData_Size);
            }*/
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
                //if (data <= 0)
                //    Flag = true;
                File_Stream_Writer.Write(data);
                File_Stream_Writer.Write(' ');
                PortData.Add(data);
                Chart[index].Series[0].Points.AddY(data);
            }
        }
        private void Ok_Btn_Click(object sender, EventArgs e)
        {
            System.Text.RegularExpressions.Regex reg1 = new System.Text.RegularExpressions.Regex("[0-9]");

            if ( ( FileName_Tb.Text != "" && Acupuncture_Points_Count_Tb.Text != "" ) && 
                reg1.IsMatch(Acupuncture_Points_Count_Tb.Text) &&
                (Convert.ToInt16(Acupuncture_Points_Count_Tb.Text) > 0 && Convert.ToInt16(Acupuncture_Points_Count_Tb.Text) < 11) &&
                BaudRate.Text != "" && ComPort_Combobox.Text != "")
            {
                port = new SerialPort(ComPort_Combobox.GetItemText(ComPort_Combobox.SelectedItem), Convert.ToInt16(BaudRate.Text), Parity.None, 8, StopBits.One);
                port.Open();

                File[0] = FileName_Tb.Text;
                File[1] = Acupuncture_Points_Count_Tb.Text;

                File_Stream_Writer = new StreamWriter(@"C:\Users\LuoLab\Desktop\" + File[0] + ".txt");

                FileName_Tb.Enabled = false;
                Acupuncture_Points_Count_Tb.Enabled = false;

                for (int i = 0; i < Convert.ToInt16(File[1]); i++)
                {
                    Btn[i].Enabled = true;
                    Tb[i].Enabled = true;
                }
                Ok_Btn.Enabled = false;
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
            area.AxisY.Maximum = 1200;
            area.AxisX.Interval = 10;
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
            if (Tb[index].Text != "")
            {
                Tb[index].Enabled = false;
                SetupChart();
                port.Write("S");
                PrintCahrt = new Thread(new ThreadStart(this.producer));
                PrintCahrt.Start();
            }
            else MessageBox.Show("未輸入穴位名");
        }
    }
}
