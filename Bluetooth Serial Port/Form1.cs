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
using System.Management;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.IO;


namespace Bluetooth_Serial_Port
{
    public partial class Form1 : Form
    {
        string InputData = String.Empty;
        delegate void SetTextCallback(string text);
        Excel.Application xlApp = new Excel.Application();
        
        
        
        public Form1()
        {
            InitializeComponent();
            serialPort1.DataReceived += new SerialDataReceivedEventHandler(DataReceive);
        }
        
        public class Devices
        {
            public string COM { get; set; }
            public string NAME { get; set; }
        }
        public string PORT;
        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                serialPort1.PortName = PORT;
                serialPort1.BaudRate = 9600;
                serialPort1.Open();
                progressBar1.Value = 100;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        public List<Devices> items = new List<Devices>();
        private void Form1_Load(object sender, EventArgs e)
        {

            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = SerialPort.GetPortNames();
                var portList = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());
                foreach (string s in portList)
                {
                    foreach (string j in portnames)
                    {
                        if (s.Contains(j))
                        {
                            items.Add(new Devices { NAME = s, COM = j });
                        }
                    }
                }
                comboBox1.DataSource = items;
                comboBox1.DisplayMember = "NAME";
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            if (cb.SelectedValue != null)
            {
                Devices d = cb.SelectedValue as Devices;
                PORT = d.COM.ToString();
            }
        }
        private void DataReceive(object obj, SerialDataReceivedEventArgs e)
        {
            String InputData = serialPort1.ReadLine();
            if (InputData != String.Empty)
            {
                
                SetText(InputData);  
            }
        }
        private void SetText(string text)
        {
            if (this.label1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else this.label1.Text = text;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");

                FileInfo excelFile = new FileInfo(@"D:\test222.xlsx");
                excel.SaveAs(excelFile);
            }
        }
    }
}
