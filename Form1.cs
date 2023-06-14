using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using System.Xml;
using static System.Windows.Forms.LinkLabel;
using Excel = Microsoft.Office.Interop.Excel;



namespace xRite_Interface_v1
{
    public partial class Form1 : Form
    {
        string portName;
        SerialPort port;
        int measureCount = 0;

        //initialize Form to look for Arduino and set form interactions
        public Form1()
        {
            //string portName = AutodetectArduinoPort();
            //port = new SerialPort(portName, 9600);
           // port.Open();
            InitializeComponent();
            if (portName == null)
            {

            }
            else
            {
                string portName = AutodetectArduinoPort();
                port = new SerialPort(portName, 9600);
                port.Open();
            }
            messageBar.Text = "Power on Tool then click 'Connect Tool'";
            measure.Enabled = false;
            reset.Enabled = false;
            inputcount.Enabled = false;
        }

        //routine to connect tool to datacatcher and initialize stage maching
        private void toolConnect_Click(object sender, EventArgs e)
        {
            port.Write("1");
            messageBar.Text = "Wait for Data Catcher popup, ensure connection";
            saver();
            if (dataCatcher() == true)
            {
                messageBar.Text = "Input # of measurements in textbox";

                reset.Enabled = true;
                inputcount.Enabled = true;
                toolConnect.Enabled = false;
            }
            else
            {
                messageBar.Text = "Click 'Connect Tool' to try connecting again";

                toolConnect.Enabled = true;
            }
        }

        
        private void Activate_Click(object sender, EventArgs e)
        {
            toolConnect.Enabled = false;
            inputcount.Enabled = true;
        }

        //measure on stages and fill in text box with data
        private void measure_Click(object sender, EventArgs e)
        {
            port.Write("2");

            datafill.Select(datafill.Text.Length, 0);
            datafill.Focus();
            datafill.ScrollToCaret();
            measure.Enabled = false;
            inputcount.Enabled = false;
            measureCount++;
            Thread.Sleep(1000);
            measure.Enabled = true;
            progressBar();

        }

        //routine to start datacatcher to connect to tool
        private bool dataCatcher()
        {
            Process dataCatcher = new Process();


            try
            {
                dataCatcher.StartInfo.FileName = @"C:\Program Files (x86)\X-Rite\DataCatcher\DataCatcher.exe";
                dataCatcher.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
                dataCatcher.Start();
                Thread.Sleep(1000);
                dataCatcher.WaitForInputIdle();
                MessageBoxButtons button = MessageBoxButtons.YesNo;
                string message = "Is Data Catcher Connected?";
                string title = " Check Data Catcher";
                DialogResult result = MessageBox.Show(message, title, button);
                if (result == DialogResult.Yes)
                {
                    messageBar.Text = "Tool Connected";
                    toolConnect.Enabled = false;
                    return true;
                }

                else if (result == DialogResult.No)
                {
                    messageBar.Text = "Please connect to Data Catcher";
                    return false;

                }

                else
                {
                    messageBar.Text = "Error, Please Try Again";
                    return false;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                return false;
            }
        }

        //used in case of looking for existing workbook and not a new one
        public string FileName()
        {
            string filePath = null;
            filePath = (@"C:\\Users\\delisle\\Desktop\\HP\\xriteLAB.xlsx");
            return filePath;
        }

        //routine to open excel and fill in data in a new spreadsheet, saves, then closes
        private void SS(string filePath, string filesave)
        {
            
            if (filePath != null)
            {                
                Excel.Application xlApp = new Excel.Application();
                //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "\t", true, false, 0, true, false, false);
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Add("");
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                
                Excel.Worksheet rSheet = xlWorkbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                xlApp.Visible = true;

                int p = 0;
                int z = 0;
                int h = 1;            
                int number = 0;
                int c =0;
                                           
                string value = inputcount.Text;
                bool isParsable = Int32.TryParse(value, out number);
                string phrase = datafill.Text;
                char[] delims = new[] { '\r', '\n' };
                string[] words = phrase.Split(delims, StringSplitOptions.RemoveEmptyEntries);              
                int arrlength = words.Length;
                int collength = words.Length / 47;
                string[,] words2 = new string[words.Length, collength];

                Excel.Range length2left = xlWorksheet.Cells[arrlength, 1];
                Excel.Range length2right = xlWorksheet.Cells[arrlength/collength, collength];
               

                for (c = 0; c < collength; c++)
                {

                    for (p = c*arrlength/collength; p < (c+1) * arrlength / collength; p++)
                    {
                        words2[z, c] = words[p];
                        z++;
                    }
                    z= 0;
                    xlWorksheet.get_Range("A1", length2right).Value2 = words2;
                }

                xlRange = xlWorksheet.get_Range("A1", length2right);
                xlRange.EntireColumn.AutoFit();


                rSheet.Cells[1, 2] = "Density";
                rSheet.Cells[1, 3] = "C l*a*b1";
                rSheet.Cells[1, 4] = "C l*a*b2";
                rSheet.Cells[1, 5] = "C l*a*b3";
                for (h = 0; h < number; h++)
                {
                    
                    rSheet.Cells[1 + (5 * h), 1] = h + 1;
                    rSheet.Cells[2 + (5 * h), 1] = "M0";                
                    rSheet.Cells[3 + (5 * h), 1] = "M1";
                    rSheet.Cells[4 + (5 * h), 1] = "M2";                                      
                    rSheet.Cells[5 + (5 * h), 1] = "M3";
                    if (xlWorksheet.Cells[18, h + 1].Value2 == null)
                    {
                        string M0 = "";
                        string M1 = "";
                        string M2 = "";
                        string M3 = "";
                        string[] data = { M0, M1, M2, M3 };
                        for (int d = 0; d < data.Length; d++)
                        {
                            rSheet.Cells[5 * h + 2 + d, 2] = data[d];
                        }
                    }
                    else 
                    {
                        string M0 = xlWorksheet.Cells[18, h + 1].Value2.ToString();
                        string M1 = xlWorksheet.Cells[21, h + 1].Value2.ToString();
                        string M2 = xlWorksheet.Cells[24, h + 1].Value2.ToString();
                        string M3 = xlWorksheet.Cells[27, h + 1].Value2.ToString();
                        string[] data = { M0, M1, M2, M3 };
                        for (int d = 0; d < data.Length; d++)
                        {
                            rSheet.Cells[5 * h + 2 + d, 2] = data[d];
                        }
                    }

                    if (xlWorksheet.Cells[30, h + 1].Value2 == null)
                    {
                        string C1M0 = "";
                        string C2M0 = "";
                        string C3M0 = "";

                        string C1M1 = "";
                        string C2M1 = "";
                        string C3M1 = "";

                        string C1M2 = "";
                        string C2M2 = "";
                        string C3M2 = "";

                        string C1M3 = "";
                        string C2M3 = "";
                        string C3M3 = "";

                        string[] dataC1 = { C1M0, C1M1, C1M2, C1M3 };
                        string[] dataC2 = { C2M0, C2M1, C2M2, C2M3 };
                        string[] dataC3 = { C3M0, C3M1, C3M2, C3M3 };

                        for (int d = 0; d < dataC1.Length; d++)
                        {
                            rSheet.Cells[5 * h + 2 + d, 3].Value = dataC1[d];
                        }

                        for (int d = 0; d < dataC2.Length; d++)
                        {
                            rSheet.Cells[5 * h + 2 + d, 4].Value = dataC2[d];
                        }

                        for (int d = 0; d < dataC3.Length; d++)
                        {
                            rSheet.Cells[5 * h + 2 + d, 5].Value = dataC3[d];
                        }
                    }
                    else
                    {

                        string C1M0 = xlWorksheet.Cells[30, h + 1].Value2.ToString();
                        string C2M0 = xlWorksheet.Cells[31, h + 1].Value2.ToString();
                        string C3M0 = xlWorksheet.Cells[32, h + 1].Value2.ToString();

                        string C1M1 = xlWorksheet.Cells[35, h + 1].Value2.ToString();
                        string C2M1 = xlWorksheet.Cells[36, h + 1].Value2.ToString();
                        string C3M1 = xlWorksheet.Cells[37, h + 1].Value2.ToString();

                        string C1M2 = xlWorksheet.Cells[40, h + 1].Value2.ToString();
                        string C2M2 = xlWorksheet.Cells[41, h + 1].Value2.ToString();
                        string C3M2 = xlWorksheet.Cells[42, h + 1].Value2.ToString();

                        string C1M3 = xlWorksheet.Cells[45, h + 1].Value2.ToString();
                        string C2M3 = xlWorksheet.Cells[46, h + 1].Value2.ToString();
                        string C3M3 = xlWorksheet.Cells[47, h + 1].Value2.ToString();

                        string[] dataC1 = { C1M0, C1M1, C1M2, C1M3 };
                        string[] dataC2 = { C2M0, C2M1, C2M2, C2M3 };
                        string[] dataC3 = { C3M0, C3M1, C3M2, C3M3 };

                        for (int d = 0; d < dataC1.Length; d++)
                        {
                            rSheet.Cells[5 * h + 2 + d, 3].Value = dataC1[d];
                        }

                        for (int d = 0; d < dataC2.Length; d++)
                        {
                            rSheet.Cells[5 * h + 2 + d, 4].Value = dataC2[d];
                        }

                        for (int d = 0; d < dataC3.Length; d++)
                        {
                            rSheet.Cells[5 * h + 2 + d, 5].Value = dataC3[d];
                        }

                        c--;
                    }
                }


                rSheet.get_Range("A1", "E1").Font.Bold = true;
                rSheet.get_Range("A:A").Font.Bold = true;
                

                xlWorksheet.SaveAs(filesave, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing);
                xlWorkbook.Close();
                xlApp.Quit();

            }


            else
            {
                MessageBox.Show("Error: Select a spreadsheet to continue");
            }
        }

        //routine to detect arduino
        private string AutodetectArduinoPort()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(connectionScope, serialQuery);

            try
            {
                foreach (ManagementObject item in searcher.Get())
                {
                    string desc = item["Description"].ToString();
                    string deviceId = item["DeviceID"].ToString();

                    if (desc.Contains("Arduino"))
                    {
                        return deviceId;
                    }
                }
            }
            catch (ManagementException e)
            {
                /* Do Nothing */
            }

            return null;
        }

        //routine to allow measuring as soon as measuring # is changed
        private void inputcount_TextChanged(object sender, EventArgs e)
        {
            string value = inputcount.Text;
            measure.Enabled = true;

        }

        //routine to move progress bar as measuring occurs
        private void progressBar()
        {
            int number;
            string value = inputcount.Text;
            bool isParsable = Int32.TryParse(value, out number);

            if (isParsable)
            {
                progressBar1.Maximum = number;
                progressBar1.Minimum = 0;
                if (measureCount >= number)
                {
                    messageBar.Text = "Measurements Complete, click 'Done' to export data";
                    measure.Enabled = false;
                    progressBar1.Value = measureCount;
                    measureStatus.Text = (measureCount.ToString() + "/" + number.ToString());
                    //progressBar1.Value = 0;
                    //reload();

                }
                else
                {
                    progressBar1.Value = measureCount;
                    measureStatus.Text = (measureCount.ToString() + "/" + number.ToString());
                }



            }
            else
                number = 0;
        }

        //routine to restart whole measuring process
        private void reset_Click(object sender, EventArgs e)
        {
            port.Write("1");
            measureCount = 0;
            inputcount.Text = "0";
            progressBar();
            inputcount.Text = "Input # of Measurements";
            inputcount.Enabled = false;
            datafill.Text = "";
            measureStatus.Text = measureCount.ToString();
            measure.Enabled = false;
            toolConnect.Enabled = true;
        }

        //routine to separate string of data in filled in text boc
        private void delimeter()
        {
            int p;
            string phrase = datafill.Text;
            string[] words = phrase.Split(' ');
            int arrlength = words.Length;
            string sentence = ("");
            //datafill.Text = string.Join(" " , words );
            for (p = 0; p < arrlength; p++)
            {
                if (p % 2 != 0)
                {
                    sentence += words[p] + Environment.NewLine;
                }
                else
                {
                    sentence += words[p] + " ";
                }
            }
            datafill.Text = (sentence);
        }

        //routine to save excel sheet to specific location
        private string saver()
        {
            string thedate = date.Value.ToString("MM_dd_yyyy");
            string filemaker = (thedate + "_XRite.xlsx");
            string desktop = @"C:\Users\delisle\Desktop\HP\";
            string filesave = Path.Combine(desktop, filemaker);
            return filesave;
        }

        //routine similar to reset but doesn't include progress bar
        private void reload()
        {
            port.Write("1");
            measureCount = 0;
            inputcount.Text = "0";
            inputcount.Enabled = false;
            measureStatus.Text = measureCount.ToString() + "/0";
            measure.Enabled = false;
            toolConnect.Enabled = true;
        }

        //routine to finish measuring and start excel data processing
        private void done_Click(object sender, EventArgs e)
        {
            measureStatus.Text = measureCount.ToString();
            measure.Enabled = false;
            toolConnect.Enabled = true;
            string filePath = FileName();
            string filesave = saver();
            SS(filePath, filesave);
            progressBar1.Value = 0;
            progressBar();
            reload();
        }
    }
}