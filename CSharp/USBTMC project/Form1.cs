using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ivi.Visa.Interop;
using NationalInstruments.Visa;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Reflection;


namespace VISA_Example
{
    public partial class Form1 : Form
    {
        Ivi.Visa.Interop.ResourceManager rMgr = new ResourceManagerClass();//Create a resource manager
        FormattedIO488 src = new FormattedIO488Class(); //Create a new IEEE 488.2-like message based session 

        public Form1()
        {
            InitializeComponent();
            getAvailableResources(); // calls for the function GetAvailable resources 

        }

        public void getAvailableResources()//function find all available resources using the FindRsrc command and populates combobox1
        {
            try
            {
                string[] resources = rMgr.FindRsrc("?*");   //?* Matches all resources 

                comboBox1.Items.AddRange(resources);        //List all avaible resources in comboBox1
            }
            catch (Exception)
            {
                textBox2.Text = "No Resources Available";   //If no resources where found NO Resources Available will be displayed in the read terun Stings textBox2

            }


        }




        public string sendstring; // instanciating strings
        public string sendlogstring;

        public string FilePath{ get; }
        public string ID;
        public string resp;

        public void CreateExcel()
        {

            this.StartButton.Text = "Stop Logging";

            string execPath =
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);

            DateTime date1 = new DateTime();
            DateTime dateOnly = date1.Date;
            DateTime timeOnly = date1.ToLocalTime();
            string dateOnly1 = dateOnly.ToString("d");

            string FilePath = textBox7.Text; 



            //C:\Users\ARamirez.BK1\Desktop\VISA Example.xlsx
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = false;
            excel.DisplayAlerts = false;

            //Create a new workbook

            Workbook wb;
            Worksheet ws;
            

            wb = excel.Workbooks.Add(Type.Missing);
            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
            ws.Name = "Log";




            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created.");
                return;
            }

            ws.get_Range("A1", "C1").Merge();
            Range Date = ws.Range["A1:D1"];
            Date.Value = dateOnly;

            ws.get_Range("A2", "D2").Merge();
            Range IDN = ws.Range["A2:D2"];
            IDN.Value = ID;



            wb.SaveAs(FilePath);
            wb.Close();


        }

        public void WriteToExcel()
        {


            //C:\Users\ARamirez.BK1\Desktop\VISA Example.xlsx
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = false;
            excel.DisplayAlerts = false;

            //Create a new workbook

            Workbook wb;
            Worksheet ws;


            wb = excel.Workbooks.Open(Path.GetFileName(textBox7.Text));
            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;


            for (int i = 3; i < Convert.ToInt32(textBox5.Text); i = i + 1)

                try
                {
                resp = src.ReadString() + "\r\n";                             // read instrument buffer and append return string with string in textBox2

                ws.get_Range("Ai", "Di").Merge();
                Range RES = ws.Range["Ai:Di"];
                RES.Value = resp;

                wb.SaveAs(FilePath);
                wb.Close();

                }
                catch (TimeoutException)
                {
                textBox2.Text = "timeout exception";                                    // tineout provided (12s) has expired (nothing was returned)

                }




        }





        public void Form1_Load(object sender, EventArgs e)      //disables all groupboxes upon loading Form1
        {
            groupBox1.Enabled = false;  //Disables the Send String box
            groupBox2.Enabled = false;  //Disables the Read Return Strings box
            groupBox3.Enabled = false;  //Disables the LOGGING.CSV box
            
        }

       public void comboBox1_SelectedIndexChanged(object sender, EventArgs e)   //action after clicking the Available Resources drop down menu
        {
            label1.Text = comboBox1.Text;
            button3.Enabled = true;         //Enables the open session button
            button4.Enabled = true;         //Enables the close session button


        }

   
        public void button1_Click(object sender, EventArgs e)       //action after clicking the Write Button
        {

            sendstring = textBox1.Text;                                                 // assigns the command to sendstring variable
            src.WriteString(sendstring + "\n");                                         // appends linefeed terminataion character to command and sends it            
            textBox2.AppendText(textBox1.Text + "\r\n");                                // appends the command to the return string box (read) to indicate what the write commadn was
            textBox1.Text = "";                                                         // sets the text in textBox 1 to blank 
            //  textBox2.Text = src.ReadString();

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {


            

        }

        public void button3_Click(object sender, EventArgs e)       //action after clicking the the OPEN button
        {
            string srcAddress = label1.Text;

            src.IO = (IMessage)rMgr.Open(srcAddress, AccessMode.NO_LOCK, 2000, "");     // Open a session
            src.IO.Timeout = 10000;                                                     // sets a 12 second timeout

            comboBox1.Enabled = false;                  //disables the availble resource box once a resource is selected
            groupBox1.Enabled = true;                   //enables the Send String box once a resource is selected
            groupBox2.Enabled = true;                   //enables the Read Returned box once a resource is selected
            groupBox3.Enabled = true;                   //enables the LOGGING.CSV box once a resourece is selected

            src.WriteString("*IDN?\n");                         //write *IDN 
                try
                {
                textBox8.Text += src.ReadString() + "\r\n";     // read instrument buffer and return string to Instrument ID textbox
                }
                catch (TimeoutException)
                {
                textBox8.Text = "timeout exception";            // tineout provided (10s) has expired (nothing was returned)

                }

        }

        public void textBox2_TextChanged(object sender, EventArgs e)
        {
            
                  

            
        }
        

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        public void button2_Click_1(object sender, EventArgs e)                         // action after clicking the Read button
        {

            try
            {
                textBox2.Text += src.ReadString() + "\r\n";                             // read instrument buffer and append return string with string in textBox2
            }
            catch (TimeoutException)
            {
                textBox2.Text = "timeout exception";                                    // tineout provided (12s) has expired (nothing was returned)

            }
                        
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            
            
        }

        private void button4_Click(object sender, EventArgs e)      // action after clicking the CLOSE button
        {
            Close();                                                // closes the application 
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)    // action of pressing the enter key when in textbox1 will click the Write button
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(this, new EventArgs());
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            label3.Text = textBox3.Text;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            label4.Text = textBox4.Text;
        }

        private async void button5_Click(object sender, EventArgs e)        //stop button
        {

            src.WriteString("*IDN?\n");
            string ID = src.ReadString();

            label8.Text = "";
           
            for (int a = 0; a < Convert.ToInt32(textBox5.Text); a = a + 1)
            {
                
                
                    

                    try
                    {
                        sendlogstring = label3.Text + ";" + label4.Text;


                        src.WriteString("MEAS:ALL?\n");

                        

                        string fullstring = src.ReadString(); //textBox2.Text;
                        var onestring = fullstring.Split(',');
                        label6.Text = onestring[0];
                        label7.Text = onestring[1];
                        textBox2.Text += fullstring + "\r\n";
                    
                        // CSV LOGGING
                        StringBuilder csvcontent = new StringBuilder();
                    
                        csvcontent.AppendLine(DateTime.Now.ToString("HH:mm:ss:fff") + "," + label3.Text + "," + label6.Text + "," + label4.Text + "," + label7.Text);
                    // string csvpath = "C:\\Users\\aramirez\\Documents\\Current Projects\\trial_juerves.csv";
                    string csvpath = textBox7.Text;

                    File.AppendAllText(csvpath, csvcontent.ToString());
                    // csvcontent.AppendLine("alex, 895");
                        int b = Convert.ToInt32(textBox6.Text);
                        await Task.Delay(b*1000);
                    
                }

                 
                    catch (TimeoutException)
                    {
                        textBox2.Text = "timeout exception";

                    }
                

            }
            label8.Text = "Logging Completed";

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private CancellationTokenSource _canceller;

        private async void StartButton_Click(object sender, EventArgs e)
        {
            StartButton.Enabled = false;
            StopButton.Enabled = true;

            src.WriteString("*IDN?\n");
            ID = src.ReadString();
            CreateExcel();

            _canceller = new CancellationTokenSource();
            await Task.Run(() =>
            {
                do
                {
                    WriteToExcel();




                    if (_canceller.Token.IsCancellationRequested)
                        break;

                } while (true);
            });

            _canceller.Dispose();
            StartButton.Enabled = true;
            StopButton.Enabled = false;
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            _canceller.Cancel();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}

