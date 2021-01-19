using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScreenerIn_Automation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            
            InitializeComponent();
            linkLabel2.Text = $@"C:\Users\{Environment.UserName}\Downloads";

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            BusinessExecution businessExecution = new BusinessExecution();
            businessExecution.Setup();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog();
            string file = openFileDialog1.FileName;
            textBox1.Text = file;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel2.Visible = false;
            label5.Visible = false;
            textBox1.Visible = false;
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            textBox2.Text = Properties.Settings.Default.UserName;
            textBox3.Text = Properties.Settings.Default.Password;
            groupBox1.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if(button5.Text=="Edit")
            {
                textBox2.Enabled = true;
                textBox3.Enabled = true;
                button5.Text = "Save";
            }
            else
            {
                Properties.Settings.Default.UserName = textBox2.Text;
                Properties.Settings.Default.Password = textBox3.Text;
                Properties.Settings.Default.Save();
                Application.Restart();
            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

       /* private void backgroundWorker7_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(Properties.Settings.Default.UserName) && !string.IsNullOrEmpty(Properties.Settings.Default.Password))
                {
                    BusinessExecution execution = new BusinessExecution();
                    List<string> company = execution.MainExecution(textBox1.Text);
                    double counter = Math.Round(100.00 / company.Count, 2);
                    double inital = counter;
                    foreach (string comp in company)
                    {
                        SetText("Getting Data from " + comp);
                        execution.ExecuteQuerysteps(comp);
                        setprogval(Convert.ToInt32(inital));
                        inital += counter;
                    }
                    execution.Final();
                }
                else
                    MessageBox.Show("UserName and password is empty. Please configure in the configuration setting");
            }
            catch(Exception ea)
            {
                MessageBox.Show("Unexcepted Error : " + ea);
            }
            
            
        }*/

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(Properties.Settings.Default.UserName) && !string.IsNullOrEmpty(Properties.Settings.Default.Password))
                {
                    ExcelModel excelModel = new ExcelModel();
                    ApiCalls apiCalls = new ApiCalls();
                    logindata initial = apiCalls.login(Properties.Settings.Default.UserName.ToString(), Properties.Settings.Default.Password.ToString());
                    apiCalls.readExcelFile(textBox1.Text, excelModel);
                    //List<string> comapnyurl = apiCalls.readcomapnyurl(textBox1.Text);
                    double counter = Math.Round(100.00 / excelModel.comapny.Count, 2);
                    double inital = counter;
                    for(int i=0;i<excelModel.comapny.Count;i++)
                    {
                        SetText("Getting Data from " + excelModel.comapny[i]);
                        apiCalls.MainExecution(excelModel.comapny[i], initial, excelModel.url[i]);
                        setprogval(Convert.ToInt32(inital));
                        inital += counter;
                    }
                    apiCalls.Final();
                    MessageBox.Show("Completed", "Completed");
                }
                else
                    MessageBox.Show("UserName and password is empty. Please configure in the configuration setting");
            }
            catch (Exception ea)
            {
                MessageBox.Show("Unexcepted Error : " + ea);
            }


        }

        delegate void SetTextCallback(string text);

        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.textBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                label4.Text = text;
            }
        }

        delegate void setProgressvalue(int val);

        private void setprogval(int val)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.progressBar1.InvokeRequired)
            {
                setProgressvalue d = new setProgressvalue(setprogval);
                this.Invoke(d, new object[] { val });
            }
            else
            {
                progressBar1.Value = val;
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start($@"C:\Users\{Environment.UserName}\Downloads");
        }
    }
}
