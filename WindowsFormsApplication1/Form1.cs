using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public static Process PLTS_pro;
        PLTS plts = new PLTS();

        //声明PAI函数，便于将第三方软件最小化
        [DllImport("kernel32.dll", EntryPoint = "WinExec")]
        public static extern int WinExec(string lpCmdLine, int nCmdShow);

        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = OpenFiles();

        }

        private string OpenFiles()
        {
            string _filename = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Touchstone files(*.s*p)|*.s*p";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                _filename = ofd.FileName;

            }
            return _filename;

        }

        private string OpenDialog()
        {
            string _DialogPath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            if(fbd.ShowDialog() ==DialogResult.OK)
            {
                _DialogPath = fbd.SelectedPath;

            }
            return _DialogPath;

        }
        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = OpenFiles();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox3.Text = OpenDialog();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox4.Text = OpenDialog();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            ProcessStartInfo Info = new ProcessStartInfo();
            Info.FileName = "C:\\Program Files (x86)\\Keysight\\PLTS2018\\PLTS.exe";
            Info.Arguments = "";
            Info.WorkingDirectory = "";
            Info.WindowStyle = ProcessWindowStyle.Minimized;

            Process[] pProcess;
            pProcess = Process.GetProcesses();
            for (int i=1;i<=pProcess.Length-1;i++)
            {
                if(pProcess[i].ProcessName=="PLTS")
                {
                    PLTS_pro = pProcess[i];
                    break;
                }

            }

            if (PLTS_pro == null)
            {
                //PLTS_pro = new System.Diagnostics.Process();
                //PLTS_pro.StartInfo.FileName = "C:\\Program Files (x86)\\Keysight\\PLTS2018\\PLTS.exe";
                //PLTS_pro = Process.Start(Info);
                WinExec(Info.FileName, 0);
                IntPtr hand = FindWindow(null, "Physical Layer Test System");
                ShowWindow(hand, 2);
                System.Threading.Thread.Sleep(5000);
                //pProcess = Process.GetProcesses();
                //for (int i = 1; i <= pProcess.Length - 1; i++)
                //{
                //    if (pProcess[i].ProcessName == "PLTS")
                //    {
                //        break;
                //    }

                //}

            }

            //PLTS plts = new PLTS();
            plts.Connect("TCPIP0::localhost::hislip0::INSTR");
            if (plts.Connected)
            {
                MessageBox.Show("连接成功");
                //plts.HideDisplay();

            } 
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            plts.ShowDisplay();

        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            if (textBox1.Text!=null & textBox2.Text!=null & textBox3.Text!=null & textBox4.Text!=null)
            {
                try
                {
                    string fixture1File, fixture2File; 
                    string DUT_filename = System.IO.Path.GetFileNameWithoutExtension(textBox1.Text);
                    string Thru_filename = System.IO.Path.GetFileNameWithoutExtension(textBox2.Text);
                    string suffix;
                    StringBuilder sBuilder = new StringBuilder();
                    if(DUT_filename[0]!=Thru_filename[0])
                    {
                        MessageBox.Show("DUT和Thru不匹配！");
              
                    }
                    else
                    {
                        sBuilder.Append(DUT_filename[0]);
                        for (int i = 1; i < DUT_filename.Length; i++)
                        {
                            if (DUT_filename[i] == Thru_filename[i])
                            {
                                sBuilder.Append(DUT_filename[i]);
                            }
                            else continue;
                        }
                        suffix = sBuilder.ToString();

                        string fileExtension = System.IO.Path.GetExtension(textBox1.Text);
                        int PortsNum=0;
                        if(fileExtension==".s2p")
                        {
                            PortsNum = 2;
                        }
                        else
                        {
                            if(fileExtension==".s4p")
                            {
                                PortsNum = 4;
                            }
                        }

                        int fileIndex;
                        string deembeddedDUTFile;

                        if (PortsNum!=0)
                        {
                            if (PortsNum == 2)
                            {
                                fixture1File = textBox3.Text + "\\fixture1.s2p";
                                fixture2File = textBox3.Text + "\\fixture2.s2p";
                                plts.Do2XThruAFR(textBox2.Text, fixture1File, fixture2File, "SYSZ");
                                plts.ImportSnpFile(textBox1.Text);
                                fileIndex = plts.GetActiveFileIndex();
                                plts.DoTwoPortDembedding(fileIndex, fixture1File, fixture2File);
                                deembeddedDUTFile = textBox4.Text + "\\" + suffix + "_AFR_2.s2p";
                                plts.ExportFile(fileIndex, 2, deembeddedDUTFile);

                            }
                            if (PortsNum == 4)
                            {
                                fixture1File = textBox3.Text + "\\fixture1.s4p";
                                fixture2File = textBox3.Text + "\\fixture2.s4p";
                                plts.DoDifferentialThruAFR(textBox2.Text, fixture1File, fixture2File, "SYSZ");
                                plts.ImportSnpFile(textBox1.Text);
                                fileIndex = plts.GetActiveFileIndex();
                                plts.DoFourPortDeembedding(fileIndex, fixture1File, fixture2File);
                                deembeddedDUTFile = textBox4.Text + "\\" + suffix + "_AFR_2.s4p";
                                plts.ExportFile(fileIndex, 4, deembeddedDUTFile);

                            }
                        }

                        plts.CloseAllFiles();

                    }

                }
                catch
                {
                    MessageBox.Show(e.ToString());

                }
            }

            this.Enabled = true;

        }
    }
}
