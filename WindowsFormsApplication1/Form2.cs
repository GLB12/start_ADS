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
using System.IO;


namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        public static Process ADS_pro;

        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProcessStartInfo Info = new ProcessStartInfo();
            Info.FileName = @"C:\Program Files\Keysight\ADS2020\bin\ads.exe";
            Info.Arguments = "";
            Info.WorkingDirectory = @"C:\Work\ADS_Proj";
            Info.WindowStyle = ProcessWindowStyle.Minimized;

            Process[] pProcess;
            pProcess = Process.GetProcesses();
            for (int i = 1; i <= pProcess.Length - 1; i++)
            {
                if (pProcess[i].ProcessName == "hpeesofde")
                {
                    pProcess[i].Kill();
                }

                if(pProcess[i].ProcessName == "hpeesofdds")
                {
                    pProcess[i].Kill();

                }
            }

            ADS_pro = new System.Diagnostics.Process();
            ADS_pro = Process.Start(Info);

            MessageBox.Show("启动完毕！");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Process proc = null;
            try
            {
                string targetDir = string.Format(@"C:\Work\ADS_Proj\AEL_command");//这是bat存放的目录
                //string targetDir1 = AppDomain.CurrentDomain.BaseDirectory; //或者这样写，获取程序目录
                //proc = new Process();
                //proc.StartInfo.WorkingDirectory = targetDir;
                //proc.StartInfo.FileName = "ADS_env2020.bat";//bat文件名称
                //proc.StartInfo.Arguments = string.Format("10");//this is argument
                ////proc.StartInfo.CreateNoWindow = true;
                ////proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;//这里设置DOS窗口不显示，经实践可行
                //proc.Start();
                //proc.WaitForExit();

                //proc.StartInfo.FileName = "AEL_Run.bat";
                //proc.StartInfo.Arguments = string.Format("10");//this is argument
                //proc.Start();
                //proc.WaitForExit();

                string _AEL_command = "hpeesofemx hpeesofde -env de_sim  -m dem_temp.ael";
                Process p = new Process();
                p.StartInfo.FileName = "cmd.exe";
                p.StartInfo.UseShellExecute = false;    //是否使用操作系统shell启动
                p.StartInfo.RedirectStandardInput = true;//接受来自调用程序的输入信息
                p.StartInfo.RedirectStandardOutput = true;//由调用程序获取输出信息
                p.StartInfo.RedirectStandardError = true;//重定向标准错误输出
                p.StartInfo.CreateNoWindow = true;//不显示程序窗口
                p.StartInfo.WorkingDirectory = targetDir;//Setting the Dir
                p.Start();//启动程序
                p.StandardInput.WriteLine("ADS_env2020.bat");
                p.StandardInput.WriteLine(_AEL_command + "&exit");
                p.StandardInput.AutoFlush = true;
                //p.WaitForExit();
                p.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //Edit Variables
        public static void Edit_Variables(string _AELfilename, string _NewAELfile, string[] Varname, string[] Val)
        {
            int Varnum = Varname.Length;
            string _Var,_Val;
            if (File.Exists(_NewAELfile))
                File.Delete(_NewAELfile);

            try
            {
                //int i = 0;
                StreamWriter sw = new StreamWriter(_NewAELfile);

                using (StreamReader sr = new StreamReader(_AELfilename))
                {
                    string line= sr.ReadLine();
                    while(line!=null)
                    {
                        for(int i=0; i<=Varnum-1;i++)
                        {
                            _Var= "_"+Varname[i];
                            if (line.Contains(_Var))
                            {
                                _Val = Val[i];
                                line = line.Replace(_Var, _Val);

                            }
                            
                        }
                        sw.WriteLine(line);
                        line = sr.ReadLine();

                    }
                }
                sw.Close();
            }
            catch
            {}
        }


        private void button5_Click(object sender, EventArgs e)
        {
            string _AELSource = @"C:\Work\ADS_Proj\AEL_command\dem_new.ael";
            string _AELCurrent = @"C:\Work\ADS_Proj\AEL_command\dem_temp.ael";

            string[] Var = { "Rz_drum", "Rz_matte","tuneMin","tuneMax","tuneStep","OptMin","OptMax" };
            string[] Val = { "0.265", "1.725","2.2","4.0","0.01","2.4","3.8"};
            
            Edit_Variables(_AELSource, _AELCurrent, Var, Val);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            ADS_setting ads_set = new ADS_setting();
            ads_set.initializeADS();

            string _netlistSource = @".\Template\netlist.log";
            string _netlistNew = @".\netlist.log";
            string snp = @"C:\Users\leiboguo\source\repos\startADS\WindowsFormsApplication1\bin\Debug\PCB_Deemb_wrk\data\WUS_2_B_L3_Deemb_2.s2p";
            string filename = "\"" + snp + "\"";
            string[] Var = {"filename","er","Rz_drum" };
            string[] Val = { filename, "3.14111", "1.184" };
            Edit_Variables(_netlistSource, _netlistNew, Var, Val);

            richTextBox1.Text = ads_set.Simulate();
            ads_set.ShowDDS("B3_R2_v2.dds");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process[] temp = Process.GetProcessesByName("hpeesofdds");
            foreach(Process i in temp)
            {
                i.Kill();
            }
            temp = Process.GetProcessesByName("hpeesofde");
            foreach(Process i in temp)
            { i.Kill(); }
        }
    }
}
