using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;

namespace WindowsFormsApplication1
{
    class ADS_setting
    {
        public void initializeADS()
        {
            string[] lines;
            string install_path;
            try
            {
                string current_directory = System.Reflection.Assembly.GetEntryAssembly().Location;
                Directory.SetCurrentDirectory(Path.GetDirectoryName(current_directory));
                lines = System.IO.File.ReadAllLines("./config.log");
                install_path = lines[0];
            }
            catch
            {
                File.WriteAllText("config.log", @"C:\Program Files\Keysight\ADS2020");
                lines = System.IO.File.ReadAllLines("config.log");
                install_path = lines[0];
            }
            if (!File.Exists(install_path + @"\bin\hpeesofsim.exe"))
            {
                MessageBox.Show("Please set correct install path in config.log and restart ADS SOP.", "Wrong ADS Installation Path", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Process.Start("notepad.exe", "config.log");
                throw new Exception();
            }

            Environment.SetEnvironmentVariable("HPEESOF_DIR", install_path);
            Environment.SetEnvironmentVariable("SIMARCH", "win32_64");
            string s0 = install_path + @"\bin;";
            string s1 = install_path + @"\circuit\lib.win32_64;";
            string s2 = install_path + @"\adsptolemy\lib.win32_64;";
            string s3 = install_path + @"\SystemVue\2016.08\win32_64\bin";
            Environment.SetEnvironmentVariable("PATH", s0 + s1+s2+s3);

        }

        public Process HiddenProcess()
        {
            Process prc = new Process();
            prc.StartInfo.RedirectStandardOutput = true;
            prc.StartInfo.UseShellExecute = false;
            prc.StartInfo.CreateNoWindow = true;
            prc.StartInfo.ErrorDialog = false;
            prc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            return prc;

        }

        public void Generate_netlist()
        {

        }
        public string Simulate()
        {
            String Simlog = null;
            Process prc = HiddenProcess();
            prc.StartInfo.FileName = "hpeesofsim";
            prc.StartInfo.Arguments = "netlist.log";

            prc.Start();
            while (!prc.StandardOutput.EndOfStream)
            {
                Simlog = Simlog + prc.StandardOutput.ReadLine() + "\r\n";
            }
            prc.WaitForExit();

            if (!System.IO.Directory.Exists(@".\data"))
            {
                System.IO.Directory.CreateDirectory(@".\data");
            }

            DirectoryInfo source = new DirectoryInfo(@".\data\");

            //ADX to DS Conversion
            //FileInfo[] adx_files = source.GetFiles("*.adx");
            //foreach (FileInfo i in adx_files)
            //{
            //    string name = Path.GetFileNameWithoutExtension(i.ToString());
            //    prc.StartInfo.FileName = "adx_to_ds";
            //    string argument = String.Format("{0}.adx {0}.ds", name);
            //    prc.StartInfo.Arguments = argument;
            //    prc.Start();
            //    prc.WaitForExit();
            //    File.Delete(name + ".adx");
            //    File.Delete(name + ".log");
            //}

            //Move dataset
            FileInfo[] ds_files = source.GetFiles("*.ds");
            if (ds_files.Length == 0)
            {
                throw new Exception();
            }
            foreach (FileInfo i in ds_files)
            {
                File.Copy(i.ToString(), @".\data\" + i.ToString(), true);
                File.Delete(i.ToString());
            }
            
            return Simlog;
        }

        public void ShowDDS(string dds_path)
        {
            Process[] temp = Process.GetProcessesByName("hpeesofdds");
            if (temp.Length == 0)
            {
                Process prc = HiddenProcess();
                prc.StartInfo.FileName = "dds";
                prc.StartInfo.Arguments = dds_path;

                prc.Start();
            }
        }
    }
}
