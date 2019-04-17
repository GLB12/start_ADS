using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ivi.Visa.Interop;

namespace WindowsFormsApplication1
{
    public class PLTS
    {
        private FormattedIO488 ioDmm = new FormattedIO488();
        private bool _connected = false;

        public void Connect(string visaAddress)
        {
            ResourceManager grm = new ResourceManager();
            ioDmm.IO = (IMessage)grm.Open(visaAddress, AccessMode.NO_LOCK, 0, "");
            _connected = true;
        }
        public bool Connected
        {
            get
            {
                return _connected;
            }
        }

        public void Disconnect()
        {
            ioDmm.IO.Close();
            _connected = false;
        }

        public void HideDisplay()
        {
            string cmd = "DISPlay:STATe HIDE";
            ioDmm.WriteString(cmd);

        }

        public void ShowDisplay()
        {
            string cmd = "DISPlay:STATe SHOW";
            ioDmm.WriteString(cmd);

        }

        public string ReadString()
        {
            return ioDmm.ReadString().TrimEnd('\n').Trim('\"');
        }

        private void OPC()
        {
            ioDmm.WriteString("*opc?");
            ReadString();
        }

        public void CloseAllFiles()
        {
            ioDmm.WriteString(":ACLS");
        }

        public void ImportSnpFile(string snpFile)
        {
            string cmd = ":IMPort:DOMain FREQuency";
            ioDmm.WriteString(cmd);

            cmd = ":IMPort:TYPE TSOne";
            ioDmm.WriteString(cmd);

            cmd = ":IMPort:RANGe ALL";
            ioDmm.WriteString(cmd);

            cmd = string.Format(":IMPort '{0}'", snpFile);
            ioDmm.WriteString(cmd);
        }

        public int GetActiveFileIndex()
        {
            string cmd = "File:Active?";
            ioDmm.WriteString(cmd);

            string result = ReadString();
            return int.Parse(result);
        }

        public void Do2XThruAFR(string thruFileName, string fixture1, string fixture2, string RefZ)
        {
            int timeout = ioDmm.IO.Timeout;
            ioDmm.IO.Timeout = 100000;

            string cmd = ":AFR:INIT";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:INPuts SENDed";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:MEASurement 2";
            ioDmm.WriteString(cmd);

            cmd = string.Format(":AFR:FIXTure:REFZ '{0}'",RefZ);
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:CLENgth OFF";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:CMATch ON";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:BLIMited OFF";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:STANdard:USE THRU,ON";
            ioDmm.WriteString(cmd);

            cmd = string.Format(":AFR:STANdard:LOAD THRU,'{0}'", thruFileName);
            ioDmm.WriteString(cmd);

            cmd = ":AFR:SAVE:TYPE TSONe";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:SAVE:PORTs PLTS";
            ioDmm.WriteString(cmd);

            cmd = string.Format(":AFR:SAVE:FILename '{0}','{1}'", fixture1, fixture2);
            ioDmm.WriteString(cmd);

            OPC();

            ioDmm.IO.Timeout = timeout;

        }

        public void DoTwoPortDembedding(int dutFileNumber, string fixture1File, string fixture2File)
        {
            int timeout = ioDmm.IO.Timeout;
            ioDmm.IO.Timeout = 100000;

            string cmd = string.Format(":FILE{0}:DEEMbed:FILename 1,'{1}'", dutFileNumber, fixture1File);
            ioDmm.WriteString(cmd);

            cmd = string.Format(":FILE{0}:DEEMbed:FILename 2,'{1}'", dutFileNumber, fixture2File);
            ioDmm.WriteString(cmd);

            cmd = string.Format("FILE{0}:DEEMbed:STATe ON", dutFileNumber);
            ioDmm.WriteString(cmd);

            OPC();

            ioDmm.IO.Timeout = timeout;
        }

        public void DoDifferentialThruAFR(string thruFileName, string fixture1, string fixture2, string RefZ)
        {
            int timeout = ioDmm.IO.Timeout;
            ioDmm.IO.Timeout = 100000;

            string cmd = ":AFR:INIT";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:INPuts DIFFerential";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:MEASurement 4";
            ioDmm.WriteString(cmd);

            cmd = string.Format(":AFR:FIXTure:REFZ '{0}'", RefZ);
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:CLENgth OFF";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:FIXTure:BLIMited OFF";
            ioDmm.WriteString(cmd);

            cmd = ":AFR:STANdard:USE THRU,ON";
            ioDmm.WriteString(cmd);

            cmd = string.Format(":AFR:STANdard:LOAD THRU,'{0}'", thruFileName);
            ioDmm.WriteString(cmd);

            cmd = ":AFR:SAVE:PORTs PLTS";
            ioDmm.WriteString(cmd);

            cmd = string.Format(":AFR:SAVE:FILename '{0}','{1}'", fixture1, fixture2);
            ioDmm.WriteString(cmd);

            OPC();

            ioDmm.IO.Timeout = timeout;
        }

        public void DoFourPortDeembedding(int dutFileNumber, string fixture1File, string fixture2File)
        {
            int timeout = ioDmm.IO.Timeout;
            ioDmm.IO.Timeout = 100000;

            string cmd = string.Format(":FILE{0}:DEEMbed:FILename 1,'{1}'", dutFileNumber, fixture1File);
            ioDmm.WriteString(cmd);

            cmd = string.Format(":FILE{0}:DEEMbed:FILename 2,'{1}'", dutFileNumber, fixture2File);
            ioDmm.WriteString(cmd);

            cmd = string.Format("FILE{0}:DEEMbed:STATe ON", dutFileNumber);
            ioDmm.WriteString(cmd);

            OPC();

            ioDmm.IO.Timeout = timeout;
        }


        public void ExportFile(int fileNumber, int portCnt, string fileName)
        {
            string cmd = ":EXPort:INITialize";
            ioDmm.WriteString(cmd);

            cmd = ":EXPort:FORMat MA";
            ioDmm.WriteString(cmd);

//            cmd = string.Format(":EXPort:PORTs:COUNt {0}", portCnt);
//            ioDmm.WriteString(cmd);

            cmd = ":EXPort:SDATa SENDed";
            ioDmm.WriteString(cmd);

            cmd = string.Format(":FILE{0}:EXPort '{1}'", fileNumber, fileName);
            ioDmm.WriteString(cmd);

        }

    }
}
