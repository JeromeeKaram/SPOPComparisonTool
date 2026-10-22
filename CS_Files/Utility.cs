using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace Cleaning_Records_Sheet
{
    public enum RUN
    {
        IHI = 0,
        CYIENT = 1
    }
    class Utility
    {
        public static RUN m_eRUN = RUN.IHI;
        static int myCnt = 0;
        public static string m_sTempPath = "";
        public static string m_sBinPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);

        public static string m_sToolName = "SPM Comparison and Cleaning Records Sheet Highlighting";
        public static string m_sVersion = "1.0";
        public static string m_sDate = "22-Oct-26";

        public static Dictionary<string, ColorInfo> m_dtColorCode = new Dictionary<string, ColorInfo>();

        //*******************************************************
        //Function  : GetTempDir
        //Purpose   : Get the temporaory directory
        //*******************************************************
        public static void GetTempDir()
        {
            //Get temp Directory
            string path = null;
            try
            {
                //Get the TEMP Path
                path = System.Environment.GetEnvironmentVariable("TEMP");
                if (path == null || path.Length < 1)//Not found
                {
                    //Get the TMP path
                    path = System.Environment.GetEnvironmentVariable("TMP");
                    if (path == null || path.Length < 1)//Not found
                    {
                        path = System.Windows.Forms.Application.ExecutablePath.ToString();
                    }
                }
                m_sTempPath = path + @"\";
            }
            catch { }
        }
        //*******************************************************
        //Function  : WriteErrorLog
        //Purpose   : Write the error data to log
        //*******************************************************
        public static void WriteErrorLog(string sWarning, string sStackTrace, string sMsg)
        {
            try
            {
                if (sWarning.Trim().Length > 0)
                    MessageBox.Show(sWarning.Trim());
                string sLogFilepath = m_sTempPath + "ErrorLog.dat";
                if (sLogFilepath != null && sLogFilepath.Length > 0)
                {
                    try
                    {
                        //Write the Error Log
                        StreamWriter sw;
                        if (myCnt == 0)//First Error
                        {
                            sw = System.IO.File.CreateText(sLogFilepath);
                            DateTime date = DateTime.Now;
                            sw.WriteLine("********** Engine Manual Revision-up Application *********");
                            sw.WriteLine("Version           : " + m_sVersion);
                            sw.WriteLine("Release Date      : " + m_sDate);
                            sw.WriteLine("Time              : " + date.ToString());
                            sw.WriteLine("**********************************************");
                        }
                        else//Next Errors Append
                        {
                            sw = System.IO.File.AppendText(sLogFilepath);
                        }
                        //----- This is for error
                        if (sWarning.Length > 0 || sStackTrace.Length > 0)
                        {
                            sw.WriteLine("-----------------");
                            if (sWarning.Length > 0)
                                sw.WriteLine("Warning     : " + sWarning);//Write the Error Message
                            if (sStackTrace.Length > 0)
                                sw.WriteLine("StackTrace  : " + sStackTrace);//Write the StackTrace
                            if (sStackTrace.Length > 0)
                                sw.WriteLine("Msg  : " + sMsg);//Write the StackTrace
                            sw.WriteLine("-----------------");
                        }
                        else if (sMsg.Length > 0)
                        {
                            sw.WriteLine(sMsg);
                        }
                        //close
                        sw.Close();
                        myCnt++;
                    }
                    catch
                    { }
                }
            }
            catch { }
        }
        //*******************************************************
        //Function  : WriteErrorLog
        //Purpose   : Write the error data to log
        //*******************************************************
        public static void WriteErrorLog(Exception ee)
        {
            try
            {
                if (ee != null)
                {
                    WriteErrorLog(ee.Message, ee.StackTrace, "");
                }
            }
            catch { }
        }
        public static void WriteErrorLog(string str)
        {
            try
            {
                WriteErrorLog("", "", str);
            }
            catch { }
        }
        //*******************************************************
        //Function  : WarnUser
        //Purpose   : Warn the user
        //*******************************************************
        public static void WarnUser(string smsg)
        {
            try
            {
                //show
                MessageBox.Show(smsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        //*******************************************************
        //Function  : InformationUser
        //Purpose   : information message box to user
        //*******************************************************
        public static void InformationUser(string smsg)
        {
            try
            {
                //show
                MessageBox.Show(smsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public static List<string> SplitString(string sValue, string schar)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(new string[1] { schar }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string ss in sArr)
                {
                    if (ss.Trim().Length > 0)
                    {
                        lst.Add(ss.Trim());
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static List<string> SplitString(string sValue, string[] schars, bool Toupper)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(schars, StringSplitOptions.RemoveEmptyEntries);
                foreach (string ss in sArr)
                {
                    if (ss.Trim().Length > 0)
                    {
                        if (Toupper == false)
                            lst.Add(ss.Trim());
                        else
                            lst.Add(ss.ToUpper().Trim());
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static void CopyFile1(string source, string dest)
        {
            try
            {
                Utility.DeleteFile(dest);
                System.IO.File.Copy(source, dest, true);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public static string CopyFiletoLocal(string sInputFileName)
        {
            string sNewInputFileName = "";
            try
            {
                if (System.IO.File.Exists(sInputFileName))
                {
                    string sFileName = System.IO.Path.GetFileNameWithoutExtension(sInputFileName);
                    string sFileExtn = System.IO.Path.GetExtension(sInputFileName);
                    //--copy the file...
                    bool bYes = false;
                    int nCnt = 0;
                    while (true)
                    {
                        nCnt = nCnt + 1;
                        sNewInputFileName = m_sTempPath + "\\" + sFileName + "_" + nCnt.ToString() + sFileExtn;
                        try
                        {
                            CopyFile1(sInputFileName, sNewInputFileName);
                            bYes = true;
                            break;
                        }
                        catch //(Exception ee)
                        {
                            // Utility.WriteErrorLog("", "", ee.Message);
                        }
                    }
                    //wait till file copied...
                    int ncnt = 0;
                    while (bYes)
                    {
                        System.Threading.Thread.Sleep(2000);
                        ncnt = ncnt + 1;
                        if (System.IO.File.Exists(sNewInputFileName))
                        {
                            break;
                        }
                        //wait for 5 min..
                        if (ncnt == 150)
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return sNewInputFileName;
        }
        public static void DeleteFile(string sFile)
        {
            try
            {
                if (System.IO.File.Exists(sFile) == true)
                {
                    System.IO.File.Delete(sFile);
                }
            }
            catch
            {
                // Utility.WriteErrorLog(ee);
            }
        }
        public static string GetURlData(string linkPath)
        {
            string contents = "";
            try
            {
                using (var client = new System.Net.WebClient())
                {
                    try
                    {
                        contents = client.DownloadString(linkPath);
                    }
                    catch (Exception ee)
                    {
                        Utility.WriteErrorLog("", linkPath, ee.Message + " **** " + ee.StackTrace);
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return contents;
        }
        public static int GetInteger(string str)
        {
            try
            {
                str = str.Trim();
                if (str.Length > 0)
                {
                    return int.Parse(str);
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return 0;
        }

        public static List<string> ReadFile(string sFile)
        {
            List<string> lstOp = new List<string>();
            try
            {
                if (System.IO.File.Exists(sFile))
                {
                    string[] sArr = System.IO.File.ReadAllLines(sFile);
                    foreach (string ss in sArr)
                    {
                        if (ss.Trim().Length > 0)
                        {
                            lstOp.Add(ss.Trim());
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lstOp;
        }
    }
}
