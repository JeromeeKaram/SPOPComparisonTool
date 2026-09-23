using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SVC
{
    class Utility
    {
        static int myCnt = 0;

        public static string m_sLogFilePath = "";
        public static string m_sTempPath = "";
        public static string m_sToolName = "SPOP Cleaning Record Comparison Tool";
        public static string m_sVersion = "1.0";
        public static string m_sDate = "28-September-26";
        public static string m_sBinPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);

        public static Dictionary<string, string> m_dtReplaceTexts = new Dictionary<string, string>();
        //public static StringComparison.StringComparison m_oStrCompareObj = null;
        public static Dictionary<string, ColorInfo> m_dtColorCode = new Dictionary<string, ColorInfo>();
        public static System.Drawing.Color m_oWaringNoteClr = System.Drawing.Color.FromArgb(237, 237, 237);
        //public static List<TableConfiguraton> m_lstTableconfig = new List<TableConfiguraton>();
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
        public static void WriteErrorLog(string sWarning, string sStackTrace, string sMsg)
        {
            try
            {
                if (sWarning.Trim().Length > 0)
                    MessageBox.Show(sWarning.Trim());
                string sLogFilepath = m_sLogFilePath + "ErrorLog.dat";
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
                            sw.WriteLine("********** " + m_sToolName + " *********");
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

        public static void WriteErrorLog(string err)
        {
            try
            {
                WriteErrorLog("", "", err);
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
        public static int GetInteger(string sValue)
        {
            int cnt = 0;
            try
            {
                sValue = sValue.Trim();
                if (sValue.Length > 0)
                {
                    int n1 = 0;
                    if (int.TryParse(sValue, out n1))
                    {
                        cnt = n1;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return cnt;
        }
        public static bool IsPositiveInteger(string sValue)
        {

            try
            {
                sValue = sValue.Trim();
                if (sValue.Length > 0)
                {
                    int cnt = 0;
                    try
                    {
                        int n1 = 0;
                        if (int.TryParse(sValue, out n1))
                        {
                            cnt = n1;
                        }
                    }
                    catch { }
                    if (cnt > 0) return true;
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
        public static bool IsReal(string sValue)
        {
            try
            {
                sValue = sValue.Trim();
                if (sValue.Length > 0)
                {
                    double d1 = 0.0;
                    if (double.TryParse(sValue, out d1))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
        public static List<string> SplitString2(string sValue, char schar)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(schar);
                foreach (string ss in sArr)
                {
                    lst.Add(ss.Trim());
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static double GetReal(string sValue)
        {
            double cnt = 0;
            try
            {
                sValue = sValue.Trim();
                if (sValue.Length > 0)
                {
                    double n1 = 0;
                    if (double.TryParse(sValue, out n1))
                    {
                        cnt = n1;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return cnt;
        }
        public static string SubString(string str, string substr)
        {
            try
            {
                if (str.Length > 0 && substr.Length > 0)
                {
                    return str.Substring(substr.Length, str.Length - substr.Length).Trim();
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return str;
        }
        public static bool GetInteger(string sValue, ref int nVal)
        {
            try
            {
                sValue = sValue.Trim();
                if (sValue.Length > 0)
                {
                    if (int.TryParse(sValue, out nVal))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
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
        public static string RemoveUnicodeAtStart(string strval)
        {
            try
            {
                int firstCharCode = strval[0];
                if (firstCharCode >= 9312 && firstCharCode <= 9331)
                {
                    return strval.Substring(1, strval.Length - 1).Trim();
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return strval;
        }
        public static List<string> SplitStringByKeepingSplitString(string sValue, string[] schars)
        {
            List<string> lstx = new List<string>() { sValue };
            try
            {
                foreach (string chars in schars)
                {
                    List<string> lstAll = new List<string>();
                    foreach (string s in lstx)
                    {
                        int n1 = s.IndexOf(chars);
                        if (n1 > 0)
                        {
                            string s1 = s.Substring(0, n1);
                            string s2 = s.Substring(n1, s.Length - n1);
                            lstAll.Add(s1);
                            lstAll.Add(s2);
                        }
                        else
                        {
                            lstAll.Add(s);
                        }
                    }
                    lstx.Clear();
                    foreach (string s in lstAll) lstx.Add(s);
                }


                ////Create pattern
                //string pattern = @"";
                //foreach (string s in schars)
                //    pattern = pattern + "(?<=\b" + s + "\b)|(?=\b" + s + "\b)|";
                //pattern = pattern.Substring(0, pattern.Length - 1);
                ////Split
                //string[] parts = System.Text.RegularExpressions.Regex.Split(sValue, pattern, RegexOptions.IgnoreCase);
                //foreach (string part in parts)
                //{
                //    if (!string.IsNullOrWhiteSpace(part))
                //    {
                //        lst.Add(($"[{part.Trim()}]"));
                //    }
                //}
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lstx;
        }
        public static string FindHtmlText(string strp, bool considerOnlyText)
        {
            try
            {
                string str = strp;
                if (str.StartsWith(">") == false) str = ">" + str;
                if (str.EndsWith("<") == false) str = str + "<";
                string strnew = "";
                for (int i = 0; i < str.Length; i++)
                {
                    if (str[i] == '>')
                    {

                        string sx = "";
                        for (int j = i + 1; j < str.Length; j++)
                        {
                            if (str[j] == '<')
                            {
                                i = j - 1;
                                break;
                            }
                            sx = sx + str[j].ToString();
                        }
                        if (sx.Length > 0) strnew = strnew + " " + sx;
                    }
                }
                //If user want only text and no tags...if empty also return
                if (considerOnlyText)
                {
                    return strnew.Trim();
                }


                if (strnew.Length > 0) return strnew.Trim();
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return strp;
        }
        public static List<string> DiffWords(List<string> L1, List<string> L2)
        {
            List<string> lst = new List<string>();
            try
            {
                foreach (string ss in L1)
                {
                    if (L2.Contains(ss) == false)
                    {
                        if (ss.Trim().Length > 0)
                        {
                            lst.Add(ss.Trim());
                        }
                    }
                }

            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static string CombineString(List<string> lst, string schar)
        {
            string sval = "";
            try
            {
                foreach (string ss in lst)
                {
                    sval = sval + ss + schar;
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return sval.Trim();
        }
        public static List<string> SplitString(string sValue, string[] schars)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(schars, StringSplitOptions.RemoveEmptyEntries);
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
        public static string cleanString(string str)
        {
            try
            {
                string s1 = str.Replace(((char)8211).ToString(), ((char)45).ToString());
                string s2 = s1.Replace("”", "\"");
                return s2.Trim();
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return str.Trim();
        }
        public static void releaseExcelObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ee)
            {
                obj = null;
                Utility.WriteErrorLog(ee);
            }
            finally
            {
                GC.Collect();
            }
        }

        internal static List<string> Split(string v1, string v2)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = v1.Split(new string[1] { v2 }, StringSplitOptions.None);
                foreach (string ss in sArr)
                {
                    //if (ss.Trim().Length > 0)
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
        public static string NoSpaceWithUpper(string str)
        {
            string strnew = str.Replace(" ", "").ToUpper().Trim();
            if (strnew.EndsWith(".")) strnew = strnew.Substring(0, strnew.Length - 1);
            return strnew;
        }
        public static void doit()
        {
            int cnt = 0;
            while (cnt < 6)
            {
                System.Threading.Thread.Sleep(10000);
                //--Getting tickets
                int nticket = 0;
                //
                if (nticket > 0) break;
            }
        }
        public static List<string> SplitStringWithoutTrim(string sValue, string schar)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(new string[1] { schar }, StringSplitOptions.None);
                foreach (string ss in sArr)
                {
                    lst.Add(ss);
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static string GetBullet(int index, int pos)
        {
            string bullet = "";
            try
            {
                int n1 = pos / 26;
                int n2 = pos % 26;
                if (n2 == 0)
                {
                    if (n1 > 1)
                    {
                        bullet = ((char)(index + n1 - 1)).ToString() + ((char)(index + 26)).ToString();
                    }
                    else
                    {
                        bullet = ((char)(index + 26)).ToString();
                    }
                }
                else
                {
                    if (n1 > 0)
                    {
                        bullet = bullet = ((char)(index + n1)).ToString() + ((char)(index + n2)).ToString();
                    }
                    else
                    {
                        bullet = ((char)(index + n2)).ToString();
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return bullet;
        }
        public static string ReplaceNonASCIIChars(string val)
        {
            try
            {
                string sx = System.Text.RegularExpressions.Regex.Replace(val, "[^\u0000-\u007F]", "");
                sx = sx.Replace("&lt;A&gt;", "<A>");
                return sx;
                //return System.Text.RegularExpressions.Regex.Replace(val, "[^\u0000-\u007F]", "");
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return val;
        }
        public static string FindHtmlText(string str)
        {
            try
            {
                string strnew = "";
                for (int i = 0; i < str.Length; i++)
                {
                    if (str[i] == '>')
                    {

                        string sx = "";
                        for (int j = i + 1; j < str.Length; j++)
                        {
                            if (str[j] == '<') break;
                            sx = sx + str[j].ToString();
                        }
                        if (sx.Length > 0) strnew = strnew + " " + sx;
                    }
                }
                return strnew.Trim();
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return str;
        }
        public static bool Comparelist(List<string> lnew, List<string> lold, bool bCaseSensitive)
        {
            try
            {
                if (lnew != null && lold != null)
                {
                    if (lnew.Count == lold.Count)
                    {
                        List<string> l1 = new List<string>();
                        List<string> l2 = new List<string>();
                        if (bCaseSensitive == false)
                        {
                            foreach (string s in lnew) l1.Add(s);
                            foreach (string s in lold) l2.Add(s);
                        }
                        else
                        {
                            foreach (string s in lnew) l1.Add(s.ToUpper());
                            foreach (string s in lold) l2.Add(s.ToUpper());
                        }
                        //Now compare
                        for (int i = 0; i < l1.Count; i++)
                        {
                            int n = l2.IndexOf(l1[i]);
                            if (n >= 0)
                            {
                                l2.RemoveAt(n);
                                l1.RemoveAt(i);
                                i--;
                            }
                        }
                        if (l1.Count == 0 && l2.Count == 0) return true;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
        public static List<string> SplitStringAddItChar(string sValue, string schar)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(new string[1] { schar }, StringSplitOptions.None);
                //foreach (string ss in sArr)
                for (int x = 0; x < sArr.Length; x++)
                {
                    if (x == 0) lst.Add(sArr[x]);
                    else lst.Add(schar + sArr[x]);
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static bool IsDeletedProcedurePoint(string str)
        {
            try
            {
                str = str.Replace(".", "").ToUpper().Trim();
                if (str == "Deleted".ToUpper()) return true;
                if (str == "Delete".ToUpper()) return true;
                if (str == "(Deleted)".ToUpper()) return true;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
        public static List<string> cleanList(List<string> lst)
        {
            try
            {
                List<string> lstNew = new List<string>();
                foreach (string ss in lst)
                {
                    if (ss.Trim().Length > 0) lstNew.Add(ss);
                }
                return lstNew;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static bool IsMatch(string s1, string s2)
        {
            try
            {
                string s3 = Utility.CombineString(Utility.SplitString(s1, new string[5] { "\t", "\n", "\r\n", "\r", " " }, false), " ");
                string s4 = Utility.CombineString(Utility.SplitString(s2, new string[5] { "\t", "\n", "\r\n", "\r", " " }, false), " ");

                s3 = s3.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper().Trim();
                s4 = s4.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper().Trim();
                if (s3 == s4) return true;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
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
        public static string GetSBNumber(string str)
        {
            string strSB = "";
            try
            {
                if (str.ToUpper().StartsWith("PRE SB"))
                {
                    string s1 = str.Substring(6, str.Length - 6);
                    List<string> lst = SplitString(s1, new string[2] { " ", "(" }, false);
                    if (lst.Count >= 1) strSB = lst[0];
                }
                else if (str.ToUpper().StartsWith("POST SB"))
                {
                    string s1 = str.Substring(7, str.Length - 7);
                    List<string> lst = SplitString(s1, new string[2] { " ", "(" }, false);
                    if (lst.Count >= 1) strSB = lst[0];
                }

                if (strSB.Length > 2 && strSB.StartsWith("C-"))
                {
                    strSB = strSB.Substring(2, strSB.Length - 2);
                }
                if (strSB.Length > 1 && strSB.StartsWith("C"))
                {
                    strSB = strSB.Substring(1, strSB.Length - 1);
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return strSB;
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
        public static List<string> TrimList(List<string> lst)
        {
            try
            {
                while (true)
                {
                    if (lst.Count == 0) break;
                    if (lst[lst.Count - 1].Length == 0)
                    {
                        lst.RemoveAt(lst.Count - 1);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Utility.WriteErrorLog(ex);
            }
            return lst;
        }
        public static bool IsMatch1(string s1, string s2)
        {
            try
            {
                string s3 = s1.Replace(" ", "").Replace("-", "").Replace(".", "").Replace(",", "").ToUpper().Trim();
                string s4 = s2.Replace(" ", "").Replace("-", "").Replace(".", "").Replace(",", "").ToUpper().Trim();
                if (s3 == s4) return true;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
    }
}
