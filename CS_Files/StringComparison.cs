using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
//using System.Windows.Forms;

namespace StringComparison1
{
    class Utility
    {
        public static string m_sTempPath = "";
        public static Dictionary<string, string> m_dtReplaceTexts = new Dictionary<string, string>();
        public static void WriteErrorLog(string sWarning, string sStackTrace, string sMsg)
        {
            try
            {

            }
            catch { }
        }
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
        public static List<string> ReadFile(string sFile)
        {
            List<string> lstOp = new List<string>();
            try
            {
                if (File.Exists(sFile))
                {
                    string[] sArr = File.ReadAllLines(sFile);
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
                WriteErrorLog(ee);
            }
            return lstOp;
        }
        public static bool IsDMC(string text)
        {
            try
            {
                text = text.TrimStart('\"').Trim();
                if (text.ToUpper().StartsWith("PW1000G")) return true;
                if (text.ToUpper().StartsWith("PW30K")) return true;
                if (text.ToUpper().StartsWith("V2500")) return true;

            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return false;
        }
        public static bool IsMatch1(string s1, string s2)
        {
            try
            {
                string s3 = s1.Replace(" ", "").Replace(":", "").Replace("-", "").Replace(".", "").Replace(",", "").ToUpper().Trim();
                string s4 = s2.Replace(" ", "").Replace(":", "").Replace("-", "").Replace(".", "").Replace(",", "").ToUpper().Trim();
                if (s3 == s4) return true;
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
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
        public static bool IsBulletin(string text)
        {
            try
            {
                bool isBullet = Regex.IsMatch(text, @"^\d{1,2}[A-Za-z](\d{1,2}([A-Za-z])?)?$");
                return isBullet;
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return false;
        }
    }

    public class LineInfo
    {

        public string m_sValue = "";
        public string m_sBullet = "";
        public string m_sEndDot = "";
        public int m_nPageNo = 0;
        public int m_nPageLineNo = -1;
        public bool m_bIsLastlineOfPage = false;
        // public PdfLineInfo m_oPDFLine = new PdfLineInfo();
        //public LineInfo()
        //{

        //}
        //public LineInfo(string sval)
        //{
        //    m_sValue = sval;
        //}
        public LineInfo(string sval, int lineNo)
        {
            m_sValue = sval;
            m_nPageLineNo = lineNo;
        }

    }
    public class TextInfo
    {
        public LineInfo m_oCurrentLine = null;
        public LineInfo m_oComparedLine = null;
        // public Report m_oRpt = null;
        //
        public int m_nIndex = -1;
        //Row error info
        //public TextInfo m_oOldRef = null;
        public List<TextInfo> m_lstRowErrors = new List<TextInfo>();
        //public List<TextInfo> m_lstOldRowErrors = new List<TextInfo>();
        //
        //  public int m_nOldType = -1;
        public int m_nType = 0;//0-Same, 1-Modified, 2-New, 3-Deleted, 4-table data is moved to next line, 5-Upper Text, 6-Row number representation
        //public TextInfo(LineInfo lneNew, int type)
        //{
        //    m_oNewLine = lneNew;
        //    m_nType = type;
        //}
        public TextInfo(LineInfo lneNew, int type, LineInfo lneOld)
        {
            if (lneNew == null && lneOld != null)
            {
                lneNew = new LineInfo("", lneOld.m_nPageLineNo);
            }
            else if (lneNew != null && lneOld == null)
            {
                lneOld = new LineInfo("", lneNew.m_nPageLineNo);
            }
            else if (lneNew == null && lneOld == null)
            {
                lneNew = new LineInfo("", -1);
                lneOld = new LineInfo("", -1);
            }


            m_oCurrentLine = lneNew;
            m_oComparedLine = lneOld;
            m_nType = type;
        }
    }
    public enum Launguage
    {
        English = 0,
        Japanese = 1
    }
    class StringComparison
    {
        public Launguage e_launguage = Launguage.English;
        public string m_sTempPath = "";
        public static List<string> lstStopWords = new List<string>();
        public static List<string> lstStopWordsNew = new List<string>();
        public StringComparison(string tmpPath)
        {
            try
            {
                m_sTempPath = tmpPath;
                //read the stop words
                lstStopWordsNew = Utility.ReadFile(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\bin\\stopwords.txt");
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
        }
        private void WriteErrorLog(string sWarning, string sStackTrace, string sMsg)
        {
            try
            {
                //if (sWarning.Trim().Length > 0)
                //    MessageBox.Show(sWarning.Trim());
                string sLogFilepath = m_sTempPath + "ErrorLog.dat";
                if (sLogFilepath != null && sLogFilepath.Length > 0)
                {
                    try
                    {
                        StreamWriter sw = File.AppendText(sLogFilepath);
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
                    }
                    catch
                    { }
                }
            }
            catch { }
        }
        private void WriteErrorLog(Exception ee)
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
        private bool IsEnglishLetter(char chr)
        {
            try
            {
                if (char.IsDigit(chr)) return true;
                if(chr >= 65 && chr <= 90)return true;
                if (chr >= 97 && chr <= 122) return true;
                if (chr == '-') return true;
                if (chr == ',') return true;
                if (chr == '.') return true;
                if (chr == '(') return true;
                if (chr == ')') return true;
            }
            catch { }
            return false;
        }
        public List<string> SplitString(string sValue, string schar)
        {
            List<string> lst = new List<string>();
            try
            {
                if (e_launguage == Launguage.English)
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
                else
                {
                    //Split by space
                    sValue = sValue.Replace("．", ".").Trim();
                    string[] sArr = sValue.Split(new string[1] { schar }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string ss in sArr)
                    {
                        if (ss.Trim().Length > 0)
                        {
                            lst.Add(ss.Trim());
                        }
                    }
                    //Split for Japanese---
                    List<string> lstNew = new List<string>();
                    foreach (string s in lst)
                    {
                        for (int i = 0; i < s.Length; i++)
                        {
                            int pos = -1;
                            bool y1 = IsEnglishLetter(s[i]);
                            for (int j = i + 1; j < s.Length; j++)
                            {
                                bool y2 = IsEnglishLetter(s[j]);
                                if (y2 != y1) break;
                                pos = j;
                            }
                            if (pos == -1)
                            {
                                lstNew.Add(s.Substring(i, 1));
                            }
                            else
                            {
                                lstNew.Add(s.Substring(i, pos + 1 - i));
                                i = pos;
                            }
                        }
                    }
                    //Update , . to list
                    for (int i = 0; i < lstNew.Count; i++)
                    {
                        if (i > 0 && (lstNew[i] == "," || lstNew[i] == "."))
                        {
                            lstNew[i - 1] = lstNew[i - 1] + lstNew[i];
                            lstNew.RemoveAt(i);
                            i--;
                        }
                    }

                    //Split based on ( )
                    for (int i = 0; i < lstNew.Count; i++)
                    {
                        if (lstNew[i].Contains("("))
                        {
                            try
                            {
                                int n = lstNew[i].IndexOf("(");
                                string s1 = lstNew[i].Substring(0, n);
                                string s2 = lstNew[i].Substring(n, lstNew[i].Length - s1.Length);
                                if (s1.Length > 0 && s2.Length > 0)
                                {
                                    lstNew[i] = s1;
                                    lstNew.Insert(i + 1, s2);
                                }
                            }
                            catch { }
                        }
                        if (lstNew[i].Contains(")"))
                        {
                            try
                            {
                                int n = lstNew[i].IndexOf(")") + 1;
                                string s1 = lstNew[i].Substring(0, n);
                                string s2 = lstNew[i].Substring(n, lstNew[i].Length - s1.Length);
                                if (s1.Length > 0 && s2.Length > 0)
                                {
                                    lstNew[i] = s1;
                                    lstNew.Insert(i + 1, s2);
                                }
                            }
                            catch { }
                        }
                    }
                    return lstNew;
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return lst;
        }
        List<string> SplitString(string sValue, string[] splits)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(splits, StringSplitOptions.RemoveEmptyEntries);
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
        private string CombineString(List<string> Split, string character)
        {
            string value = "";
            try
            {
                foreach (string ss in Split)
                    value = value + ss + character;
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return value.Trim();
        }
        private bool IsReal(string sValue)
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
                WriteErrorLog(ee);
            }
            return false;
        }
        private void removeStopWords(ref List<string> l1)
        {
            try
            {
                for (int i = 0; i < l1.Count; i++)
                {
                    if (lstStopWords.Contains(l1[i].ToUpper()))
                    {
                        l1.RemoveAt(i);
                        i--;
                    }
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
        }
        bool bIsWarningNoteCaution(string text)
        {
            if (text.ToUpper().StartsWith("CAUTION")) return true;
            if (text.ToUpper().StartsWith("NOTE")) return true;
            if (text.ToUpper().StartsWith("WARNING")) return true;
            return false;
        }
        public double PercentageMatch(string oldStr, string newStr, int typeOfMatch, bool maxMatch = true)
        {
            try
            {
                if (oldStr.Length == 0 || newStr.Length == 0) return 0;
                //Handle caution note warning
                bool b1 = bIsWarningNoteCaution(oldStr);
                bool b2 = bIsWarningNoteCaution(newStr);
                if (b1 && !b2 || !b1 && b2) return 0;

                //if(oldStr.Contains("1.5 hours") && newStr.Contains("1.5 hours"))
                //{
                //    int p = 1;
                //}

                //changeType 0- equal match, 1-equal/Modified

                ////-----------Only single word
                List<string> lp1 = SplitString(oldStr, " ");
                List<string> lp2 = SplitString(newStr, " ");
                if (lp1.Count == 1 && lp2.Count == 1)
                {
                    if (Utility.IsMatch1(oldStr, newStr))//.ToUpper() == newStr.ToUpper())
                    {
                        return 1.0;
                    }
                }
                else
                {
                    //Find comparison logic with new approach
                    List<TextInfo> lstTxtsold = new List<TextInfo>();
                    List<TextInfo> lstTxtsnew = new List<TextInfo>();
                    Compare(oldStr, newStr, Launguage.English, ref lstTxtsold, ref lstTxtsnew);

                    //remove stop words from list
                    for (int i = 0; i < lstTxtsold.Count; i++)
                    {
                        if (lstTxtsold[i].m_oComparedLine != null && lstStopWords.Contains(lstTxtsold[i].m_oComparedLine.m_sValue.ToUpper()) ||
                            lstTxtsnew[i].m_oCurrentLine != null && lstStopWords.Contains(lstTxtsnew[i].m_oCurrentLine.m_sValue.ToUpper()))
                        //if(lstTxtsold[i].m_nType == 0 && lstStopWords.Contains(lstTxtsold[i].m_oOldLine.m_sValue.ToUpper()))
                        {
                            lstTxtsold.RemoveAt(i);
                            lstTxtsnew.RemoveAt(i);
                            i--;
                        }
                    }

                    //Split the words and find %
                    int oldMatchCnt = 0, newMatchCnt = 0;
                    for (int x = 0; x < lstTxtsold.Count; x++)
                    {
                        TextInfo ti = lstTxtsold[x];
                        if (ti.m_nType <= typeOfMatch)
                        {
                            List<string> lst = SplitString(ti.m_oCurrentLine.m_sValue, " ");
                            removeStopWords(ref lst);
                            oldMatchCnt = oldMatchCnt + lst.Count;
                        }
                    }
                    for (int x = 0; x < lstTxtsnew.Count; x++)
                    {
                        TextInfo ti = lstTxtsnew[x];
                        if (ti.m_nType <= typeOfMatch)
                        {
                            List<string> lst = SplitString(ti.m_oCurrentLine.m_sValue, " ");
                            removeStopWords(ref lst);
                            newMatchCnt = newMatchCnt + lst.Count;
                        }
                    }
                    //remove stop words from full array
                    List<string> l1 = SplitString(oldStr, " ");
                    List<string> l2 = SplitString(newStr, " ");
                    removeStopWords(ref l1);
                    removeStopWords(ref l2);
                    //for (int i = 0; i < l1.Count; i++)
                    //{
                    //    if (lstStopWords.Contains(l1[i].ToUpper()))
                    //    {
                    //        l1.RemoveAt(i);
                    //        i--;
                    //    }
                    //}
                    //for (int i = 0; i < l2.Count; i++)
                    //{
                    //    if (lstStopWords.Contains(l2[i].ToUpper()))
                    //    {
                    //        l2.RemoveAt(i);
                    //        i--;
                    //    }
                    //}

                    //% match
                    int oldAllCnt = l1.Count;
                    int newAllCnt = l2.Count;
                    double per1 = oldMatchCnt / (double)oldAllCnt;
                    double per2 = newMatchCnt / (double)newAllCnt;

                   // if (per1 >= 0.49999 && per2 >= 0.49999)
                    {

                        if (maxMatch)
                        {
                            if (per1 > per2) return per1;
                            else return per2;
                        }
                        else
                        {
                            if (per1 < per2) return per1;
                            else return per2;
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return 0.0;
        }
        public void Compare(string old, string newstr, Launguage lauguage, ref List<TextInfo> lstTxtsold, ref List<TextInfo> lstTxtsnew)
        {
            try
            {


                e_launguage = lauguage;
                old = RecreateString(old);
                newstr = RecreateString(newstr);


                //------------------------------------------------------------------------------
                //-----------------Split the string into max match functionality----------------
                List<string> lstOld = new List<string>() { old };
                List<string> lstNew = new List<string>() { newstr };
                int count = 1, total = 10;
                for (int c = 0; c < total; c++)
                {
                    if (c == total - 3) count = 0;
                    for (int i = 0; i < lstOld.Count; i++)
                    {
                        if (lstOld[i].ToUpper() != lstNew[i].ToUpper())
                        {
                            string sold = lstOld[i];
                            string snew = lstNew[i];
                            List<string> o1 = new List<string>(), n1 = new List<string>();
                            DivideBasedOnLongestString(sold, snew, count, ref o1, ref n1);
                            //Add to list
                            for (int j = 0; j < o1.Count; j++)
                            {
                                if (j == 0)
                                {
                                    lstOld[i] = o1[0];
                                    lstNew[i] = n1[0];
                                }
                                else
                                {
                                    lstOld.Insert(i + j, o1[j]);
                                    lstNew.Insert(i + j, n1[j]);
                                }
                            }
                            i = i + o1.Count - 1;
                        }
                    }
                }
                //------------------------------------------------------------
                //----------- Handle the Start or Ends strings
                //------------------------------------------------------------
                for (int i = 0; i < lstOld.Count; i++)
                {
                    if (lstOld[i].ToUpper() != lstNew[i].ToUpper())
                    {
                        List<string> l1 = SplitString(lstOld[i], " ");
                        List<string> l2 = SplitString(lstNew[i], " ");
                        if (l1.Count > 0 && l2.Count > 0)
                        {
                            if (IsSameString1(l1[0], l2[0]))
                            {
                                lstOld[i] = l1[0];
                                l1.RemoveAt(0);
                                lstOld.Insert(i + 1, CombineString(l1, " "));
                                lstNew[i] = l2[0];
                                l2.RemoveAt(0);
                                lstNew.Insert(i + 1, CombineString(l2, " "));
                            }
                            else if (IsSameString1(l1[l1.Count - 1], l2[l2.Count - 1]))
                            {
                                string s1 = l1[l1.Count - 1];
                                string s2 = l2[l2.Count - 1];
                                l1.RemoveAt(l1.Count - 1);
                                l2.RemoveAt(l2.Count - 1);
                                lstOld[i] = CombineString(l1, " ");
                                lstOld.Insert(i + 1, s1);
                                lstNew[i] = CombineString(l2, " ");
                                lstNew.Insert(i + 1, s2);
                            }
                        }
                    }
                }
                //Fill the list
                for (int i = 0; i < lstOld.Count; i++)
                {
                    int type1 = 0, type2 = 0;
                    //Upper
                    if (lstOld[i].Trim().ToUpper() == lstNew[i].Trim().ToUpper())
                    {
                        type1 = 0;
                        type2 = 0;
                    }
                    //DMC change
                    else if (DMCChange(lstOld[i].Trim(), lstNew[i]))
                    {
                        type1 = 0;
                        type2 = 0;
                    }
                    //Do it many ways
                    else if (IsSameString2(lstOld[i], lstNew[i]))
                    {
                        type1 = 1;
                        type2 = 1;
                    }
                    else if (lstOld[i].Length == 0 && lstNew[i].Length > 0)
                    {
                        type1 = 2;
                        type2 = 2;
                    }
                    else if (lstOld[i].Length > 0 && lstNew[i].Length == 0)
                    {
                        type1 = 3;
                        type2 = 3;
                    }
                    else
                    {
                        type1 = 3;
                        type2 = 2;
                    }
                    if (type1 == 3 && type2 == 2)
                    {
                        //LineInfo lneOld = new LineInfo(lstOld[i], 1);
                        //LineInfo lneNew = new LineInfo(lstNew[i], 1);
                        //TextInfo ox = new TextInfo(lneOld, type1, lneNew);
                        //TextInfo nx = new TextInfo(lneNew, type2, lneOld);
                        //Deleted
                        LineInfo lneOld = new LineInfo(lstOld[i], 1);
                        LineInfo lneNew = new LineInfo("", 1);
                        TextInfo ox = new TextInfo(lneOld, type1, lneNew);
                        TextInfo nx = new TextInfo(lneNew, type1, lneOld);
                        lstTxtsold.Add(ox);
                        lstTxtsnew.Add(nx);
                        //New
                        lneOld = new LineInfo("", 1);
                        lneNew = new LineInfo(lstNew[i], 1);
                        ox = new TextInfo(lneOld, type2, lneNew);
                        nx = new TextInfo(lneNew, type2, lneOld);
                        lstTxtsold.Add(ox);
                        lstTxtsnew.Add(nx);
                    }
                    else
                    {
                        LineInfo lneOld = new LineInfo(lstOld[i], 1);
                        LineInfo lneNew = new LineInfo(lstNew[i], 1);
                        TextInfo ox = new TextInfo(lneOld, type1, lneNew);
                        TextInfo nx = new TextInfo(lneNew, type2, lneOld);
                        lstTxtsold.Add(ox);
                        lstTxtsnew.Add(nx);
                    }
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
        }
        private bool DMCChange(string oldstr, string newstr)
        {
            try
            {
                //if(m_bHandleDMCChange)
                //{
                //    string sNew = newstr;
                //    foreach (var v in m_dtDMCChange)
                //    {
                //        sNew = sNew.Replace(v.Key, v.Value);
                //    }
                //    if (oldstr.ToUpper() == sNew.ToUpper())
                //    {
                //        return true;
                //    }
                //}
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return false;
        }
        private bool IsSameString1(string oldstr, string newstr)
        {
            try
            {
                return IsSameString2(oldstr, newstr);
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return false;
        }
        private bool IsSameString2(string oldstr, string newstr)
        {
            try
            {
                if (oldstr.Length == 0 || newstr.Length == 0) return false;

                //if (oldstr.Contains("9002") && newstr.Contains("9002"))
                //{
                //    int y = 1;
                //}

                string sold = Trim1(oldstr);
                string snew = Trim1(newstr);
                if (sold.ToUpper() == snew.ToUpper())
                {
                    return true;
                }
                else
                {
                    //Same value
                    if (sold.Replace(" ", "").ToUpper().Trim() == snew.Replace(" ", "").ToUpper().Trim())
                    {
                        return true;
                    }
                    //Real value
                    if (sold.Length > 0 && snew.Length > 0)
                    {
                        if (IsReal(sold) && IsReal(snew))
                        {
                            return true;
                        }
                    }
                    //Any Technical
                    if (IsTechnicalString(oldstr, newstr))
                    {
                        return true;
                    }
                    //Single word comparisons
                    List<string> l1 = SplitString(sold, " ");
                    List<string> l2 = SplitString(snew, " ");
                    if (l1.Count == 1 && l2.Count == 1)
                    {
                        //Bulletin
                        if (IsBulletin(l1[0], l2[0]))
                        {
                            return true;
                        }
                        //Any integer found in the word
                        if (IsIntegerExists(l1[0]) && IsIntegerExists(l2[0]))
                        {
                            return true;
                        }
                        //Subset string
                        else if (oldstr.ToUpper().Contains(newstr.ToUpper()) || newstr.ToUpper().Contains(oldstr.ToUpper()))
                        {
                            if (lstStopWordsNew.Contains(oldstr.ToLower()) == false && lstStopWordsNew.Contains(newstr.ToLower()) == false)
                                return true;
                        }
                        //A. B.
                        if (l1[0].Length == 1 && l2[0].Length == 1)
                            return true;
                    }
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return false;
        }
        private bool IsBulletin(string s1, string s2)
        {
            try
            {
                if (Utility.IsBulletin(s1) && Utility.IsBulletin(s2)) return true;

            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return false;
        }
        private bool IsIntegerExists(string str)
        {
            try
            {
                return false;


                for (int i = 0; i < str.Length; i++)
                {
                    if (char.IsDigit(str[i])) return true;
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return false;
        }
        private bool IsTechnicalString(string str1, string str2)
        {
            try
            {

                List<string> arr1 = SplitString(TrimStartEndWithSprcialLetter(str1), new string[2] { "-", "." });
                List<string> arr2 = SplitString(TrimStartEndWithSprcialLetter(str2), new string[2] { "-", "." });
                if (arr1.Count >= 1 && arr2.Count >= 1)
                {
                    if (arr1[0] == arr2[0]) return true;
                    //Check DMC
                    if (Utility.IsDMC(arr1[0]) && Utility.IsDMC(arr2[0])) return true;
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return false;
        }
        public string RecreateString(string str)
        {
            try
            {
                str = str.Replace("\t", " ").Replace("\n", " ").Trim();
                List<string> lst = SplitString(str, " ");
                for (int i = 0; i < lst.Count; i++)
                {
                    if (lst[i] == "." && i > 0)
                    {
                        lst[i - 1] = lst[i - 1] + ".";
                        lst.RemoveAt(i);
                        i--;
                    }
                    else if (lst[i] == "," && i > 0)
                    {
                        lst[i - 1] = lst[i - 1] + ",";
                        lst.RemoveAt(i);
                        i--;
                    }
                    else if (lst[i] == ")" && i > 0)
                    {
                        lst[i - 1] = lst[i - 1] + ")";
                        lst.RemoveAt(i);
                        i--;
                    }
                    else if (lst[i] == "<" && i < lst.Count - 1)
                    {
                        lst[i + 1] = "(" + lst[i + 1];
                        lst.RemoveAt(i);
                        i--;
                    }
                }
                //Remove end . and ,
                for (int i = 0; i < lst.Count; i++)
                {
                    if (lst[i].Length > 1 && (lst[i].EndsWith(".") || lst[i].EndsWith(",") || lst[i].EndsWith(":")))//veeru 15 Sept 2026
                    {
                        lst[i] = lst[i].Substring(0, lst[i].Length - 1);
                    }
                }
                string snew = CombineString(lst, " ").Trim();
                //Added
                snew = snew.Replace(",", " ");
                snew = CombineString(SplitString(snew, " ")," ").Trim();


                return snew;
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return str;
        }
        private string Trim1(string str)
        {
            try
            {
                List<string> lstIgonrePoints = new List<string>() { ",", ".", "(", ")", ":", "-" };
                string sx = str;
                foreach (string s in lstIgonrePoints)
                {
                    sx = sx.Replace(s, "");
                }
                return sx;
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return str;
        }
        private string TrimStartEndWithSprcialLetter(string str)
        {
            try
            {
                List<string> lstIgonrePoints = new List<string>() { ",", ".", "(", ")", ":", "-" };
                string sx = str;
                foreach (string s in lstIgonrePoints)
                {
                    sx = sx.TrimStart(s[0]);
                    sx = sx.TrimEnd(s[0]);
                }
                return sx;
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return str;
        }

        private void DivideBasedOnLongestString(string old, string newstr, int count, ref List<string> lstOld, ref List<string> lstNew)
        {
            try
            {
                //Find the Long common string and split the data---
                string sLongStr = LongestCommonSubstring(old, newstr);
                if (SplitString(sLongStr, " ").Count > count)
                {
                    lstOld = SplitLongestString(old, sLongStr);
                    lstNew = SplitLongestString(newstr, sLongStr);
                }
                else
                {
                    lstOld = new List<string>() { "", old, "" };
                    lstNew = new List<string>() { "", newstr, "" };
                }

                for (int i = 0; i < lstOld.Count; i++)
                {
                    if (lstOld[i].Length == 0 && lstNew[i].Length == 0)
                    {
                        lstOld.RemoveAt(i);
                        lstNew.RemoveAt(i);
                        i--;
                    }
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
        }
        private List<string> SplitLongestString(string str, string sLongStr)
        {
            List<string> lst = new List<string>();
            try
            {
                //change the logic
                int n1 = str.ToUpper().IndexOf(" " + sLongStr.ToUpper() + " ");
                if (n1 == -1)
                {
                    n1 = str.ToUpper().IndexOf(" " + sLongStr.ToUpper());
                    if (n1 == -1)
                    {
                        n1 = str.ToUpper().IndexOf(sLongStr.ToUpper() + " ");
                        if (n1 == -1)
                        {
                            n1 = str.ToUpper().IndexOf(sLongStr.ToUpper());
                        }
                    }
                    else
                    {
                        n1 = n1 + 1;
                    }
                }
                else
                {
                    n1 = n1 + 1;
                }
                if (n1 != -1)
                {
                    List<string> lstx = new List<string>();
                    string s1 = "", s2 = "", s3 = "";
                    if (n1 > 0)
                    {
                        s1 = str.Substring(0, n1);
                    }
                    s2 = str.Substring(n1, sLongStr.Length);
                    int l = str.Length - (n1 + sLongStr.Length);
                    if (l > 0)
                    {
                        s3 = str.Substring(n1 + sLongStr.Length, l);
                    }
                    lst.Add(s1.Trim());
                    lst.Add(s2.Trim());
                    lst.Add(s3.Trim());
                }
                else
                {
                    lst.Add(str);
                }
            }
            catch (Exception ee)
            {
                WriteErrorLog(ee);
            }
            return lst;
        }
        private string LongestCommonSubstring(string old, string newstr)
        {
            List<string> source = SplitString(old, " ");
            List<string> target = SplitString(newstr, " ");
            if (source.Count == 0 || target.Count == 0) return "";

            var L = new int[source.Count, target.Count];
            var maximumLength = 0;
            var lastSubsBegin = 0;
            var stringBuilder = new StringBuilder();

            for (var i = 0; i < source.Count; i++)
            {
                for (var j = 0; j < target.Count; j++)
                {
                    if (source[i].ToUpper() != target[j].ToUpper())
                    {
                        L[i, j] = 0;
                    }
                    else
                    {
                        if (i == 0 || j == 0)
                            L[i, j] = 1;
                        else
                            L[i, j] = 1 + L[i - 1, j - 1];

                        if (L[i, j] > maximumLength)
                        {
                            maximumLength = L[i, j];
                            var thisSubsBegin = i - L[i, j] + 1;
                            if (lastSubsBegin == thisSubsBegin)
                            {
                                //if the current LCS is the same as the last time this block ran
                                stringBuilder.Append(source[i] + " ");
                            }
                            else //this block resets the string builder if a different LCS is found
                            {
                                lastSubsBegin = thisSubsBegin;
                                stringBuilder.Length = 0; //clear it
                                //stringBuilder.Append(source.Substring(lastSubsBegin, i + 1 - lastSubsBegin));
                                int cnt = lastSubsBegin + i + 1 - lastSubsBegin;
                                for (int s = lastSubsBegin; s < cnt; s++)
                                    stringBuilder.Append(source[s] + " ");
                            }
                        }
                    }
                }
            }
            string output = stringBuilder.ToString();
            return output.Trim();
        }
        private string LongestCommonSubstringOld(string source, string target)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target)) return "";

            var L = new int[source.Length, target.Length];
            var maximumLength = 0;
            var lastSubsBegin = 0;
            var stringBuilder = new StringBuilder();

            for (var i = 0; i < source.Length; i++)
            {
                for (var j = 0; j < target.Length; j++)
                {
                    if (source[i].ToString().ToUpper() != target[j].ToString().ToUpper())
                    {
                        L[i, j] = 0;
                    }
                    else
                    {
                        if (i == 0 || j == 0)
                            L[i, j] = 1;
                        else
                            L[i, j] = 1 + L[i - 1, j - 1];

                        if (L[i, j] > maximumLength)
                        {
                            maximumLength = L[i, j];
                            var thisSubsBegin = i - L[i, j] + 1;
                            if (lastSubsBegin == thisSubsBegin)
                            {
                                //if the current LCS is the same as the last time this block ran
                                stringBuilder.Append(source[i]);
                            }
                            else //this block resets the string builder if a different LCS is found
                            {
                                lastSubsBegin = thisSubsBegin;
                                stringBuilder.Length = 0; //clear it
                                stringBuilder.Append(source.Substring(lastSubsBegin, i + 1 - lastSubsBegin));
                            }
                        }
                    }
                }
            }
            string output = stringBuilder.ToString();
            return output;
        }
    }
}
namespace SequenceAlignment
{
    public class AlignmentResult
    {
        public string OldSentence { get; set; }
        public string NewSentence { get; set; }
        public string Status { get; set; }

        public override string ToString()
        {
            return $"{Status} : {OldSentence} -> {NewSentence}";
        }
    }

    public class SentenceAligner
    {
        private const double MinMatch = 0.5;
        private const double GapPenalty = -0.2;

        public static List<AlignmentResult> Align(
            string[] oldSentences,
            string[] newSentences,
            double[,] similarity)
        {
            int m = oldSentences.Length;
            int n = newSentences.Length;

            double[,] dp = new double[m + 1, n + 1];

            for (int i = 1; i <= m; i++)
                dp[i, 0] = dp[i - 1, 0] + GapPenalty;

            for (int j = 1; j <= n; j++)
                dp[0, j] = dp[0, j - 1] + GapPenalty;

            for (int i = 1; i <= m; i++)
            {
                for (int j = 1; j <= n; j++)
                {
                    double sim = similarity[i - 1, j - 1];

                    double matchScore =
                        sim >= MinMatch
                        ? dp[i - 1, j - 1] + sim
                        : double.MinValue;

                    double deleteScore =
                        dp[i - 1, j] + GapPenalty;

                    double addScore =
                        dp[i, j - 1] + GapPenalty;

                    dp[i, j] = Math.Max(
                        matchScore,
                        Math.Max(deleteScore, addScore));
                }
            }

            List<AlignmentResult> result =
                new List<AlignmentResult>();

            int row = m;
            int col = n;

            while (row > 0 || col > 0)
            {
                if (row > 0 && col > 0)
                {
                    double sim = similarity[row - 1, col - 1];

                    if (sim >= MinMatch)
                    {
                        double matchScore =
                            dp[row - 1, col - 1] + sim;

                        if (Math.Abs(dp[row, col] - matchScore) < 0.0001)
                        {
                            result.Add(new AlignmentResult
                            {
                                OldSentence = oldSentences[row - 1],
                                NewSentence = newSentences[col - 1],
                                Status = "Matched"
                            });

                            row--;
                            col--;
                            continue;
                        }
                    }
                }

                if (col > 0 &&
                    Math.Abs(dp[row, col] -
                             (dp[row, col - 1] + GapPenalty))
                    < 0.0001)
                {
                    result.Add(new AlignmentResult
                    {
                        OldSentence = "",
                        NewSentence = newSentences[col - 1],
                        Status = "Added"
                    });

                    col--;
                }
                else
                {
                    result.Add(new AlignmentResult
                    {
                        OldSentence = oldSentences[row - 1],
                        NewSentence = "",
                        Status = "Deleted"
                    });

                    row--;
                }
            }

            result.Reverse();
            return result;
        }
    }
}