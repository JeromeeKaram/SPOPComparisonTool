using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EngineManualReader
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
        public static List<string> SplitString1(string sValue, string schars, bool trim)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(new string[1] { schars }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string ss in sArr)
                {
                    if (trim)
                    {
                        if (ss.Trim().Length > 0)
                        {
                            lst.Add(ss.Trim());
                        }
                    }
                    else
                    {
                        lst.Add(ss);
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
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
        public static List<string> SplitString4(string str, string[] schars)
        {
            List<string> lst = new List<string>();
            try
            {
                string extension = "";
                foreach (string s in schars)
                {
                    if (str.StartsWith(s))
                    {
                        extension = "ihi";
                    }
                }
                lst = SplitString(extension + str, schars, false);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static List<string> SplitString5(string str, string[] schars)
        {
            List<string> lst = new List<string>();
            try
            {
                string extension = "";
                foreach (string s in schars)
                {
                    if (str.StartsWith(s))
                    {
                        extension = "ihi";
                    }
                }
                lst = SplitString(extension + str, schars, true);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static string CombineString(List<string> Split, string character)
        {
            string value = "";
            try
            {
                foreach (string ss in Split)
                    value = value + ss + character;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return value.Trim();
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
    }


    public class SeleniumReader
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;  // Hide the window
        const int SW_SHOW = 5;  // Show the window
        const int SW_MINIMIZE = 6;
        const int SW_RESTORE = 9;

        public WebDriver driver1 = null;

        public SeleniumReader()
        {
            var options = new ChromeOptions();

            options.AddArgument("--headless=new");
            options.AddArgument("--disable-gpu");
            options.AddArgument("--window-size=1920,1080");
            options.AddArgument("--log-level=3");

            var service = ChromeDriverService.CreateDefaultService(".", "chromedriver.exe");
            service.HideCommandPromptWindow = true;

            driver1 = new ChromeDriver(service, options);
        }

        public string Read(string link)
        {
            string innerHtml = "";
            try
            {

                driver1.Navigate().GoToUrl(link);
                IWebElement webElement = null;
                System.Threading.Thread.Sleep(3000);
                int cntx = 0;
                while (cntx < 3)
                {
                    cntx = cntx + 1;
                    try
                    {
                        webElement = driver1.FindElement(By.ClassName("header"));
                        break;
                    }
                    catch
                    {
                        link = link.Replace("PWSPROC", "SP1100G");
                        driver1.Navigate().GoToUrl(link);
                    }
                }
                int cnt = 0;
                if (webElement != null)
                {
                    string text = getText(webElement);
                    text = text.Replace("</h1>", "").Replace("<h1>", "").Replace("\r\n", "").Replace(" ", "").Trim();
                    //---Check IETP
                    if (text.Contains("IETP Terms and Conditions".Replace(" ", "").Trim()) || text.Contains("(IETP) Terms and Conditions".Replace(" ", "").Trim()))
                    {
                        cnt = 0;
                        while (true)
                        {
                            driver1.FindElement(By.ClassName("legalButton")).Click();
                            System.Threading.Thread.Sleep(1000);
                            webElement = driver1.FindElement(By.ClassName("header"));
                            System.Threading.Thread.Sleep(1000);
                            text = getText(webElement);
                            text = text.Replace("</h1>", "").Replace("<h1>", "").Replace("\r\n", "").Replace(" ", "").Trim();
                            if (text.Contains("IETP Terms and Conditions".Replace(" ", "").Trim()) == false)
                            {
                                break;
                            }
                            if (text.Contains("(IETP) Terms and Conditions".Replace(" ", "").Trim()) == false)
                            {
                                break;
                            }
                            cnt = cnt + 1;
                            if (cnt == 20) break;
                        }
                    }
                }
                //---Now read the data
                cnt = 0;
                while (true)
                {
                    try
                    {
                        webElement = driver1.FindElement(By.Id("main"));
                        innerHtml = getText(webElement);
                        break;
                    }
                    catch
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    cnt = cnt + 1;
                    if (cnt == 20) break;
                }


            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return innerHtml;
        }
        public string getText(IWebElement webElement)
        {
            string innerHtml = "";
            try
            {
                IJavaScriptExecutor js = driver1 as IJavaScriptExecutor;
                if (js != null)
                {
                    innerHtml = (string)js.ExecuteScript("return arguments[0].innerHTML;", webElement);
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return innerHtml;
        }
        public void close()
        {
            try
            {
                try
                {
                    if (driver1 != null)
                    {
                        driver1.Quit();
                        driver1.Dispose();
                        driver1 = null;
                    }
                }
                catch
                {
                }
                // Kill ChromeDriver if it is still running
                Process[] processlist = Process.GetProcesses();
                foreach (Process theprocess in processlist)
                {
                    try
                    {
                        if (theprocess.ProcessName.Equals("chromedriver", System.StringComparison.OrdinalIgnoreCase))
                        {
                            theprocess.Kill();
                            theprocess.WaitForExit();
                        }
                    }
                    catch
                    {
                        // Ignore process termination errors
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public string RemoveTableData(string givenText)
        {
            try
            {
                if (givenText.Contains("<table") && givenText.Contains("</table>"))
                {
                    string sx = "";
                    List<string> lst1 = Utility.SplitString(givenText, "<table");
                    foreach (string ss in lst1)
                    {
                        if (ss.Contains("</table>"))
                        {
                            List<string> lst2 = Utility.SplitString(ss, "</table>");
                            sx = sx + lst2[1];
                        }
                        else
                        {
                            sx = sx + ss;
                        }
                    }
                    return sx;
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return givenText;
        }
        public void splitElements(string givenText, string element, ref List<string> elements, ref List<string> elementsWithOutTables)
        {
            try
            {
                string splitText = "test " + givenText;
                List<string> lst = Utility.SplitStringAddItChar(splitText, "<" + element);

                int cnt1 = lst.Count;
                //Avoid other step movements like -,. etc..PWSPROC-C-72-35-00-45A-066A-B, Step1
                for (int i = 1; i < lst.Count; i++)
                {
                    if (lst[i].ToLower().Contains("class=\"steplevel".ToLower()) == false)
                    {
                        lst[i - 1] = lst[i - 1] + lst[i];
                        lst.RemoveAt(i);
                        i--;
                    }
                }
                if (cnt1 - lst.Count > 0)
                {
                    Utility.WriteErrorLog("", "", "VEERU : Found some step movements");
                }

                string end = "</" + element + ">";
                for (int i = 1; i < lst.Count; i++)
                {
                    if (lst[i].StartsWith(end) == false)
                    {
                        string sss = lst[i];
                        if (sss.StartsWith(">")) sss = sss.Substring(1, sss.Length - 1);


                        //---Added below code on 5-Mar-24 to address PW1000G-C-72-00-0000-00A-250A-C (Requirements After Job Completion) issue
                        if (sss.Contains(end) && sss.EndsWith(end) == false)
                        {
                            int n1 = sss.IndexOf(end) + end.Length;
                            sss = sss.Substring(0, n1);
                        }

                        //----
                        elements.Add(sss);
                    }
                    else
                    {
                    }
                }

                foreach (string ss in elements) elementsWithOutTables.Add(ss);

                //Remove the table data
                for (int i = 0; i < elements.Count; i++)
                {
                    elementsWithOutTables[i] = RemoveTableData(elements[i]);
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
    }

    //--------------------- DATA STRUCTURES ------------
    #region Data_structures
    public enum CompareType
    {
        SAME = 0,
        MODIFIED = 1,
        NEW = 2,
        DELETED = 3,
        CHECK = 4,
        NONE
    }
    public class ProcedureOutput
    {
        //int index = 0;
        public CompareType m_eStatus = CompareType.NEW;
        public ProcedureInfo m_old = null;
        public ProcedureInfo m_new = null;
        public NoteInfo m_oldNote = null;
        public NoteInfo m_newNote = null;
    }
    public class ProcedureOutputNew
    {
        //int index = 0;
        public CompareType m_eStatus = CompareType.SAME;
        public ProcedureInfo m_old = null;
        public ProcedureInfo m_new = null;
        public ProcedureInfo m_Original = null;
        //
        public List<ProcedureOutput> lstWarningComparison = new List<ProcedureOutput>();
        public List<ProcedureOutput> lstNotesComparison = new List<ProcedureOutput>();
        //
        public ProcedureOutputNew(ProcedureInfo oo)
        {
            m_Original = oo;
        }
    }
    public class ProcedureInfo
    {
        public string m_sDMApplicability = "";
        public string m_sCommonInformation = "";
        public string m_sCommonInformationTitle = "";
        //
        public bool m_bIsNewlyAdded = false;
        public int m_nPos = 0;
        public string m_sBullet = "";
        public string m_sCompleteBullet = "";
        //public string m_sCompleteBullet = "";
        public List<NoteInfo> lstSafetyRequirements = new List<NoteInfo>();
        //
        public ProcedureOutput m_oComparedObject = null;
        public CompareType m_eStatus = CompareType.NEW;
        public int m_nIndex = 0;
        public string m_sID = "";
        public string m_sText = "";
        public string m_sFullText = "";
        //public List<ProcedureInfo> m_lstDeleted = new List<ProcedureInfo>();
        public List<ProcedureInfo> m_lstChilds = new List<ProcedureInfo>();
        public ProcedureInfo m_oParent = null;

        public List<ProcedureOutputNew> lstOps = new List<ProcedureOutputNew>();

        public List<ProcedureOutputNew> lstComparisonOutput = new List<ProcedureOutputNew>();

        public List<NoteInfo> lstBeforeNotes = new List<NoteInfo>();
        public List<NoteInfo> lstAfterNotes = new List<NoteInfo>();


        public ProcedureInfo() { }
        public ProcedureInfo(string str)
        {
            m_sText = str;
        }
        public void AddNoteWarningCaution(string type, string noteTxt, int pos)
        {
            try
            {
                //yes found...add the note and caution objects
                if (type == "NOTE")
                    lstAfterNotes.Add(new NoteInfo("NOTE " + noteTxt, pos));
                else
                    lstBeforeNotes.Add(new NoteInfo(type + " " + noteTxt, pos));
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public NoteInfo AddNoteWarningCautionNew(string type, string noteTxt, int pos)
        {
            NoteInfo onote = null;
            try
            {
                //before
                if (pos == 0)
                {
                    onote = new NoteInfo(type + " " + noteTxt, pos);
                    lstBeforeNotes.Add(onote);
                }
                //after
                else
                {
                    onote = new NoteInfo(type + " " + noteTxt, pos);
                    lstAfterNotes.Add(onote);
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return onote;
        }
    }
    public class NoteInfo
    {
        public int m_nTextLoc = 0;//0-starts with, 1 - ends with
                                  // public HtmlAgilityPack.HtmlNode m_node = null;
        public string m_sText = "";
        public string m_sNewText = "";
        public NoteType m_eType = NoteType.NONE;
        public int m_nPos = 0;
        public NoteInfo(string txt, int loc)
        {
            m_sText = txt;
            m_sNewText = txt;
            m_nTextLoc = loc;
            m_eType = NoteType.NOTE;

            if (m_sText.StartsWith("WARNING"))
            {
                m_sNewText = m_sText.Substring(7, m_sText.Length - 7).Trim();
                m_eType = NoteType.WARNING;
            }
            else if (m_sText.StartsWith("CAUTION"))
            {
                m_sNewText = m_sText.Substring(7, m_sText.Length - 7).Trim();
                m_eType = NoteType.CAUTION;
            }
            else if (m_sText.StartsWith("NOTE"))
            {
                m_sNewText = m_sText.Substring(4, m_sText.Length - 4).Trim();
            }
        }
        public void AddText(string str)
        {
            m_sText = m_sText + "\n" + str;
            m_sNewText = m_sNewText + "\n" + str;
        }
    }
    public enum NoteType
    {
        WARNING = 0,
        CAUTION = 1,
        NOTE = 2,
        NONE
    }

    #endregion

    //------------------------------------------------------------

    public class ComparisonHelper
    {
        public List<ComparisonResult> Compare(ProcedureInfo oOldParent, ProcedureInfo onewParent)
        {
            List<ProcedureInfo> lstAllOld = new List<ProcedureInfo>();
            List<ProcedureInfo> lstAllNew = new List<ProcedureInfo>();
            Recursive(oOldParent, ref lstAllOld);
            Recursive(onewParent, ref lstAllNew);
            return Compare(lstAllOld, lstAllNew);
        }
        public List<ComparisonResult> Compare(List<ProcedureInfo> lstOld, List<ProcedureInfo> lstNew)
        {
            List<ComparisonResult> lstCompares = new List<ComparisonResult>();
            try
            {
                var strCompareObj = new StringComparison1.StringComparison(Utility.m_sTempPath );

                Dictionary<string, ProcedureInfo> dtOld = new Dictionary<string, ProcedureInfo>();
                Dictionary<string, ProcedureInfo> dtNew = new Dictionary<string, ProcedureInfo>();

                string[] oldSentences = new string[lstOld.Count];
                string[] newSentences = new string[lstNew.Count];
                for (int i = 0; i < lstOld.Count; i++)
                {
                    oldSentences[i] = "Old" + i.ToString();
                    dtOld.Add(oldSentences[i], lstOld[i]);
                }
                for (int i = 0; i < lstNew.Count; i++)
                {
                    newSentences[i] = "New" + i.ToString();
                    dtNew.Add(newSentences[i], lstNew[i]);
                }
                double[,] similarity = new double[lstOld.Count, lstNew.Count];
                for (int i = 0; i < lstOld.Count; i++)
                {
                    for (int j = 0; j < lstNew.Count; j++)
                    {
                        similarity[i, j] = strCompareObj.PercentageMatch(lstOld[i].m_sText, lstNew[j].m_sText, 0, true);
                    }
                }
                //--
                List<SequenceAlignment.AlignmentResult> lstAlignmentResult = SequenceAlignment.SentenceAligner.Align(oldSentences, newSentences, similarity);
                for (int i = 0; i < lstAlignmentResult.Count; i++)
                {
                    if (lstAlignmentResult[i].Status == "Matched")
                    {
                        CompareType typep = CompareType.MODIFIED;
                        if (Utility.IsMatch1(dtNew[lstAlignmentResult[i].NewSentence].m_sText, dtOld[lstAlignmentResult[i].OldSentence].m_sText)) typep = CompareType.SAME;
                        ComparisonResult oCmp = new ComparisonResult(dtNew[lstAlignmentResult[i].NewSentence], dtOld[lstAlignmentResult[i].OldSentence], typep);
                        lstCompares.Add(oCmp);
                    }
                    else if (lstAlignmentResult[i].Status == "Added")
                    {
                        ComparisonResult oCmp = new ComparisonResult(dtNew[lstAlignmentResult[i].NewSentence], null, CompareType.NEW);
                        lstCompares.Add(oCmp);
                    }
                    else if (lstAlignmentResult[i].Status == "Deleted")
                    {
                        ComparisonResult oCmp = new ComparisonResult(null, dtOld[lstAlignmentResult[i].OldSentence], CompareType.DELETED);
                        lstCompares.Add(oCmp);
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lstCompares;
        }
        private void Recursive(EngineManualReader.ProcedureInfo proc, ref List<ProcedureInfo> lstAll)
        {
            if(proc.m_nIndex > 0) lstAll.Add(proc);
            foreach (EngineManualReader.ProcedureInfo child in proc.m_lstChilds)
                Recursive(child, ref lstAll);
        }
    }
    public class ComparisonResult
    {
        public CompareType m_eCompareType = CompareType.NEW;
        public ProcedureInfo m_new = null;
        public ProcedureInfo m_old = null;
        public ComparisonResult(ProcedureInfo n, ProcedureInfo o, CompareType e)
        {
            m_new = n;
            m_old = o;
            m_eCompareType = e;
        }
    }

    public class EMReader
    {
        List<string> lstThings = new List<string>() { "WARNING", "CAUTION", "NOTE" };
        public ProcedureInfo Read(string htmlLink, string linktype, SeleniumReader m_oSeleniumReader)
        {
            ProcedureInfo oMainProcedure = new ProcedureInfo();
            try
            {
                List<ProcedureInfo> lstProcedureObjs = new List<ProcedureInfo>();
                List<ProcedureInfo> lstDuplicateProcedureObjs = new List<ProcedureInfo>();
                int toplevel = 0, lastIndex = 0;
                List<string> lstpoints = new List<string>();


                //-->>> Read from Web browser
                int nProcedureHeaderIndex = -1;
                List<string> lstProcIndex = new List<string>() { "<h1>Procedure</h1>", "<h1>Description</h1>" };
                string outerHtml = m_oSeleniumReader.Read(htmlLink);
                string m_sProcDesc = "Procedure";
                foreach (string ss in lstProcIndex)
                {
                    nProcedureHeaderIndex = outerHtml.IndexOf(ss);
                    if (nProcedureHeaderIndex != -1)
                    {
                        m_sProcDesc = ss.Replace("<h1>", "").Replace("</h1>", "").Trim();
                        nProcedureHeaderIndex = nProcedureHeaderIndex + ss.Length;
                        break;
                    }
                }
                //DM Applicability 
                try
                {
                    IWebElement sElement = m_oSeleniumReader.driver1.FindElement(By.Id("dmApplicability"));
                    List<string> lstz = Utility.SplitString(sElement.Text, "\r\n");
                    for (int x = 1; x < lstz.Count; x++)
                    {
                        oMainProcedure.m_sDMApplicability = oMainProcedure.m_sDMApplicability + lstz[x] + " ";
                    }
                    oMainProcedure.m_sDMApplicability = Utility.CombineString(Utility.SplitString(oMainProcedure.m_sDMApplicability, " "), " ");
                }
                catch { }
                //common information
                try
                {
                    IWebElement sElement = m_oSeleniumReader.driver1.FindElement(By.Id("commonInfo"));
                    string txtAll = sElement.Text;
                    IWebElement[] webElements = sElement.FindElements(By.TagName("table")).ToArray();
                    foreach (IWebElement ooo in webElements)
                    {
                        txtAll = txtAll.Replace(ooo.Text, "");
                    }
                    List<string> lstz = Utility.SplitString(txtAll, "\r\n");
                    for (int x = 1; x < lstz.Count; x++)
                    {
                        if (x == 1)
                            oMainProcedure.m_sCommonInformationTitle = lstz[x];
                        else
                            oMainProcedure.m_sCommonInformation = oMainProcedure.m_sCommonInformation + lstz[x] + " ";
                    }
                    oMainProcedure.m_sCommonInformationTitle = Utility.CombineString(Utility.SplitString(oMainProcedure.m_sCommonInformationTitle, " "), " ");
                    oMainProcedure.m_sCommonInformation = Utility.CombineString(Utility.SplitString(oMainProcedure.m_sCommonInformation, " "), " ");
                }
                catch { }

                if (outerHtml.Contains("Safety Requirements"))
                {
                    try
                    {
                        List<string> lp = new List<string>();
                        string sSafetyData = Utility.SplitString(Utility.SplitString(outerHtml, "Safety Requirements")[1], "<h1>" + m_sProcDesc + "</h1>")[0];
                        string sNewText = GetElementText(sSafetyData, ref lp);
                        List<string> lstTxts = Utility.SplitString(sNewText, "\n");
                        for (int i = 0; i < lstTxts.Count; i++)
                        {
                            if (lstTxts[i] == "WARNING" || lstTxts[i] == "CAUTION" || lstTxts[i] == "NOTE")
                            {
                                string str = "";
                                int iIndex = i;
                                for (int j = i + 1; j < lstTxts.Count; j++)
                                {
                                    if (lstTxts[j] == "WARNING" || lstTxts[j] == "CAUTION" || lstTxts[j] == "NOTE")
                                    {
                                        i = j - 1;
                                        break;
                                    }
                                    str = str + lstTxts[j] + " ";
                                    if (j == lstTxts.Count)
                                    {
                                        i = j;
                                    }
                                }
                                if (str.Trim().Length > 0)
                                {
                                    NoteInfo nt = new NoteInfo(lstTxts[iIndex] + " " + str.Trim(), oMainProcedure.lstSafetyRequirements.Count);
                                    oMainProcedure.lstSafetyRequirements.Add(nt);
                                }
                            }
                        }
                    }
                    catch { }
                }

                //-->>> Find the bulletin points
                //doit(outerHtml);
                List<string> elements = new List<string>();
                List<string> elementsWithOutTables = new List<string>();
                m_oSeleniumReader.splitElements(outerHtml, "li", ref elements, ref elementsWithOutTables);
                if (nProcedureHeaderIndex != -1 && elements.Count > 0)
                {
                    for (int i = 0; i < elements.Count; i++)
                    {
                        string text = elementsWithOutTables[i];
                        int nlocIndex = outerHtml.IndexOf(elements[i]);
                        if (nlocIndex > nProcedureHeaderIndex)
                        {
                            //Divide table text---

                            //get the level
                            int parentnum = 0;
                            if (lstProcedureObjs.Count > 0) parentnum = lstProcedureObjs[lstProcedureObjs.Count - 1].m_nIndex;
                            int n1 = getLevelIndex(text, parentnum);
                            if (n1 == -1)
                            {
                                if (lstDuplicateProcedureObjs.Count > 0)
                                    n1 = lstDuplicateProcedureObjs[lstDuplicateProcedureObjs.Count - 1].m_nIndex;
                                else
                                    n1 = 0;
                            }

                            if (lstDuplicateProcedureObjs.Count == 88)
                            {
                                int x = 5;
                            }

                            if (i == 0) toplevel = n1;
                            int level = n1 - toplevel;
                            if (level > lastIndex) lastIndex = level;
                            //get the text
                            string sNewText = GetElementText(text, ref lstpoints);/// Utility.CombineString(Utility.SplitString(liElement.Text, "\r\n"), "\n");
                            if (sNewText.Length == 0) continue;
                            //---create procedure object
                            ProcedureInfo oo = new ProcedureInfo();
                            oo.m_sText = Utility.ReplaceNonASCIIChars(sNewText);
                            oo.m_nIndex = level;
                            //oo.m_node = element;
                            lstProcedureObjs.Add(oo);
                            lstDuplicateProcedureObjs.Add(oo);

                            //EMPTY TEXT in FIRST LEVEL
                            if (oo.m_sText.Length == 0 && oo.m_nIndex == 0)
                            {
                                oo.m_sText = "NULL";
                            }
                        }
                    }
                }
                else if (nProcedureHeaderIndex != -1)
                {
                    string str = Utility.SplitString(outerHtml, "class=\"figure\"")[0];
                    string str11 = str.Substring(nProcedureHeaderIndex, str.Length - nProcedureHeaderIndex);
                    string str22 = m_oSeleniumReader.RemoveTableData(str11);
                    string sNewText = Utility.FindHtmlText(str22);
                    ProcedureInfo oo = new ProcedureInfo();
                    oo.m_sText = Utility.ReplaceNonASCIIChars(sNewText);
                    oo.m_nIndex = 0;
                    lstProcedureObjs.Add(oo);
                    lstDuplicateProcedureObjs.Add(oo);
                }

                //Tables reading
               // oMainProcedure.m_lstProcTables = Utility.m_oSeleniumReader.ReadTables(linktype, htmlLink);
              //  UpdateForMCMI(oMainProcedure, htmlLink);
                //----------------------------------------------------------------
                //------------------ MAKE PARENT CHILD RELATION ------------------
                for (int level = lastIndex; level >= 1; level--)
                {
                    for (int x = 0; x < lstProcedureObjs.Count; x++)
                    {
                        if (lstProcedureObjs[x].m_nIndex == level)
                        {
                            //find parent
                            ProcedureInfo oparent = null;
                            for (int y = x - 1; y >= 0; y--)
                            {
                                if (lstProcedureObjs[y].m_nIndex == level - 1)
                                {
                                    oparent = lstProcedureObjs[y]; break;
                                }
                            }
                            //Add to parent
                            if (oparent != null)
                            {
                                oparent.m_lstChilds.Add(lstProcedureObjs[x]);
                                lstProcedureObjs[x].m_oParent = oparent;
                                lstProcedureObjs.RemoveAt(x);
                                x--;
                            }
                        }
                    }
                }

                //-------------------------------------------------------------------
                //------------ FIND THE NEW PROCEDURE POINT FRON CHILD VALUE --------
                //-------------------------------------------------------------------
                foreach (ProcedureInfo oo in lstDuplicateProcedureObjs)
                {
                    if (oo.m_lstChilds.Count > 0)
                    {
                        //Find the last index of the list from child data
                        List<string> lst1 = Utility.SplitString(oo.m_sText, "\n");
                        int startloc = GetNoteStartIndex(lst1);
                        int lastloc = lst1.Count;
                        List<string> lx = Utility.SplitString(oo.m_lstChilds[0].m_sText, "\n");
                        if (lx.Count == 0) lx.Add("");
                        string nextVal = lx[0];// Utility.SplitString(oo.m_lstChilds[0].m_sText, "\n")[0];
                        for (int j = startloc; j < lst1.Count; j++)
                        {
                            if (lst1[j] == nextVal)
                            {
                                lastloc = j; break;
                            }
                        }
                        //push the newly formed string
                        List<string> lst3 = new List<string>();
                        for (int j = 0; j < lastloc; j++)
                            lst3.Add(lst1[j]);
                        oo.m_sText = Utility.CombineString(lst3, "\n");
                    }
                }

                //-----NOTE are NEXT to point
                //this false condition modified on 5-Sept-23 as notes are present some times above the step number...
                //So below condition is not required...
                if (false)
                {
                    //foreach (ProcedureInfo oo in lstDuplicateProcedureObjs)
                    for (int i = 0; i < lstDuplicateProcedureObjs.Count; i++)
                    {
                        while (true)
                        {
                            List<string> lstTxts = Utility.SplitString(lstDuplicateProcedureObjs[i].m_sText, "\n");
                            if (lstTxts.Count >= 2 && lstTxts[0] == "NOTE" && i > 0)
                            {
                                lstDuplicateProcedureObjs[i - 1].m_sText = lstDuplicateProcedureObjs[i - 1].m_sText + "\n" + "NOTE" + "\n" + lstTxts[1];
                                lstTxts = Utility.SplitString(lstDuplicateProcedureObjs[i].m_sText, "\n");
                                lstTxts.RemoveAt(0);
                                lstTxts.RemoveAt(0);
                                lstDuplicateProcedureObjs[i].m_sText = Utility.CombineString(lstTxts, "\n");
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }

                //-------------------------------------------------------------------------
                //--------------------- Seperate the WARNINGS/CAUTION/NOTES ---------------
                //-------------------------------------------------------------------------
                //changed the logic on 5-Sept-23, this is because note & cautions any location in step numbers
                if (false)
                {
                    //"NOTE's" are exists below the procedure bullet
                    foreach (ProcedureInfo oo in lstDuplicateProcedureObjs)
                    {
                        List<string> lstTxts = Utility.SplitString(oo.m_sText, "\n");
                        for (int p = 0; p < lstTxts.Count - 1; p++)
                        {
                            if (lstTxts[p] == "NOTE")
                            {
                                oo.AddNoteWarningCaution("NOTE", lstTxts[p + 1], 1);
                                lstTxts.RemoveAt(p);
                                lstTxts.RemoveAt(p);
                                p--;
                            }
                        }
                        oo.m_sText = Utility.CombineString(lstTxts, "\n");
                    }

                    //"WARNING", "CAUTION", "NOTE" are exists above the procedure bullet
                    foreach (ProcedureInfo oo in lstDuplicateProcedureObjs)
                    {
                        List<string> lstTxts = Utility.SplitString(oo.m_sText, "\n");
                        //loop it
                        int cnt = 0;
                        while (true)
                        {
                            if (lstTxts.Count < 2) break;
                            if (cnt == 20) break;
                            int n1 = lstThings.IndexOf(lstTxts[0]);
                            if (n1 != -1)
                            {
                                //yes found...add the note and caution objects
                                oo.AddNoteWarningCaution(lstThings[n1], lstTxts[1], 0);
                                //remove
                                lstTxts.RemoveAt(0);
                                lstTxts.RemoveAt(0);
                            }
                            else
                            {
                                break;
                            }
                            cnt = cnt + 1;
                        }
                        oo.m_sText = Utility.CombineString(lstTxts, "\n");
                    }
                }
                else
                {
                    int x = 0;
                    foreach (ProcedureInfo oo in lstDuplicateProcedureObjs)
                    {
                        int pos = 0;
                        List<string> lstTxts = Utility.SplitString(oo.m_sText, "\n");
                        for (int i = 0; i < lstTxts.Count; i++)
                        {
                            int n1 = lstThings.IndexOf(lstTxts[i]);
                            if (n1 != -1 && i + 1 < lstTxts.Count)
                            {
                                NoteInfo note = oo.AddNoteWarningCautionNew(lstThings[n1], lstTxts[i + 1], pos);
                                lstTxts.RemoveAt(i);
                                lstTxts.RemoveAt(i);
                                //Added on 27-Aug-24
                                for (int j = i; j < lstTxts.Count; j++)
                                {
                                    if (lstTxts[j].StartsWith("-"))
                                    {
                                        note.AddText(lstTxts[j]);
                                        lstTxts.RemoveAt(j);
                                        j--;
                                    }
                                }
                                i--;
                            }
                            else
                            {
                                pos = 1;
                            }
                        }
                        oo.m_sText = Utility.CombineString(lstTxts, "\n");
                    }
                }

                //-------------------------------------------------------------------------
                //--------------------- FIND THE FINAL TEXT OF NOTE ---------------
                //-------------------------------------------------------------------------
                foreach (ProcedureInfo oo in lstDuplicateProcedureObjs)
                {
                    List<string> lstTxts = Utility.SplitString(oo.m_sText, "\n");
                    oo.m_sText = Utility.CombineString(lstTxts, " ");
                }


                //Add childs to  main procedure
                oMainProcedure.m_lstChilds = lstProcedureObjs;
                UpdateTheProcedureChilds(oMainProcedure);


                //find the complete text
                int poscnt = 0;
                foreach (ProcedureInfo oo in lstDuplicateProcedureObjs)
                {
                    string scompleteText = "";
                    int tablecnt = 0;
                    //find the position
                    if (oo.m_oParent != null)
                    {
                        oo.m_nPos = oo.m_oParent.m_lstChilds.IndexOf(oo) + 1;
                        for (int x = 0; x < oo.m_nPos; x++)
                        {
                            if (oo.m_oParent.m_lstChilds[x].m_sText.StartsWith("Table "))
                            {
                                //tablecnt = tablecnt + 1;//Modified by veeru on 2-Aug-23
                            }
                        }
                    }
                    else
                    {
                        poscnt = poscnt + 1;
                        oo.m_nPos = poscnt;
                    }
                    int pos = oo.m_nPos - tablecnt;
                    //---find the bullet
                    if (oo.m_nIndex == 0)
                    {
                        oo.m_sBullet = pos.ToString() + ".";
                    }
                    else if (oo.m_nIndex == 1)
                    {
                        oo.m_sBullet = Utility.GetBullet(64, pos) + ".";
                    }
                    else if (oo.m_nIndex % 2 == 0)
                    {
                        oo.m_sBullet = "(" + pos.ToString() + ")";
                    }
                    else if (oo.m_nIndex % 2 == 1)
                    {
                        oo.m_sBullet = "(" + Utility.GetBullet(96, pos) + ")";
                    }
                    //table
                    //if (oo.m_sText.StartsWith("Table "))//Modified by veeru on 2-Aug-23
                    //    oo.m_sBullet = "";
                    //find the complete text
                    FindCompleteText(oo, ref scompleteText);
                    oo.m_sFullText = scompleteText.Trim();
                }

                //Read ICN
               // oMainProcedure.m_lstICN = GetICNNumbers(htmlLink, outerHtml);
                //create procedure outputs
                List<string> lst = new List<string>();
                foreach (ProcedureInfo oo in lstDuplicateProcedureObjs)
                {
                    foreach (NoteInfo ox in oo.lstBeforeNotes)
                    {
                        lst.Add("".PadLeft(5 * oo.m_nIndex, ' ') + "".PadLeft(20, '-'));
                        lst.Add("".PadLeft(5 * oo.m_nIndex, ' ') + ox.m_eType.ToString() + " = " + ox.m_sText);
                        lst.Add("".PadLeft(5 * oo.m_nIndex, ' ') + "".PadLeft(20, '-'));
                    }
                    lst.Add("".PadLeft(5 * oo.m_nIndex, ' ') + oo.m_sBullet + " " + oo.m_sText);
                    foreach (NoteInfo ox in oo.lstAfterNotes)
                    {
                        lst.Add("".PadLeft(5 * oo.m_nIndex, ' ') + "".PadLeft(20, '*'));
                        lst.Add("".PadLeft(5 * oo.m_nIndex, ' ') + ox.m_eType.ToString() + " = " + ox.m_sText);
                        lst.Add("".PadLeft(5 * oo.m_nIndex, ' ') + "".PadLeft(20, '*'));
                    }
                }
                System.IO.File.WriteAllLines(Utility.m_sTempPath + System.IO.Path.GetFileNameWithoutExtension(htmlLink) + "_" + linktype + ".txt", lst.ToArray());
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            finally
            {
                
            }
            return oMainProcedure;
        }
        private void FindCompleteText(ProcedureInfo oprodure, ref string text)
        {
            try
            {
                //form the string
                foreach (NoteInfo oo in oprodure.lstBeforeNotes)
                {
                    text = text + oo.m_sText + " ";
                }
                text = text + oprodure.m_sText + " ";
                foreach (NoteInfo oo in oprodure.lstAfterNotes)
                {
                    text = text + oo.m_sText + " ";
                }
                //loop for children
                foreach (ProcedureInfo oo in oprodure.m_lstChilds)
                    FindCompleteText(oo, ref text);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        string GetElementText(string textz, ref List<string> lstpoints)
        {
            string bulletPoint = "";
            int type = 0;
            try
            {
                if (type == 0)
                {
                    //Replace--
                    string htmlTxt = textz;// liElement.InnerHtml;
                    foreach (KeyValuePair<string, string> entry in Utility.m_dtReplaceTexts)
                    {
                        htmlTxt = htmlTxt.Replace(entry.Key, entry.Value);
                    }

                    //update--
                    char[] array = htmlTxt.ToCharArray();
                    for (int i = 0; i < array.Length; i++)
                    {
                        if (array[i] == '>')
                        {
                            for (int j = i + 1; j < array.Length; j++)
                            {
                                if (array[j] == '<') break;
                                if (array[j] == '\n')
                                {
                                    array[j] = ' ';
                                }
                            }
                        }
                    }
                    string innerhtml = new string(array);
                    //---
                    //innerhtml = innerhtml.Replace("\n", " ");
                    List<string> lst = Utility.SplitStringWithoutTrim(innerhtml, "firstpara");
                    //before 12-May-26 veeru
                    //List<string> lstAll = new List<string>() { lst[0] };
                    //if (lst.Count >= 2) lstAll.Add(lst[1]);
                    //After 12-May-26 veeru
                    List<string> lstAll = lst;
                    foreach (string stTxt in lstAll)
                    {
                        bool isnote = false;
                        List<string> lst1 = Utility.SplitStringWithoutTrim(stTxt, ">");
                        //---- Update the list for  spaces....(21-Sept-25, Task : PW30K-A-72-21-0111-01A-609A-C, step 1)
                        for (int p = 1; p < lst1.Count; p++)
                        {
                            if (lst1[p].Length > 3 && lst1[p].EndsWith("</a"))
                            {
                                //lst1[p] = lst1[p].Substring(0, lst1[p].Length - 3) + " </a";
                            }
                        }

                        //----------------
                        for (int p = 1; p < lst1.Count; p++)
                        {
                            //-----
                            List<string> lst2 = Utility.SplitStringWithoutTrim(lst1[p], "<");
                            if (lst2.Count > 1)
                            {
                                //  ReplaceNextLineChar(ref lst2);//veeru 7-june-23
                                //Some times image headers are coming..avoid it
                                if (lst2[0].StartsWith("Figure 1:")) break;
                                //NOTE/CAUTION/WARNING
                                if (lst2[0] == "NOTE" || lst2[0] == "CAUTION" || lst2[0] == "WARNING")
                                {
                                    string textVal = "";
                                    //bulletPoint = bulletPoint + "\n" + lst2[0] + "\n";
                                    isnote = false;
                                    //get all the notes/warning text
                                    for (int q = p + 1; q < lst1.Count; q++)
                                    {
                                        List<string> lst3 = Utility.SplitStringWithoutTrim(lst1[q], "<");
                                        // ReplaceNextLineChar(ref lst3);//veeru 7-june-23

                                        if (lst3.Count > 1)
                                        {
                                            textVal = textVal + lst3[0];
                                            // textVal = textVal + " " + lst3[0] + " ";
                                            //bulletPoint = bulletPoint + " " + lst3[0] + " ";
                                            isnote = true;
                                        }
                                        if (lst1[q].Contains("/div"))
                                        {
                                            p = q;
                                            break;
                                        }
                                    }
                                    if (textVal.Length > 0)
                                    {
                                        bulletPoint = bulletPoint + "\n" + lst2[0] + "\n";
                                        bulletPoint = bulletPoint + "\n" + textVal + "\n";
                                    }
                                }
                                //FIRST PARA
                                else if (lst1[p - 1].ToUpper().Contains("firstpara".ToUpper()))
                                {
                                    // bulletPoint = bulletPoint + "\n" + lst2[0] + " ";
                                    bulletPoint = bulletPoint + "\n" + lst2[0];
                                    isnote = false;
                                }
                                //cautionText
                                else if (lst1[p - 1].ToUpper().Contains("cautionText".ToUpper()))
                                {
                                    isnote = true;
                                    //bulletPoint = bulletPoint + " " + lst2[0] + " ";
                                    bulletPoint = bulletPoint + lst2[0] + " ";
                                }
                                else
                                {
                                    if (isnote)
                                    {
                                        //bulletPoint = bulletPoint.Trim() + "\n" + lst2[0] + " ";
                                        bulletPoint = bulletPoint.Trim() + "\n" + lst2[0];
                                    }
                                    else if (lst2[0] == ".")
                                    {
                                        //bulletPoint = bulletPoint.Trim() + lst2[0] + " ";
                                        bulletPoint = bulletPoint.Trim() + lst2[0];
                                    }
                                    else
                                    {
                                        //bulletPoint = bulletPoint + " " + lst2[0] + " ";
                                        bulletPoint = bulletPoint + lst2[0];
                                    }
                                    isnote = false;
                                }
                            }
                        }
                        bulletPoint = bulletPoint + "\n";
                    }
                    string sx = "", sy = "";
                    string[] sArr = bulletPoint.Split(new string[1] { " " }, StringSplitOptions.None);
                    foreach (string ss in sArr)
                    {
                        string s3 = ss.Replace(" ", "");
                        if (s3.Length > 0)
                        {
                            sx = sx + s3 + " ";
                        }
                    }
                    sArr = sx.Split(new string[1] { "\n" }, StringSplitOptions.None);
                    foreach (string ss in sArr)
                    {
                        string s3 = ss.Replace("\n", "");
                        if (s3.Length > 0)
                        {
                            sy = sy + s3 + "\n";
                        }
                    }
                    bulletPoint = sy;
                }
                else
                {
                    bulletPoint = Utility.CombineString(Utility.SplitString(textz, "\r\n"), "\n");
                }
                lstpoints.Add("".PadLeft(100, '-'));
                lstpoints.Add(bulletPoint.Trim());
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return bulletPoint.Trim();
        }
        int getLevelIndex(string html, int parentLevel)
        {
            try
            {
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);

                // Case 1 : span steplevel exists
                var spanNode = doc.DocumentNode.SelectSingleNode("//span[contains(@class,'steplevel')]");

                if (spanNode != null)
                {
                    string className = spanNode.GetAttributeValue("class", "");

                    Match match = Regex.Match(className, @"steplevel-?(\d+)");

                    if (match.Success)
                    {
                        return Convert.ToInt32(match.Groups[1].Value);// + 1;
                    }
                }

                // Case 2 : nested ol steplevel0 exists inside note
                var olNode = doc.DocumentNode.SelectSingleNode("//ol[contains(@class,'steplevel')]");

                if (olNode != null)
                {
                    string className = olNode.GetAttributeValue("class", "");

                    Match match = Regex.Match(className, @"steplevel-?(\d+)");

                    if (match.Success)
                    {
                        int nestedLevel = Convert.ToInt32(match.Groups[1].Value) + 1;

                        // Add parent hierarchy
                        return parentLevel - 1 + nestedLevel;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return parentLevel;
        }
        private int GetNoteStartIndex(List<string> lstvals)
        {
            try
            {
                for (int i = 0; i < lstvals.Count; i++)
                {
                    int n1 = lstThings.IndexOf(lstvals[i]);
                    if (n1 == -1)
                    {
                        return i;
                    }
                    else
                    {
                        i = i + 1;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return 0;
        }
        private void UpdateTheProcedureChilds(ProcedureInfo oProc)
        {
            try
            {
                for (int i = 0; i < oProc.m_lstChilds.Count; i++)
                {
                    if (oProc.m_lstChilds[i].m_sText.Trim().Length == 0 && i > 0)
                    {
                        //Add the notes and warnings to lis
                        foreach (NoteInfo oo in oProc.m_lstChilds[i].lstBeforeNotes)
                            oProc.m_lstChilds[i - 1].lstBeforeNotes.Add(oo);
                        foreach (NoteInfo oo in oProc.m_lstChilds[i].lstAfterNotes)
                            oProc.m_lstChilds[i - 1].lstAfterNotes.Add(oo);
                        oProc.m_lstChilds.RemoveAt(i);
                        i--;
                    }
                }
                foreach (ProcedureInfo oo in oProc.m_lstChilds)
                    UpdateTheProcedureChilds(oo);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
    }
}