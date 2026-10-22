using EngineManualReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
namespace Cleaning_Records_Sheet
{
    class SPMReader
    {
        Dictionary<string, (EngineManualReader.ProcedureInfo, EngineManualReader.ProcedureInfo)> dtProcedures = new Dictionary<string, (EngineManualReader.ProcedureInfo, EngineManualReader.ProcedureInfo)>();
        Dictionary<string, string> dtDMC = new Dictionary<string, string>();
        Dictionary<string, List<ComparisonResult>> dtResults = new Dictionary<string, List<ComparisonResult>>();
        Dictionary<string, List<ComparisonResult>> dtResultsAll = new Dictionary<string, List<ComparisonResult>>();
        public string GetDMC(string spop)
        {
            if (dtDMC.ContainsKey(spop))
            {
                return dtDMC[spop];
            }
            return "";
        }

        public ComparisonResult GetStep(string spop, string step, StepComparionType etype)
        {
            try
            {
                if (step.Length > 0)
                {
                    string step1 = Regex.Replace(step, @"[().]", "");
                    Dictionary<string, List<ComparisonResult>> dtResultsx = dtResults;
                    if (etype == StepComparionType.Complete) dtResultsx = dtResultsAll;
                    //Compare---
                    if (dtResults.ContainsKey(spop))
                    {
                        foreach (ComparisonResult v in dtResults[spop])
                        {
                            if (v.m_old != null)
                            {
                                if (v.m_old.m_sCompleteBullet.ToLower() == step1.ToLower())
                                {
                                    return v;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return null;
        }
        public EngineManualReader.ProcedureInfo GetProcedure(string spop, string type)
        {
            try
            {
                if (dtProcedures.ContainsKey(spop))
                {
                    if (type.ToUpper() == "OLD")
                    {
                        return dtProcedures[spop].Item1;
                    }
                    else
                    {
                        return dtProcedures[spop].Item2;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return null;
        }

        public bool IsDMCSame(string spop)
        {
            try
            {
                if (dtResults.ContainsKey(spop))
                {
                    int cnt = 0;
                    foreach (ComparisonResult v in dtResults[spop])
                    {
                        if (v.m_eCompareType == CompareType.SAME) cnt = cnt + 1;
                    }
                    if (cnt == dtResults[spop].Count) return true;
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
        public void ReadSPM(string sOldURL, string sNewUrl, List<string> lstSpops, List<string> lstSpopsTitle)//string sXlPath)
        {
            EngineManualReader.SeleniumReader oseleniumReader = new EngineManualReader.SeleniumReader();
            try
            {
                //------------- READ Cleaning ticket excel file to find the sheet names -------
                //Use epplus to read the data
                //List<string> lstSpops = ReadSPOPFromCleaningSheet(sXlPath);

                //upate spop text
                for (int i = 0; i < lstSpops.Count; i++)
                {
                    string number = Regex.Replace(lstSpops[i], @"[^0-9]", "");
                    if (number.Length > 0)
                    {
                        lstSpops[i] = "SPOP " + number;
                    }
                    else
                    {
                        lstSpops.RemoveAt(i);
                        lstSpopsTitle.RemoveAt(i);
                        i--;
                    }
                }
                //Remove duplicates 
                for (int i = 0; i < lstSpops.Count; i++)
                {
                    for (int j = i + 1; j < lstSpops.Count; j++)
                    {
                        if (lstSpops[i] == lstSpops[j])
                        {
                            lstSpops.RemoveAt(j);
                            lstSpopsTitle.RemoveAt(j);
                            j--;
                        }
                    }
                }
                //lstSpops = lstSpops.Distinct().ToList();

                //------------- READ the Engine Manual data from Old Revision --- 
                //foreach (string spop in lstSpops)
                for (int i = 0; i < lstSpops.Count; i++)
                {
                    string spop = lstSpops[i];
                    SolrSearch oo = new SolrSearch();
                    string dmc = oo.SearchTask(sOldURL, spop, lstSpopsTitle[i]);
                    if (dmc.Length > 0)
                    {
                        dtDMC.Add(spop, dmc);
                        string url = sOldURL + @"/" + dmc + ".html";
                        EngineManualReader.EMReader emReader = new EngineManualReader.EMReader();
                        EngineManualReader.ProcedureInfo oprocOld = DefineProcedure(emReader.Read(url, "old", oseleniumReader));
                        //New read
                        url = sNewUrl + @"/" + dmc + ".html";
                        emReader = new EngineManualReader.EMReader();
                        EngineManualReader.ProcedureInfo oprocNew = DefineProcedure(emReader.Read(url, "new", oseleniumReader));
                        dtProcedures.Add(spop, (oprocOld, oprocNew));

                        //compare specific
                        ComparisonHelper comHelper = new ComparisonHelper();
                        List<ComparisonResult> lstCompares = comHelper.Compare(oprocOld, oprocNew, StepComparionType.Specific);
                        dtResults.Add(spop, lstCompares);

                        //Compare all
                        comHelper = new ComparisonHelper();
                        lstCompares = comHelper.Compare(oprocOld, oprocNew, StepComparionType.Complete);
                        dtResultsAll.Add(spop, lstCompares);
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            finally
            {
                oseleniumReader.close();
                oseleniumReader = null;
            }
        }
        EngineManualReader.ProcedureInfo DefineProcedure(EngineManualReader.ProcedureInfo oproc)
        {
            //In SPM the main "procedure" is defined as child procedure, in this case update it
            EngineManualReader.ProcedureInfo finalProc = oproc, findProc = null;
            foreach (EngineManualReader.ProcedureInfo oo in oproc.m_lstChilds)
            {
                if (oo.m_sText.ToUpper() == "Procedure".ToUpper())
                {
                    findProc = oo; break;
                }
            }
            findProc = null;
            //If found update the labels and main procedure
            if (findProc != null)
            {
                finalProc = findProc;
                List<EngineManualReader.ProcedureInfo> lstAll = new List<EngineManualReader.ProcedureInfo>();
                Recursive(finalProc, ref lstAll);
                foreach (var item in lstAll)
                {
                    UpdateLevel(item);
                }
            }
            //Find complete bullet
            List<EngineManualReader.ProcedureInfo> lstAllProcs = new List<EngineManualReader.ProcedureInfo>();
            Recursive(finalProc, ref lstAllProcs);
            foreach (var item in lstAllProcs)
            {
                string CompleteBulltet = "";
                Recursive(item, ref CompleteBulltet);
                item.m_sCompleteBullet = CompleteBulltet;
            }
            RecursiveUpdateIndex(finalProc, 0);
            return finalProc;
        }
        private void RecursiveUpdateIndex(EngineManualReader.ProcedureInfo proc, int pos)
        {
            proc.m_nIndex = pos;
            foreach (EngineManualReader.ProcedureInfo child in proc.m_lstChilds)
            {
                RecursiveUpdateIndex(child, pos + 1);
            }
        }
        private void Recursive(EngineManualReader.ProcedureInfo proc, ref string CompleteBulltet)
        {
            if (proc != null)
            {
                CompleteBulltet = proc.m_sBullet.Replace("(", "").Replace(")", "").Replace(".", "").Trim() + CompleteBulltet;
                Recursive(proc.m_oParent, ref CompleteBulltet);
            }
        }
        private void Recursive(EngineManualReader.ProcedureInfo proc, ref List<EngineManualReader.ProcedureInfo> lstAll)
        {
            lstAll.Add(proc);
            foreach (EngineManualReader.ProcedureInfo child in proc.m_lstChilds)
                Recursive(child, ref lstAll);
        }
        private void UpdateLevel(EngineManualReader.ProcedureInfo oproc)
        {
            string bulletType = "bracket";
            string text = oproc.m_sBullet.Replace("(", "").Replace(")", "").Replace(".", "").Trim();
            int number = GetNumber(text);
            string newBullet = "";
            //upper the level
            if (oproc.m_nIndex == 0)
            {
                newBullet = "";
            }
            else if (oproc.m_nIndex == 1)
            {
                newBullet = number.ToString();
                bulletType = ".";
            }
            else if (oproc.m_nIndex == 2)
            {
                newBullet = NumberToColumn(number);
                bulletType = ".";
            }
            else if (oproc.m_nIndex % 2 == 1)
            {
                newBullet = number.ToString();
            }
            else if (oproc.m_nIndex > 2)
            {
                newBullet = NumberToColumn(number).ToLower();
            }
            //Add the . or ()
            if (newBullet.Length > 0)
            {
                if (bulletType == "bracket") newBullet = "(" + newBullet + ")";
                else newBullet = newBullet + ".";
            }
            oproc.m_sBullet = newBullet;
            //oproc.m_nIndex = oproc.m_nIndex - 1;
        }
        private int GetNumber(string str)
        {
            int number = 0;
            if (int.TryParse(str, out number))
            {
                return number;
            }
            else
            {
                int result = 0;
                foreach (char c in str.ToUpper())
                {
                    if (c < 'A' || c > 'Z')
                        throw new ArgumentException("Invalid column name.");

                    result = result * 26 + (c - 'A' + 1);
                }
                return result;
            }
        }
        string NumberToColumn(int number)
        {
            string result = "";

            while (number > 0)
            {
                number--;
                result = (char)('A' + (number % 26)) + result;
                number /= 26;
            }
            return result;
        }
        public static List<string> ReadSPOPFromCleaningSheet(string sXlPath)
        {
            List<string> lstSpops = new List<string>();
            try
            {
                string newfile = Utility.CopyFiletoLocal(sXlPath);
                FileInfo oFileInfo = new FileInfo(newfile);
                ExcelPackage oExcel = new ExcelPackage(oFileInfo);
                //go through all the sheets
                for (int i = 1; i <= oExcel.Workbook.Worksheets.Count; i++)
                {
                    ExcelWorksheet oXlWorkSheet = oExcel.Workbook.Worksheets[i];
                    if (oXlWorkSheet.Dimension == null) continue;
                    if (oXlWorkSheet.Hidden == eWorkSheetHidden.Visible)
                    {
                        if (oXlWorkSheet.Name.StartsWith("SPOP"))
                        {
                            string num = "SPOP " + Utility.SplitString(oXlWorkSheet.Name.Replace("SPOP", ""), " ")[0];
                            if (lstSpops.Contains(num) == false)
                            {
                                if (num == "SPOP 15")
                                    lstSpops.Add(num);
                            }
                        }
                    }
                }
                //Delete and close
                oExcel.Dispose();
                Utility.DeleteFile(newfile);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lstSpops;
        }
    }
    public class SolrSearch
    {
        public List<EngineInfo> Search(string hmtlpath, string searchtext)
        {
            List<EngineInfo> lstEngineSns = new List<EngineInfo>();
            try
            {


                //Form the search link
                string http = @"http://";
                List<string> lsts = Utility.SplitString(hmtlpath.Substring(http.Length, hmtlpath.Length - http.Length), "/");
                string s1 = lsts[0];
                string s2 = lsts[1];
                string linkPath = http + lsts[0] + "/solr/" + lsts[1] + "/select?q=" + searchtext + "&rows=500";
                string smorelinks = http + lsts[0] + "/solr/" + lsts[1];
                string htmlPath = http + lsts[0];
                //Now time for find the data from the link
                string contents = Utility.GetURlData(linkPath);
                //Get the hyper links
                List<string> lsta = Utility.SplitString(contents, "</a>");
                for (int i = 1; i < lsta.Count; i++)
                {
                    if (lsta[i].Contains("Documents Found)"))
                    {
                        List<string> ls111 = Utility.SplitString(lsta[i - 1], ">");
                        string title = ls111[ls111.Count - 1].Trim();
                        //no of documents
                        List<string> ls222 = Utility.SplitString(Utility.SplitString(lsta[i], "Documents Found)")[0], "(");
                        //create object
                        EngineInfo oengine = new EngineInfo(title, ls222[ls222.Count - 1]);
                        lstEngineSns.Add(oengine);
                        //
                        if (oengine.m_nDocuments <= 5)
                        {
                            for (int j = i; j < i + oengine.m_nDocuments; j++)
                            {
                                string sx = "";
                                List<string> lstaa = Utility.SplitString(lsta[j], "<a");
                                if (lstaa.Count == 1) sx = lstaa[0];
                                else sx = lstaa[1];
                                //--Add the task
                                TaskInfo otask = new TaskInfo(GetTitle(sx));
                                oengine.m_lstTasks.Add(otask);
                                if (otask.m_nSearchcnt > 0) oengine.m_nSearchcnt = oengine.m_nSearchcnt + 1;
                                //find html link
                                List<string> lst1 = Utility.SplitString(sx, "href=");
                                List<string> lst2 = Utility.SplitString(lst1[1], "\"");
                                otask.m_sLink = htmlPath + lst2[0];
                            }
                            i = i + oengine.m_nDocuments - 1;
                        }
                        else
                        {
                            //---Find the number of pages
                            int left = oengine.m_nDocuments - 10;
                            int ncycles = left / 10; ;
                            if (left % 10 != 0)
                            {
                                ncycles = ncycles + 1;
                            }
                            //----Get the links
                            List<string> lstlinks = new List<string>();
                            //first page
                            string mylnk = oengine.m_sTitle.Replace(" ", "%20");
                            mylnk = mylnk.Replace("/", "%2F").Trim();
                            string s10link = smorelinks + "/select?q=" + searchtext + "&fq=manual:" + "\"" + mylnk + "\"" + "&group=false";
                            lstlinks.Add(s10link);
                            //next pages
                            for (int p = 0; p < ncycles; p++)
                            {
                                string snextlink = smorelinks + "/select?q=" + searchtext + "&fq=manual%3A%22" + mylnk + "%22&start=" + (10 * (p + 1)).ToString() + "&group=false";
                                lstlinks.Add(snextlink);
                            }

                            //-----Get the data
                            foreach (string link in lstlinks)
                            {
                                string contentsNew = Utility.GetURlData(link);
                                List<string> lp = new List<string>();
                                List<string> lstanew = Utility.SplitString(contentsNew, "<div class=\"searchResultDocument\">");
                                foreach (string sx in lstanew)
                                {
                                    if (sx.Contains("searchResultDocumentID"))
                                    {
                                        //--Add the task
                                        TaskInfo otask = new TaskInfo(GetTitle(sx));
                                        oengine.m_lstTasks.Add(otask);
                                        if (otask.m_nSearchcnt > 0) oengine.m_nSearchcnt = oengine.m_nSearchcnt + 1;
                                        //find html link
                                        List<string> lst1 = Utility.SplitString(sx, "href=");
                                        List<string> lst2 = Utility.SplitString(lst1[1], "\"");
                                        otask.m_sLink = htmlPath + lst2[0];
                                    }
                                }
                            }
                            i = i + 5 - 1;
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lstEngineSns;
        }
        public string SearchTask(string hmtlpath, string searchtext, string Recordtitle)
        {
            string DMClink = "";
            try
            {
                List<EngineInfo> lstEngineSns = Search(hmtlpath, searchtext);
                foreach (var eng in lstEngineSns)
                {
                    if (eng.m_sTitle.Contains("Standard Practices Manual"))
                    {
                        //find the title in the task
                        foreach (var task in eng.m_lstTasks)
                        {
                            if (task.m_sTitleText.Contains(searchtext))
                            {
                                if (bIsTitleMatch(task.m_sTitleText, Recordtitle))
                                {
                                    if (bIsProcedureTask(task.m_sLink))
                                    {
                                        DMClink = System.IO.Path.GetFileNameWithoutExtension(task.m_sLink);
                                        break;
                                    }
                                }
                            }
                        }
                        break;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return DMClink;
        }
        bool bIsTitleMatch(string EMtitle, string recordSheetTitle)
        {
            if (EMtitle.Replace(" ", "").ToUpper().Trim().Contains(recordSheetTitle.Replace(" ", "").ToUpper().Trim()))
                return true;
            return false;
        }
        private bool bIsProcedureTask(string url)
        {
            try
            {
                string htmlText = Utility.GetURlData(url);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(htmlText);
                string text = doc.DocumentNode.InnerText;
                if (text.Contains("DM Applicability"))
                {
                    if (text.Contains("References"))
                    {
                        if (text.Contains("Common Information"))
                        {
                            //if (text.Contains("Preliminary Requirements"))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog("", ee.Message, ee.StackTrace);
            }
            return false;
        }
        private string GetTitle(string sx)
        {
            try
            {
                string sp = sx.Replace("<em>", "").Replace("</em>", "").Replace("\r\n", "").Trim();
                List<string> l1 = Utility.SplitString(sp, "href");
                List<string> l2 = Utility.SplitString(l1[1], ">");
                List<string> l3 = Utility.SplitString(l2[1], "<");
                return l3[0].Trim();
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return "";
        }
    }
    public class TaskInfo
    {
        public string m_sDMCFromPage = "";
        public string m_sTitleFromPage = "";
        public string m_sIssueNoFromPage = "";
        public string m_sIssuedateFromPage = "";


        public string m_sTitleText = "";
        public string m_sLink = "";
        public string m_sDMC = "";
        public int m_nSearchcnt = 0;

        public string m_sItem = "";
        public string m_sScheme = "";
        public TaskInfo(string txt)
        {
            m_sTitleText = txt;
        }
        public TaskInfo(string txt, string searchtext, bool bIsTitleContains)
        {
            m_sTitleText = txt;
            if (bIsTitleContains)
            {
                if (m_sTitleText.ToUpper().Contains(searchtext.ToUpper()))
                {
                    m_nSearchcnt = m_nSearchcnt + 1;
                }
            }
            else
            {
                m_nSearchcnt = m_nSearchcnt + 1;
            }
        }
    }
    public class EngineInfo
    {
        public string m_sTitle = "";
        public int m_nDocuments = 0;
        public int m_nSearchcnt = 0;
        public List<TaskInfo> m_lstTasks = new List<TaskInfo>();
        public EngineInfo(string s1, string s2)
        {
            m_sTitle = s1;
            m_nDocuments = Utility.GetInteger(s2);
        }
    }
}
