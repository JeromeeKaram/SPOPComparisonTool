using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SVC
{
    class XlUtility
    {
        static object[,] oShData = null;
        static int m_nXlShRows = 1;
        static int m_nXlnCols = 1;
        public static ExcelRange SetCellValue(ExcelWorksheet osh, string value, int FromRow, int FromCol, int ToRow, int ToCol, bool bIsBold, bool bIsMerge)
        {
            ExcelRange rng = null;
            try
            {
                rng = osh.Cells[FromRow, FromCol];
                //Value
                rng.Value = value;
                //Bold
                if (bIsBold)
                {
                    rng.Style.Font.Bold = true;
                }
                //Merge
                rng = osh.Cells[FromRow, FromCol, ToRow, ToCol];
                if (bIsMerge)
                {
                    rng.Merge = true;
                }
                //Wrap
                rng.Style.WrapText = true;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return rng;
        }
        public static void WriteExcelCellValueWithBackColor(ExcelRange rng, Color clr)
        {
            try
            {
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(clr);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public static void WriteExcelCellValueWithForeColor(ExcelRange rng, Color clr)
        {
            try
            {
                rng.Style.Font.Color.SetColor(clr);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public static void WriteExcelCellValueWithForeColor(ExcelRange rng, string value, Color clr)
        {
            try
            {
                if (value.Length > 0) rng.Value = value;
                rng.Style.Font.Color.SetColor(clr);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public static void WriteExcelCellValueWithBackColor(ExcelRange rng, string type)
        {
            try
            {
                ColorInfo clr = Utility.m_dtColorCode[type];
                WriteExcelCellValueWithBackColor(rng, clr.m_oBackColor);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }

        public static void WriteExcelCellValueWithForeColor(ExcelRange rng, string type)
        {
            try
            {
                ColorInfo clr = Utility.m_dtColorCode[type];
                WriteExcelCellValueWithForeColor(rng, clr.m_oForeColor);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public static void WriteExcelCellValueWithBackForeColor(ExcelRange rng, string type)
        {
            try
            {
                ColorInfo clr = Utility.m_dtColorCode[type];
                WriteExcelCellValueWithBackColor(rng, clr.m_oBackColor);
                WriteExcelCellValueWithForeColor(rng, clr.m_oForeColor);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        public static void WriteExcelCellValueWithBackForeColor(ExcelRange rng, string value, string type)
        {
            try
            {
                ColorInfo clr = Utility.m_dtColorCode[type];
                if (value.Length > 0) rng.Value = value;
                WriteExcelCellValueWithBackColor(rng, clr.m_oBackColor);
                WriteExcelCellValueWithForeColor(rng, clr.m_oForeColor);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }

        public static void SetRowHeight(ExcelWorksheet osh, int rownum, double rwHeight)
        {
            if (osh == null) return;
            try
            {
                if (rwHeight > 408) rwHeight = 408;
                osh.Row(rownum).Height = rwHeight;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        //------------------------ DONE -------------------



        public static string GetCellValueNoTrim(ExcelWorksheet oXlWorkSheet, int nRow, int nCol)
        {
            string sValue = "";
            try
            {
                if (nRow > 0 && nCol > 0)
                {
                    object oCell = oXlWorkSheet.Cells[nRow, nCol].Value;
                    if (oCell != null)
                    {
                        sValue = oCell.ToString();
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return sValue;
        }
        public static string GetCellValue(ExcelWorksheet oXlWorkSheet, int nRow, int nCol)
        {
            if (oXlWorkSheet == null) return "";
            string sValue = "";
            try
            {
                if (nRow > 0 && nCol > 0)
                {
                    object oCell = oXlWorkSheet.Cells[nRow, nCol].Value;
                    if (oCell != null)
                    {
                        sValue = oCell.ToString().Trim();
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return sValue;
        }
        public static string GetCellValue(int nRow, int nCol)
        {
            string sValue = "";
            try
            {
                if (oShData != null)
                {
                    if (nRow <= m_nXlShRows && nCol <= m_nXlnCols)
                    {
                        object oValue = oShData[nRow, nCol];
                        if (oValue != null)
                        {
                            sValue = oValue.ToString().Trim();
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return sValue;
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
    }
    class ExtractExcelAttributes
    {
        static object[,] oShData = null;
        static int m_nXlShRows = 1;
        static int m_nXlnCols = 1;

        public static List<string> SortIt(List<string> lstFleetcareTasks, List<string> lstXlTasks, int len)
        {
            List<string> lstCommonTasks = new List<string>();
            try
            {
                List<string> lstNonTasks = new List<string>();
                foreach (string sxlTask in lstXlTasks)
                {
                    string sTask = "";
                    if (sxlTask.Length >= len)
                        sTask = sxlTask.Remove(len, sxlTask.Length - len);
                    if (sTask.Length != 0)
                    {
                        if (lstCommonTasks.Contains(sTask) == false)
                            lstCommonTasks.Add(sTask);
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return lstCommonTasks;
        }
        public static string GetCellValue(ExcelWorksheet oXlWorkSheet, int nRow, int nCol)
        {
            string sValue = "";
            try
            {
                if (nRow > 0 && nCol > 0)
                {
                    object oCell = oXlWorkSheet.Cells[nRow, nCol].Value;
                    if (oCell != null)
                    {
                        sValue = oCell.ToString().Trim();
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return sValue;
        }

        public static void releaseExcelObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch// (Exception ee)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        public static List<string> SplitString(string sValue, string[] sDelimeter)
        {
            List<string> lst = new List<string>();
            try
            {
                string[] sArr = sValue.Split(sDelimeter, StringSplitOptions.RemoveEmptyEntries);
                foreach (string ss in sArr)
                {
                    if (ss.Trim().Length > 0)
                    {
                        lst.Add(ss.Trim());
                    }
                }
            }
            catch// (Exception ee)
            {
                //Utility.WriteErrorLog(ee);
            }
            return lst;
        }
        public static string GetCellValue(int nRow, int nCol)
        {
            string sValue = "";
            try
            {
                if (nRow <= m_nXlShRows && nCol <= m_nXlnCols)
                {
                    object oValue = oShData[nRow, nCol];
                    if (oValue != null)
                    {
                        sValue = oValue.ToString();
                    }
                }
            }
            catch// (Exception ee)
            {
                //Utility.WriteErrorLog(ee);
            }
            return sValue;
        }
    }
    public class ExcelExportData
    {
        public string taskNO;
        public string titleUpdate;
        public string procedureUpdate;
        public string impact = "";
    }

    public class ProcedureChangesInfo
    {
        public int No;
        public List<string> oldRev;
        public List<string> newRev;
    }
    public class ChangedProcedures
    {
        public string oldPnt;
        public string newPnt;
        public string color;
    }
    public class ColorInfo
    {
        public System.Drawing.Color m_oBackColor = System.Drawing.Color.FromArgb(255, 255, 255);
        public System.Drawing.Color m_oForeColor = System.Drawing.Color.FromArgb(0, 0, 0);
        public System.Drawing.Color m_oHighlightColor = System.Drawing.Color.FromArgb(0, 0, 0);
        public ColorInfo() { }
        public ColorInfo(System.Drawing.Color cl1, System.Drawing.Color cl2)
        {
            m_oBackColor = cl1;
            m_oForeColor = cl2;
            m_oHighlightColor = m_oForeColor;
        }
    }
}
