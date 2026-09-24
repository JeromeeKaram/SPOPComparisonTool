using Cleaning_Records_Sheet;
using Cleaning_Records_Sheet.About;
using EngineManualReader;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SPOPComparisonTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtOldRevisionLink.Text))
            {
                MessageBox.Show("Please enter Old Engine manual Link");
                return;
            }

            if (string.IsNullOrEmpty(txtNewRevisionLink.Text))
            {
                MessageBox.Show("Please enter New Engine manual Link");
                return;
            }

            if (string.IsNullOrEmpty(txtCleaningRecord.Text))
            {
                MessageBox.Show("Please select cleaning record");
                return;
            }

            try
            {
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.MarqueeAnimationSpeed = 30;
                progressBar1.Visible = true;

                ReadExcel(txtCleaningRecord.Text);
            }
            finally
            {
                progressBar1.Visible = false;
            }
        }

        private void btnBrowseDocument_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Title = "Select Excel File";
            dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";
            dialog.Multiselect = false;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtCleaningRecord.Text = dialog.FileName;
            }

            dialog.Dispose();
        }

        private void ReadExcel(string filePath)
        {
            var workSheetNames = new List<string>();
            var highlights = new List<Highlight>();
            var excelData = new List<ExcelData>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var seen = new HashSet<string>();

                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    var match = Regex.Match(worksheet.Name, @"^(SPOP\d+)");

                    if (match.Success)
                    {
                        string spopKey = match.Groups[1].Value;

                        if (seen.Add(spopKey)) // first occurrence only
                        {
                            workSheetNames.Add(worksheet.Name);
                        }
                    }
                }

                //foreach (var name in workSheetNames)
                {
                    var testWorkSheet = package.Workbook.Worksheets[7];

                    int rowCount = testWorkSheet.Dimension.Rows;

                    int rowStartsAt = 0;

                    for (int row = 1; row <= testWorkSheet.Dimension.Rows; row++)
                    {
                        string value = testWorkSheet.Cells[row, 2].Text;

                        if (value != null)
                        {
                            value = value.Replace("<br>", " ")
                                         .Replace("<br/>", " ")
                                         .Replace("<br />", " ");

                            if (value.Contains("Para.") && value.Contains("No."))
                            {
                                rowStartsAt = row;
                                break;
                            }
                        }
                    }

                    if (rowStartsAt > 0)
                    {
                        var mergedRanges = GetMergedRowsForColumnB(testWorkSheet, rowStartsAt);

                        foreach (var mergedRange in mergedRanges)
                        {
                            var startRow = mergedRange.StartRow;
                            var endRow = mergedRange.EndRow;

                            var manualRequirements = new List<string>();

                            for (var i = startRow; i <= endRow; i++)
                            {
                                manualRequirements.Add(testWorkSheet.Cells[i, 7].Text);
                            }

                            var excelRecord = new ExcelData();
                            excelRecord.ManualRequirement = manualRequirements;
                            excelRecord.SPOPName = testWorkSheet.Name;
                            excelRecord.Row = startRow;

                            // Text of the merged B cell
                            excelRecord.BulletPoint = testWorkSheet.Cells[startRow, 2].Text;

                            excelData.Add(excelRecord);

                        }

                        //Integrate Veerendra's Code

                        var reader = GetReader();

                        foreach (var record in excelData)
                        {
                            if (record.BulletPoint != "1.D") continue;

                            string withoutJapanese = record.SPOPName.Split(' ')[0];
                            string spaceBetweenSPOPandNumber = Regex.Replace(withoutJapanese, @"^(SPOP)(\d+)$", "$1 $2");

                            var comparisonResult = GetComparisonResult(reader, spaceBetweenSPOPandNumber, record.BulletPoint);

                            if (comparisonResult != null && comparisonResult.m_eCompareType == CompareType.SAME)
                            {
                                continue;
                            }

                            var newRevisionBulletData = comparisonResult.m_new;
                            var bulletText = comparisonResult.m_new.m_sText;

                            var revisedManualRequirement = new List<string>();

                            var highLight = new Highlight();

                            var manualRequirementRow = record.Row;

                            foreach (var item in record.ManualRequirement)
                            {
                                if (!ManualRequirementExistsInNewRevision(bulletText, item))
                                {
                                    revisedManualRequirement.Add($"{item}_{manualRequirementRow}_7");
                                }
                                manualRequirementRow = manualRequirementRow + 1;
                            }

                            if (revisedManualRequirement.Any())
                            {
                                highLight.Row = record.Row;
                                highLight.SPOPName = record.SPOPName;
                                highLight.Columns2Hightlight = revisedManualRequirement;
                                highlights.Add(highLight);
                            }
                        }


                        //        //Do the comparison here and create the HighLight Object
                        //        var h = new Highlight();
                        //        h.SheetName = "SPOP15 表";
                        //        h.Row = 18;
                        //        h.Columns = new List<int>() { 7 };
                        //        highlights.Add(h);
                    }
                }
            }

            HighlightManualRequirement(highlights, Color.Yellow);

            // Open the Excel file
            Process.Start(new ProcessStartInfo(txtCleaningRecord.Text)
            {
                UseShellExecute = true
            });
        }

        private bool ManualRequirementExistsInNewRevision(string bulletText, string manualRequirement)
        {
            if (!string.IsNullOrEmpty(bulletText) && bulletText.Contains(manualRequirement))
            {
                return true;
            }
            return false;
        }

        //private (int merged, int mergedStartRow, int mergedEndRow, int mergedRowCount) GetMergedCellLastRow(ExcelWorksheet worksheet, int row, int column)
        //{
        //    foreach (string rangeAddress in worksheet.MergedCells)
        //    {
        //        ExcelRange mergedRange = worksheet.Cells[rangeAddress];

        //        if (row >= mergedRange.Start.Row &&
        //            row <= mergedRange.End.Row &&
        //            column >= mergedRange.Start.Column &&
        //            column <= mergedRange.End.Column)
        //        {
        //            int rowCount = mergedRange.End.Row - mergedRange.Start.Row + 1;
        //            return (1, mergedRange.Start.Row, mergedRange.End.Row, rowCount);
        //        }
        //    }

        //    return (0, row, row, 1);
        //}

        private void Form1_Load(object sender, EventArgs e)
        {

#if DEBUG
            txtCleaningRecord.Text = "D:\\CleaningRecord\\PW1100G-JM_Cleaning_Record_in_Tsurugashima_Rev.19_2026.3.9.xlsx";
            txtOldRevisionLink.Text = "http://127.0.0.1:8000/PW1000G-77445-19453-00/";
            txtNewRevisionLink.Text = "http://127.0.0.1:8001/PW1000G-77445-19453-00/";
#endif

            lblHeader.Text = Cleaning_Records_Sheet.Utility.m_sToolName + " V" + Cleaning_Records_Sheet.Utility.m_sVersion;
            this.Text = Cleaning_Records_Sheet.Utility.m_sToolName + " V" + Cleaning_Records_Sheet.Utility.m_sVersion;
        }

        private void HighlightManualRequirement(List<Highlight> highLight, Color color)
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(txtCleaningRecord.Text)))
            {
                foreach (var hl in highLight)
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[hl.SPOPName];

                    foreach (var column in hl.Columns2Hightlight)
                    {
                        var rowIndex = int.Parse(column.Split('_')[1]);
                        var columnIndex = int.Parse(column.Split('_')[2]);

                        worksheet.Cells[rowIndex, columnIndex, rowIndex, columnIndex + 1].Style.Fill.PatternType =
                            ExcelFillStyle.Solid;

                        worksheet.Cells[rowIndex, columnIndex, rowIndex, columnIndex + 1].Style.Fill.BackgroundColor
                            .SetColor(color);
                    }
                }
                excelPackage.Save();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dlg = new UIAbout();
            dlg.ShowDialog();
        }

        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dlg = new UIHelp();
            dlg.ShowDialog();
        }

        private List<(int StartRow, int EndRow)> GetMergedRowsForColumnB(ExcelWorksheet worksheet, int startSearchRow)
        {
            var result = new List<(int StartRow, int EndRow)>();

            foreach (string rangeAddress in worksheet.MergedCells)
            {
                ExcelRange mergedRange = worksheet.Cells[rangeAddress];

                // Must involve Column B
                bool containsColumnB =
                    mergedRange.Start.Column <= 2 &&
                    mergedRange.End.Column >= 2;

                // Must start after the specified row
                bool startsAfterRow =
                    mergedRange.Start.Row > startSearchRow;

                if (containsColumnB && startsAfterRow)
                {
                    result.Add((
                        mergedRange.Start.Row,
                        mergedRange.End.Row));
                }
            }

            return result;
        }

        private ComparisonResult GetComparisonResult(SPMReader reader, string spop, string bullet)
        {
            var result = reader.GetStep(spop, bullet);
            return result;
        }


        private SPMReader GetReader()
        {
            var reader = new SPMReader();
            reader.ReadSPM(txtOldRevisionLink.Text, txtNewRevisionLink.Text, txtCleaningRecord.Text);
            return reader;
        }
    }
}
