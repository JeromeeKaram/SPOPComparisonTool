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

        private async void btnRun_Click(object sender, EventArgs e)
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

                await Task.Run(() =>
                {
                    ReadExcel(txtCleaningRecord.Text);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            var changedManualRequirements = new List<Highlight>();
            var changedBulletPoints = new List<Highlight>();
            var excelData = new List<ExcelData>();

            var reader = GetReader();

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

                foreach (var name in workSheetNames)
                {

                    var testWorkSheet = package.Workbook.Worksheets[name];

                    //if (testWorkSheet.Name != "SPOP15 表") continue; //Comment this line to process all worksheets

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
                            var excelRecord = new ExcelData();
                            var startRow = mergedRange.StartRow;
                            var endRow = mergedRange.EndRow;

                            var manualRequirements = new List<string>();

                            for (var i = startRow; i <= endRow; i++)
                            {
                                manualRequirements.Add(testWorkSheet.Cells[i, 7].Text);
                            }

                            excelRecord.ManualRequirement = manualRequirements;
                            excelRecord.SPOPName = testWorkSheet.Name;
                            excelRecord.Row = startRow;

                            // Text of the merged B cell
                            excelRecord.BulletPoint = Regex.Replace(testWorkSheet.Cells[startRow, 2].Text, @"\r?\n\s*", "");

                            excelData.Add(excelRecord);
                        }
                    }
                }
            }

            foreach (var record in excelData)
            {
                //if (record.BulletPoint != "1.D") continue;

                string withoutJapanese = record.SPOPName.Split(' ')[0];
                string spaceBetweenSPOPandNumber = Regex.Replace(withoutJapanese, @"^(SPOP)(\d+)$", "$1 $2");

                //if (spaceBetweenSPOPandNumber != "SPOP 225") continue;

                var comparisonResult = GetComparisonResult(reader, spaceBetweenSPOPandNumber, record.BulletPoint);

                if (comparisonResult == null) continue;
                var bulletInfo = comparisonResult.m_new.m_sBullet;

                if (!string.IsNullOrEmpty(bulletInfo))
                {
                    var highLight = new Highlight();
                    highLight.Row = record.Row;
                    highLight.SPOPName = record.SPOPName;
                    highLight.Columns2Highlight = new List<string> { $"{record.BulletPoint}_{record.Row}_2" };
                    changedBulletPoints.Add(highLight);
                }

                if (comparisonResult != null && comparisonResult.m_eCompareType == CompareType.SAME)
                {
                    continue;
                }

                var newRevisionBulletData = comparisonResult.m_new;
                var bulletText = comparisonResult.m_new.m_sText;

                var highLightManualRequirement = new List<string>();

                var manualRequirementRow = record.Row;

                foreach (var item in record.ManualRequirement)
                {
                    var splitItem = item.Contains(":") ? item.Split(':')[1].Trim() : item.Trim();
                    var variants = GetVariants(splitItem);

                    if (!ManualRequirementExistsInNewRevision(bulletText, variants))
                    {
                        highLightManualRequirement.Add($"{splitItem}_{manualRequirementRow}_7");
                    }
                    manualRequirementRow = manualRequirementRow + 1;
                }

                if (highLightManualRequirement.Any())
                {
                    var highLight = new Highlight();
                    highLight.Row = record.Row;
                    highLight.SPOPName = record.SPOPName;
                    highLight.Columns2Highlight = highLightManualRequirement;
                    changedManualRequirements.Add(highLight);
                }
            }

            HighlightChanges(changedManualRequirements, changedBulletPoints, Color.Yellow);

            // Open the Excel file
            Process.Start(new ProcessStartInfo(txtCleaningRecord.Text)
            {
                UseShellExecute = true
            });
        }

        private bool ManualRequirementExistsInNewRevision(
    string bulletText,
    List<string> variants)
        {
            if (string.IsNullOrWhiteSpace(bulletText) || variants == null)
                return false;

            bulletText = bulletText.ToLower();

            foreach (var variant in variants)
            {
                if (!string.IsNullOrWhiteSpace(variant) &&
                    bulletText.Contains(variant.ToLower()))
                {
                    return true;
                }
            }

            return false;
        }

        private void HighlightChanges(List<Highlight> manualRequirements, List<Highlight> bulletPoints, Color color)
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(txtCleaningRecord.Text)))
            {
                HighlightList(excelPackage, manualRequirements, color, true);
                HighlightList(excelPackage, bulletPoints, color, false);

                excelPackage.Save();
            }
        }

        private void HighlightList(ExcelPackage excelPackage, List<Highlight> highLight, Color color, bool highlightNextColumn)
        {
            foreach (var hl in highLight)
            {
                ExcelWorksheet worksheet =
                    excelPackage.Workbook.Worksheets[hl.SPOPName];

                foreach (var column in hl.Columns2Highlight)
                {
                    var parts = column.Split('_');

                    var rowIndex = int.Parse(parts[1]);
                    var columnIndex = int.Parse(parts[2]);

                    var endColumn = highlightNextColumn
                        ? columnIndex + 1
                        : columnIndex;

                    var cells = worksheet.Cells[
                        rowIndex,
                        columnIndex,
                        rowIndex,
                        endColumn];

                    cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cells.Style.Fill.BackgroundColor.SetColor(color);
                }
            }
        }

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


        private List<string> GetVariants(string value)
        {
            var variants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(value))
                return variants.ToList();

            value = value.Trim();

            // Original
            variants.Add(value);

            // Normalize spaces around hyphen
            string normalized = Regex.Replace(value, @"\s*-\s*", " - ");
            variants.Add(normalized);

            // Unit aliases
            var unitAliases = new Dictionary<string, string[]>(
                StringComparer.OrdinalIgnoreCase)
    {
        { "°", new[] { "°", "degree", "degrees" } },
        { "degree", new[] { "°", "degree", "degrees" } },
        { "degrees", new[] { "°", "degree", "degrees" } },

        { "℃", new[] { "℃", "°C" } },
        { "°C", new[] { "℃", "°C" } },

        { "mm", new[] { "mm", "millimeter", "millimeters",
                        "millimetre", "millimetres" } },

        { "cm", new[] { "cm", "centimeter", "centimeters",
                        "centimetre", "centimetres" } },

        { "m", new[] { "m", "meter", "meters",
                       "metre", "metres" } },

        { "inch", new[] { "inch", "inches", "in" } },
        { "inches", new[] { "inch", "inches", "in" } },
        { "in", new[] { "inch", "inches", "in" } },

        { "ft", new[] { "ft", "foot", "feet" } },
        { "foot", new[] { "ft", "foot", "feet" } },
        { "feet", new[] { "ft", "foot", "feet" } },

        { "min", new[] { "min", "min.", "minute", "minutes" } },
        { "min.", new[] { "min", "min.", "minute", "minutes" } },
        { "minute", new[] { "min", "min.", "minute", "minutes" } },
        { "minutes", new[] { "min", "min.", "minute", "minutes" } },

        { "kg", new[] { "kg", "kilogram", "kilograms" } },
        { "g", new[] { "g", "gram", "grams" } },

        { "mpa", new[] { "mpa", "MPa" } }
    };

            // Find the unit at the end
            foreach (var unit in unitAliases)
            {
                var match = Regex.Match(
                    normalized,
                    @"^(?<number>.+?)\s*" +
                    Regex.Escape(unit.Key) +
                    @"\.?$",
                    RegexOptions.IgnoreCase);

                if (!match.Success)
                    continue;

                string numberPart = match.Groups["number"].Value.Trim();

                foreach (string alias in unit.Value)
                {
                    variants.Add(numberPart + " " + alias);
                }

                break;
            }

            // Generate different spacing around "-"
            foreach (string variant in variants.ToList())
            {
                if (!variant.Contains("-"))
                    continue;

                variants.Add(Regex.Replace(variant, @"\s*-\s*", "-"));
                variants.Add(Regex.Replace(variant, @"\s*-\s*", "- "));
                variants.Add(Regex.Replace(variant, @"\s*-\s*", " -"));
                variants.Add(Regex.Replace(variant, @"\s*-\s*", " - "));
            }

            return variants.ToList();
        }
    }
}
