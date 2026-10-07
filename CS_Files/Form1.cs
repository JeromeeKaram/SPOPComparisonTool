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

            string oldRevisionLink = txtOldRevisionLink.Text;
            string newRevisionLink = txtNewRevisionLink.Text;
            string cleaningRecord = txtCleaningRecord.Text;

            try
            {
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.MarqueeAnimationSpeed = 30;
                progressBar1.Visible = true;

                await Task.Run(() =>
                {
                    ReadExcel(oldRevisionLink, newRevisionLink, cleaningRecord);
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

        private void ReadExcel(string txtOldRevisionLink, string txtNewRevisionLink, string txtCleaningRecord)
        {
            var workSheetNames = new List<string>();
            var changedManualRequirements = new List<Highlight>();
            var changedBulletPoints = new List<Highlight>();
            var excelData = new List<ExcelData>();

            //var noData = new List<(string SPOP, string Bullet)>();
            //var data = new List<(string SPOP, string Bullet)>();

            var reader = GetReader();

            using (var package = new ExcelPackage(new FileInfo(txtCleaningRecord)))
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

                    //if (testWorkSheet.Name != "SPOP19 表") continue; //Comment this line to process all worksheets

                    //if (testWorkSheet.Name != "SPOP209 表") continue; //Comment this line to process all worksheets

                    //if (testWorkSheet.Name != "SPOP203 表") continue; //Comment this line to process all worksheets

                    if (testWorkSheet.Name != "SPOP2 表") continue; //Comment this line to process all worksheets


                    //Debug.WriteLine($"Processing sheet: {testWorkSheet.Name}");

                    //if (testWorkSheet.Name.StartsWith("SPOP209"))
                    //{
                    //    Debug.WriteLine($"Found SPOP209 sheet: {testWorkSheet.Name}");
                    //}

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
                                rowStartsAt = row + 1;
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

                            var manualRequirements = new List<(string Requirement, int Row)>();
                            var procedure = "";

                            for (var i = startRow; i <= endRow; i++)
                            {
                                var requirement = testWorkSheet.Cells[i, 7].Text.Trim();

                                if (string.IsNullOrWhiteSpace(requirement) || IgnoreManualRequirement(requirement))
                                {
                                    continue;
                                }

                                manualRequirements.Add((requirement, i));
                                procedure += testWorkSheet.Cells[i, 3].Text;
                            }

                            excelRecord.ManualRequirement = manualRequirements;
                            excelRecord.SPOPName = testWorkSheet.Name;
                            excelRecord.Row = startRow;

                            var match = Regex.Match(procedure, @"^SPOP\s*(\d+)");

                            if (match.Success)
                            {
                                excelRecord.IsNavigateToAnotherSPOP = true;
                                excelRecord.NavigateToSPOPName = $"SPOP {match.Groups[1].Value}";
                            }

                            // Text of the merged B cell
                            excelRecord.BulletPoint = Regex.Replace(testWorkSheet.Cells[startRow, 2].Text, @"\r?\n\s*", "");

                            excelData.Add(excelRecord);
                        }
                    }
                }
            }

            var manualReqs = excelData.Where(ed => ed.ManualRequirement != null && ed.ManualRequirement.Any()).ToList();

            foreach (var manualRequirement in manualReqs)
            {
                if (manualRequirement.BulletPoint != "1.A.") continue;

                string spaceBetweenSPOPandNumber = "";

                if (manualRequirement.IsNavigateToAnotherSPOP)
                {
                    spaceBetweenSPOPandNumber = manualRequirement.NavigateToSPOPName;
                }
                else
                {
                    string withoutJapanese = manualRequirement.SPOPName.Split(' ')[0];
                    spaceBetweenSPOPandNumber = Regex.Replace(withoutJapanese, @"^(SPOP)(\d+)$", "$1 $2");
                }

                var comparisonResult = GetComparisonResult(reader, spaceBetweenSPOPandNumber, manualRequirement.BulletPoint);

                if (comparisonResult == null)
                {
                    //noData.Add((record.SPOPName, record.BulletPoint));
                    continue;
                }
                //data.Add((record.SPOPName, record.BulletPoint));

                var bulletInfoNew = comparisonResult.m_new.m_sCompleteBullet;
                var bulletInfoOld = comparisonResult.m_old.m_sCompleteBullet;

                if (!string.Equals(bulletInfoOld, bulletInfoNew))
                {
                    var highLight = new Highlight();
                    highLight.Row = manualRequirement.Row;
                    highLight.SPOPName = manualRequirement.SPOPName;
                    highLight.Columns2Highlight = new List<string> { $"{manualRequirement.BulletPoint}_{manualRequirement.Row}_2" };
                    changedBulletPoints.Add(highLight);
                }

                if (comparisonResult != null && comparisonResult.m_eCompareType != CompareType.SAME)
                {
                    //var newRevisionBulletData = comparisonResult.m_new;
                    var newVersionbulletText = comparisonResult.m_new.m_sFullText;

                    var highLightManualRequirement = new List<string>();

                    //var manualRequirementRow = manualRequirement.Row;

                    foreach (var item in manualRequirement.ManualRequirement)
                    {
                        var splitItem = item.Requirement.Contains(":") ? item.Requirement.Split(':')[1].Trim() : item.Requirement.Trim();
                        var variants = GetVariants(splitItem);

                        //#if DEBUG
                        //                        foreach(var v in variants)
                        //                        {
                        //                            System.Diagnostics.Debug.WriteLine($"'{v}'");
                        //                        }
                        //#endif

                        bool matched =
ManualRequirementExistsInNewRevision(newVersionbulletText, variants) ||
RangeEndpointsExistInSameSentence(newVersionbulletText, splitItem);

                        if (!matched && splitItem.ToLowerInvariant().Contains("mpa"))
                        {
                            matched = IsEquivalentPressureRange(newVersionbulletText, splitItem);
                        }

                        if (!matched && ContainsTimeUnit(splitItem))
                        {
                            matched = IsEquivalentTimeRange(newVersionbulletText, splitItem);
                        }

                        if (!matched && (splitItem.Contains("℃") || splitItem.Contains("°C") || splitItem.ToLowerInvariant().Contains("degrees c") || splitItem.Contains("°c")))
                        {
                            matched = IsEquivalentTemperatureRange(newVersionbulletText, splitItem);
                        }

                        if (!matched)
                        {
                            highLightManualRequirement.Add($"{splitItem}_{item.Row}_7");
                        }
                    }

                    if (highLightManualRequirement.Any())
                    {
                        var highLight = new Highlight();
                        highLight.Row = manualRequirement.Row;
                        highLight.SPOPName = manualRequirement.SPOPName;
                        highLight.Columns2Highlight = highLightManualRequirement;
                        changedManualRequirements.Add(highLight);
                    }
                }
            }

            var outputFilePath = HighlightChanges(changedManualRequirements, changedBulletPoints, Color.Yellow, txtCleaningRecord);

            //CreateSPOPBulletExcel(data, noData, txtCleaningRecord);
            // Open the Excel file
            Process.Start(new ProcessStartInfo(outputFilePath)
            {
                UseShellExecute = true
            });
        }

        private bool IsEquivalentTemperatureRange(string websiteText, string excelValue)
        {
            var excelMatch = Regex.Match(
                excelValue,
                @"(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*(?:℃|°C|degrees?\s*c)",
                RegexOptions.IgnoreCase);

            if (!excelMatch.Success)
                return false;

            double excelMin = double.Parse(excelMatch.Groups[1].Value);
            double excelMax = double.Parse(excelMatch.Groups[2].Value);

            var matches = Regex.Matches(
                websiteText,
                @"(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*(?:degrees?\s*)?(c|f|℃|°c|°f)",
                RegexOptions.IgnoreCase);

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                double min = double.Parse(match.Groups[1].Value);
                double max = double.Parse(match.Groups[2].Value);
                string unit = match.Groups[3].Value.ToLowerInvariant();

                if (unit == "f" || unit == "°f")
                {
                    min = (min - 32) * 5.0 / 9.0;
                    max = (max - 32) * 5.0 / 9.0;
                }

                min = Math.Round(min);
                max = Math.Round(max);

                if (Math.Abs(min - excelMin) <= 1 &&
                    Math.Abs(max - excelMax) <= 1)
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsEquivalentTimeRange(string websiteText, string excelValue)
        {
            websiteText = ReplaceNumberWords(websiteText);

            var excelMatch = Regex.Match(
                excelValue,
                @"(?<start>\d+(?:\.\d+)?)\s*(?<startUnit>min\.?|mins?|minutes?|hr|hrs|hours?)?\s*-\s*(?<end>\d+(?:\.\d+)?)\s*(?<endUnit>min\.?|mins?|minutes?|hr|hrs|hours?)",
                RegexOptions.IgnoreCase);

            if (!excelMatch.Success)
                return false;

            string excelStartUnit = excelMatch.Groups["startUnit"].Value;
            string excelEndUnit = excelMatch.Groups["endUnit"].Value;

            // Example: "5 - 30 min."
            if (string.IsNullOrWhiteSpace(excelStartUnit))
                excelStartUnit = excelEndUnit;

            double excelStart = ToMinutes(
                double.Parse(excelMatch.Groups["start"].Value),
                excelStartUnit);

            double excelEnd = ToMinutes(
                double.Parse(excelMatch.Groups["end"].Value),
                excelEndUnit);

            var websiteMatches = Regex.Matches(
                websiteText,
                @"(?<start>\d+(?:\.\d+)?)\s*(?<startUnit>min\.?|mins?|minutes?|hr|hrs|hours?)?\s*(?:-|to)\s*(?<end>\d+(?:\.\d+)?)\s*(?<endUnit>min\.?|mins?|minutes?|hr|hrs|hours?)",
                RegexOptions.IgnoreCase);

            foreach (System.Text.RegularExpressions.Match match in websiteMatches)
            {
                string websiteStartUnit = match.Groups["startUnit"].Value;
                string websiteEndUnit = match.Groups["endUnit"].Value;

                // Example: "5 - 30 minutes"
                if (string.IsNullOrWhiteSpace(websiteStartUnit))
                    websiteStartUnit = websiteEndUnit;

                double websiteStart = ToMinutes(
                    double.Parse(match.Groups["start"].Value),
                    websiteStartUnit);

                double websiteEnd = ToMinutes(
                    double.Parse(match.Groups["end"].Value),
                    websiteEndUnit);

                if (Math.Abs(excelStart - websiteStart) <= 1 &&
                    Math.Abs(excelEnd - websiteEnd) <= 1)
                {
                    return true;
                }
            }

            return false;
        }

        private string ReplaceNumberWords(string text)
        {
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "one", "1" },
        { "two", "2" },
        { "three", "3" },
        { "four", "4" },
        { "five", "5" },
        { "six", "6" },
        { "seven", "7" },
        { "eight", "8" },
        { "nine", "9" },
        { "ten", "10" },
        { "eleven", "11" },
        { "twelve", "12" },
        { "thirteen", "13" },
        { "fourteen", "14" },
        { "fifteen", "15" },
        { "sixteen", "16" },
        { "seventeen", "17" },
        { "eighteen", "18" },
        { "nineteen", "19" },
        { "twenty", "20" }
    };

            foreach (var item in replacements)
            {
                text = Regex.Replace(
                    text,
                    $@"\b{item.Key}\b",
                    item.Value,
                    RegexOptions.IgnoreCase);
            }

            return text;
        }

        private double ToMinutes(double value, string unit)
        {
            switch (unit.ToLowerInvariant())
            {
                case "hr":
                case "hrs":
                case "hour":
                case "hours":
                    return value * 60;

                default:
                    return value;
            }
        }

        private bool ContainsTimeUnit(string text)
        {
            string lower = text.ToLowerInvariant();

            return lower.Contains("min") ||
                   lower.Contains("minute") ||
                   lower.Contains("hr") ||
                   lower.Contains("hour");
        }

        private bool IsEquivalentPressureRange(string newVersionText, string excelValue)
        {
            // Excel: 0.28 - 0.34 MPa

            // Extract Excel range
            var excelMatch = Regex.Match(
                excelValue,
                @"(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*MPa",
                RegexOptions.IgnoreCase);

            if (!excelMatch.Success)
                return false;

            double excelMin = double.Parse(excelMatch.Groups[1].Value);
            double excelMax = double.Parse(excelMatch.Groups[2].Value);

            // Find all pressure ranges in website text
            var matches = Regex.Matches(
                newVersionText,
                @"(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*(psi|kpa|mpa)",
                RegexOptions.IgnoreCase);

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                double min = double.Parse(match.Groups[1].Value);
                double max = double.Parse(match.Groups[2].Value);
                string unit = match.Groups[3].Value.ToLower();

                switch (unit)
                {
                    case "psi":
                        min *= 0.00689476;
                        max *= 0.00689476;
                        break;

                    case "kpa":
                        min /= 1000.0;
                        max /= 1000.0;
                        break;

                    case "mpa":
                        break;
                }

                min = Math.Round(min, 2);
                max = Math.Round(max, 2);

                if (min == excelMin && max == excelMax)
                    return true;
            }

            return false;
        }

        private bool IgnoreManualRequirement(string manualRequirementItem)
        {
            if (Regex.IsMatch(manualRequirementItem, @"^SPS\s+\d+(-\d+)*$"))
            {
                return true;
            }
            else if (manualRequirementItem.ToLower().StartsWith("super-bee"))
            {
                return true;
            }
            else if (manualRequirementItem.ToLower().StartsWith("tank no."))
            {
                return true;
            }
            else if (manualRequirementItem.ToLower().StartsWith("tc-"))
            {
                return true;
            }
            else if (manualRequirementItem.ToLower().StartsWith("cee bee"))
            {
                return true;
            }
            else if (manualRequirementItem.ToLower().StartsWith("polyplus prime"))
            {
                return true;
            }
            else if (manualRequirementItem.ToLower().StartsWith("pmc "))
            {
                return true;
            }
            else if (manualRequirementItem.ToLower().StartsWith("spmc "))
            {
                return true;
            }

            return false;
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

        private string HighlightChanges(List<Highlight> manualRequirements, List<Highlight> bulletPoints, Color color, string txtCleaningRecord)
        {
            var sourceFile = new FileInfo(txtCleaningRecord);

            var outputFilePath = Path.Combine(
        sourceFile.DirectoryName,
        Path.GetFileNameWithoutExtension(sourceFile.Name) + "_Highlighted.xlsx");

            using (var excelPackage = new ExcelPackage(sourceFile))
            {
                HighlightList(excelPackage, manualRequirements, color, true);
                HighlightList(excelPackage, bulletPoints, color, false);

                //        var outputFile = new FileInfo(
                //Path.Combine(
                //    sourceFile.DirectoryName,
                //    Path.GetFileNameWithoutExtension(sourceFile.Name) + "_Highlighted.xlsx"));

                excelPackage.SaveAs(new FileInfo(outputFilePath));
            }

            return outputFilePath;
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

            // Add parenthesized version of every variant
            foreach (string variant in variants.ToList())
            {
                variants.Add($"({variant})");
            }

            return variants.ToList();
        }

        public void CreateSPOPBulletExcel(
    List<(string SPOP, string Bullet)> dataList,
    List<(string SPOP, string Bullet)> noDataList,
    string savePath)
        {
            var sourceFile = new FileInfo(txtCleaningRecord.Text);

            var outputFilePath = Path.Combine(
                sourceFile.DirectoryName,
                Path.GetFileNameWithoutExtension(sourceFile.Name) + "_SPOP_Bullet.xlsx");

            using (var package = new ExcelPackage())
            {
                // Data sheet
                var dataWorksheet = package.Workbook.Worksheets.Add("Data");

                dataWorksheet.Cells[1, 1].Value = "SPOP";
                dataWorksheet.Cells[1, 2].Value = "Bullet";

                for (int i = 0; i < dataList.Count; i++)
                {
                    dataWorksheet.Cells[i + 2, 1].Value = dataList[i].SPOP;
                    dataWorksheet.Cells[i + 2, 2].Value = dataList[i].Bullet;
                }

                dataWorksheet.Column(1).AutoFit();
                dataWorksheet.Column(2).AutoFit();


                // NoData sheet
                var noDataWorksheet = package.Workbook.Worksheets.Add("NoData");

                noDataWorksheet.Cells[1, 1].Value = "SPOP";
                noDataWorksheet.Cells[1, 2].Value = "Bullet";

                for (int i = 0; i < noDataList.Count; i++)
                {
                    noDataWorksheet.Cells[i + 2, 1].Value = noDataList[i].SPOP;
                    noDataWorksheet.Cells[i + 2, 2].Value = noDataList[i].Bullet;
                }

                noDataWorksheet.Column(1).AutoFit();
                noDataWorksheet.Column(2).AutoFit();


                // Save
                package.SaveAs(new FileInfo(outputFilePath));
            }
        }

        private bool RangeEndpointsExistInSameSentence(string newVersionBulletText, string manualRequirement)
        {
            var rangeMatch = Regex.Match(
                manualRequirement,
                @"^(?<min>\d+(?:\.\d+)?)\s*-\s*(?<max>\d+(?:\.\d+)?)\s*(?<unit>mm|cm|m|in|inch|ft|degree|degrees|°)$",
                RegexOptions.IgnoreCase);

            if (!rangeMatch.Success)
                return false;

            double min = double.Parse(rangeMatch.Groups["min"].Value);
            double max = double.Parse(rangeMatch.Groups["max"].Value);
            string unit = rangeMatch.Groups["unit"].Value;

            // Split into sentences
            var sentences = Regex.Split(newVersionBulletText, @"(?<=[.!?])\s+(?=[A-Z])");

            foreach (var sentence in sentences)
            {
                bool hasMin = Regex.IsMatch(
                    sentence,
                    $@"\b{min}\s*{Regex.Escape(unit)}\b",
                    RegexOptions.IgnoreCase);

                bool hasMax = Regex.IsMatch(
                    sentence,
                    $@"\b{max}\s*{Regex.Escape(unit)}\b",
                    RegexOptions.IgnoreCase);

                if (hasMin && hasMax)
                    return true;
            }

            return false;
        }
    }
}
