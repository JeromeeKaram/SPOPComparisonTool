using Cleaning_Records_Sheet;
using Cleaning_Records_Sheet.About;
using EngineManualReader;
using HtmlAgilityPack;
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
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Security.Policy;
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
            var workSheetNamesForReader = new List<string>();
            var workSheetNames = new List<string>();
            //var changedManualRequirements = new List<Highlight>();
            //var changedBulletPoints = new List<Highlight>();
            var excelData = new List<Bullet>();

            //var noData = new List<(string SPOP, string Bullet)>();
            //var data = new List<(string SPOP, string Bullet)>();
            var spmReader = new SPMReader();

            using (var package = new ExcelPackage(new FileInfo(txtCleaningRecord)))
            {
                var seen = new HashSet<string>();

                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Name.Contains("SPOP213")) continue; //Ignore this worksheet （Not used in Tsurugashima factory)

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
            }

            //Removing the japanese character and inserting space after SPOP for all items
            for (int i = 0; i < workSheetNames.Count; i++)
            {
                workSheetNamesForReader.Add(Regex.Replace(workSheetNames[i], @"^(SPOP)\s*(\d+).*", "$1 $2"));
            }

            var reader = GetReader(txtOldRevisionLink, txtNewRevisionLink, workSheetNamesForReader);

            using (var package = new ExcelPackage(new FileInfo(txtCleaningRecord)))
            {
                foreach (var name in workSheetNames)
                {
                    var testWorkSheet = package.Workbook.Worksheets[name];

                    if (testWorkSheet == null)
                        continue;

                    //if (testWorkSheet.Name != "SPOP15 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP18 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP19 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP203 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP209 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP213 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP218 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP225 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP2 表") continue; //Comment this line to process all worksheets
                    //if (testWorkSheet.Name != "SPOP211 表") continue; //Comment this line to process all worksheets

                    var matchSpopName = Regex.Match(testWorkSheet.Name, @"SPOP\s*(\d+)", RegexOptions.IgnoreCase);
                    var spopName = matchSpopName.Success ? $"SPOP {matchSpopName.Groups[1].Value}" : testWorkSheet.Name;

                    //var isDMCSame = reader.IsDMCSame(spopName);

                    //if (isDMCSame)
                    //{
                    //    Debug.WriteLine($"No changes in {testWorkSheet.Name}");
                    //    continue;
                    //}

                    int rowStartsAt = GetDataStartRow(testWorkSheet);

                    if (rowStartsAt == 0)
                    {
                        Debug.WriteLine($"Header not found in {testWorkSheet.Name}");
                        continue;
                    }

                    if (rowStartsAt > 0)
                    {
                        var mergedRanges = GetMergedRowsForColumnB(testWorkSheet, rowStartsAt);

                        foreach (var mergedRange in mergedRanges)
                        {
                            var excelRecord = new Bullet();
                            var startRow = mergedRange.StartRow;
                            var endRow = mergedRange.EndRow;

                            var manualRequirements = new List<(string Requirement, int Row)>();
                            var procedureText = "";
                            var manualRequirement = "";

                            for (var i = startRow; i <= endRow; i++)
                            {
                                var requirement = testWorkSheet.Cells[i, 7].Text.Trim();

                                if (!string.IsNullOrEmpty(requirement))
                                {
                                    manualRequirements.Add((requirement, i));
                                }
                                procedureText += $" {testWorkSheet.Cells[i, 3].Text}";
                                manualRequirement += $" {requirement}";
                            }

                            excelRecord.ManualRequirement = manualRequirements;
                            excelRecord.SPOPName = testWorkSheet.Name;
                            excelRecord.Row = startRow;
                            excelRecord.ProcedureText = procedureText;

                            var match = Regex.Match(procedureText, @"SPOP\s*(\d+)");

                            if (match.Success)
                            {
                                excelRecord.IsNavigateToAnotherSPOP = true;
                                excelRecord.NavigateToSPOPName = $"SPOP {match.Groups[1].Value}";
                            }

                            var matchSPS = Regex.Match(manualRequirement, @"(SPS|PMC|SPMC)\s*(\d+(?:-\d+)?)", RegexOptions.IgnoreCase);

                            if (matchSPS.Success)
                            {
                                excelRecord.HasSPSinManualRequirement = true;
                                excelRecord.SPSinManualRequirement = $"{matchSPS.Groups[1].Value.ToUpper()} {matchSPS.Groups[2].Value}";
                            }

                            // Text of the merged B cell
                            excelRecord.BulletPoint = Regex.Replace(testWorkSheet.Cells[startRow, 2].Text, @"\r?\n\s*", "");

                            excelData.Add(excelRecord);
                        }
                    }
                }
            }

            var bullets = excelData.Where(ed => ed.ManualRequirement != null && ed.ManualRequirement.Any()).ToList();

            var bulletsWithDifferentSPOP = bullets.Where(b => b.IsNavigateToAnotherSPOP).ToList();
            var bulletsWithoutDifferentSPOP = bullets.Where(b => !b.IsNavigateToAnotherSPOP).ToList();

            var dataDifferentSPOP = ProcessBulletsDifferentSPOP(bulletsWithDifferentSPOP, reader, txtNewRevisionLink);
            var dataSameSPOP = ProcessBulletsSameSPOP(bulletsWithoutDifferentSPOP, reader, txtNewRevisionLink);

            var changedManualRequirements = new List<Highlight>();

            changedManualRequirements.AddRange(dataDifferentSPOP.cmrDifferentSPOP);
            changedManualRequirements.AddRange(dataSameSPOP.changedManualRequirements);

            var observationsList = new Dictionary<string, List<string>>();

            foreach (var source in new[]
            {
    dataDifferentSPOP.observationsList,
    dataSameSPOP.observationsList
})
            {
                foreach (var kvp in source)
                {
                    if (!observationsList.TryGetValue(kvp.Key, out var list))
                    {
                        list = new List<string>();
                        observationsList[kvp.Key] = list;
                    }

                    list.AddRange(kvp.Value);
                }
            }

            var outputFilePath = HighlightChanges(changedManualRequirements, dataSameSPOP.changedBulletPoints, Color.Yellow, txtCleaningRecord);

            //CreateSPOPBulletExcel(data, noData, txtCleaningRecord);
            //Open the Excel file

            Process.Start(new ProcessStartInfo(outputFilePath)
            {
                UseShellExecute = true
            });
        }

        private (List<Highlight> changedManualRequirements, List<Highlight> changedBulletPoints, Dictionary<string, List<string>> observationsList)
            ProcessBulletsSameSPOP(List<Bullet> bulletsWithoutDifferentSPOP, SPMReader reader, string txtNewRevisionLink)
        {
            var changedBulletPoints = new List<Highlight>();
            var changedManualRequirements = new List<Highlight>();
            var observationsList = new Dictionary<string, List<string>>();

            foreach (var bullet in bulletsWithoutDifferentSPOP)
            {
                //if (bullet.BulletPoint != "2.A.(2)") continue;

                var key = $"{bullet.SPOPName}_{bullet.BulletPoint}";

                if (!observationsList.ContainsKey(key))
                {
                    observationsList.Add(key, new List<string>());
                }

                var spaceBetweenSPOPandNumber = "";
                var bulletPoint = "";

                if (bullet.IsNavigateToAnotherSPOP)
                {
                    spaceBetweenSPOPandNumber = bullet.NavigateToSPOPName;
                }
                else
                {
                    string withoutJapanese = bullet.SPOPName.Split(' ')[0];
                    spaceBetweenSPOPandNumber = Regex.Replace(withoutJapanese, @"^(SPOP)(\d+)$", "$1 $2");
                    bulletPoint = bullet.BulletPoint;
                }

                var comparisonResult = GetComparisonResult(reader, spaceBetweenSPOPandNumber, bulletPoint);

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
                    highLight.Row = bullet.Row;
                    highLight.SPOPName = bullet.SPOPName;
                    highLight.Columns2Highlight = new List<string> { $"{bullet.BulletPoint}_{bullet.Row}_2" };
                    changedBulletPoints.Add(highLight);
                }

                //if (comparisonResult.m_eCompareType != CompareType.SAME)
                {
                    var newVersionbulletText1 = comparisonResult.m_new.m_sText;

                    var highLightManualRequirement = new List<string>();

                    //var procedureSteps = comparisonResult.m_new.m_oParent.m_lstChilds;

                    var validManualRequirements = new List<(string Requirement, int Row)>();

                    foreach (var item in bullet.ManualRequirement)
                    {
                        if (string.IsNullOrWhiteSpace(item.Requirement) || IgnoreManualRequirement(item.Requirement))
                        {
                            continue;
                        }

                        validManualRequirements.Add(item);
                    }

                    var linesFromProcedure = new List<string>();

                    if (bullet.IsNavigateToAnotherSPOP)
                    {
                        linesFromProcedure = GetProcedureLines(reader, bullet, validManualRequirements, txtNewRevisionLink);
                    }
                    else
                    {
                        //TODO:SPS exists for 1 item and it does not exist for another item
                        var procedures = comparisonResult.m_new.m_oParent.m_lstChilds;
                        linesFromProcedure = procedures.Select(item => item.m_sFullText).ToList();
                    }

                    foreach (var item in validManualRequirements)
                    {
                        var parts = item.Requirement.Split(new[] { ':' }, 2);

                        var splitItem = parts.Length == 2
                            ? parts[1].Trim()
                            : item.Requirement.Trim();

                        var variants = GetVariants(splitItem);

                        bool matched = false;

                        foreach (var procedure in linesFromProcedure)
                        {
                            matched = ManualRequirementExistsInNewRevision(procedure, variants) ||
RangeEndpointsExistInSameSentence(procedure, splitItem);

                            if (!matched && splitItem.ToLowerInvariant().Contains("mpa"))
                            {
                                matched = IsEquivalentPressureRange(procedure, splitItem);
                            }

                            if (!matched && ContainsTimeUnit(splitItem))
                            {
                                matched = IsEquivalentTimeRange(procedure, splitItem);
                            }

                            if (!matched && (splitItem.Contains("℃") || splitItem.Contains("°C") || splitItem.ToLowerInvariant().Contains("degrees c") || splitItem.Contains("°c")))
                            {
                                matched = IsEquivalentTemperatureRange(procedure, splitItem);
                            }

                            if (matched)
                            {
                                if (observationsList.TryGetValue($"{bullet.SPOPName}_{bullet.BulletPoint}", out List<string> values))
                                {
                                    values.Add(item.Requirement);
                                }

                                break;
                            }
                        }

                        if (!matched)
                        {
                            highLightManualRequirement.Add($"{splitItem}_{item.Row}_7");
                        }
                    }

                    if (highLightManualRequirement.Any())
                    {
                        var highLight = new Highlight();
                        highLight.Row = bullet.Row;
                        highLight.SPOPName = bullet.SPOPName;
                        highLight.Columns2Highlight = highLightManualRequirement;
                        changedManualRequirements.Add(highLight);
                    }
                }
            }

            return (changedManualRequirements, changedBulletPoints, observationsList);
        }

        private (List<Highlight> cmrDifferentSPOP, Dictionary<string, List<string>> observationsList)
            ProcessBulletsDifferentSPOP(List<Bullet> bulletsWithDifferentSPOP, SPMReader reader, string newRevisionLink)
        {
            var cmrDifferentSPOP = new List<Highlight>();
            var observationsList = new Dictionary<string, List<string>>();

            foreach (var bullet in bulletsWithDifferentSPOP)
            {
                //if (bullet.BulletPoint != "2.A.(2)") continue;

                var key = $"{bullet.SPOPName}_{bullet.BulletPoint}";

                if (!observationsList.ContainsKey(key))
                {
                    observationsList.Add(key, new List<string>());
                }

                //if (comparisonResult.m_eCompareType != CompareType.SAME)
                {
                    //var newVersionbulletText1 = comparisonResult.m_new.m_sFullText;

                    var highLightManualRequirement = new List<string>();

                    //var procedureSteps = comparisonResult.m_new.m_oParent.m_lstChilds;

                    var validManualRequirements = new List<(string Requirement, int Row)>();

                    foreach (var item in bullet.ManualRequirement)
                    {
                        if (string.IsNullOrWhiteSpace(item.Requirement) || IgnoreManualRequirement(item.Requirement))
                        {
                            continue;
                        }

                        validManualRequirements.Add(item);
                    }

                    var str = reader.GetDMC(bullet.NavigateToSPOPName);

                    //var procedureSteps = reader.GetProcedure(bullet.NavigateToSPOPName, "New");
                    var linesFromProcedure = GetProcedureLines(reader, bullet, validManualRequirements, newRevisionLink);

                    foreach (var item in validManualRequirements)
                    {
                        var parts = item.Requirement.Split(new[] { ':' }, 2);

                        var splitItem = parts.Length == 2
                            ? parts[1].Trim()
                            : item.Requirement.Trim();

                        var variants = GetVariants(splitItem);

                        bool matched = false;

                        foreach (var procedure in linesFromProcedure)
                        {
                            matched = ManualRequirementExistsInNewRevision(procedure, variants) ||
RangeEndpointsExistInSameSentence(procedure, splitItem);

                            if (!matched && splitItem.ToLowerInvariant().Contains("mpa"))
                            {
                                matched = IsEquivalentPressureRange(procedure, splitItem);
                            }

                            if (!matched && ContainsTimeUnit(splitItem))
                            {
                                matched = IsEquivalentTimeRange(procedure, splitItem);
                            }

                            if (!matched && (splitItem.Contains("℃") || splitItem.Contains("°C") || splitItem.ToLowerInvariant().Contains("degrees c") || splitItem.Contains("°c")))
                            {
                                matched = IsEquivalentTemperatureRange(procedure, splitItem);
                            }

                            if (matched)
                            {
                                if (observationsList.TryGetValue($"{bullet.SPOPName}_{bullet.BulletPoint}", out List<string> values))
                                {
                                    values.Add(item.Requirement);
                                }

                                break;
                            }
                        }

                        if (!matched)
                        {
                            highLightManualRequirement.Add($"{splitItem}_{item.Row}_7");
                        }
                    }

                    if (highLightManualRequirement.Any())
                    {
                        var highLight = new Highlight();
                        highLight.Row = bullet.Row;
                        highLight.SPOPName = bullet.SPOPName;
                        highLight.Columns2Highlight = highLightManualRequirement;
                        cmrDifferentSPOP.Add(highLight);
                    }
                }
            }

            return (cmrDifferentSPOP, observationsList);
        }

        private List<string> GetProcedureLines(SPMReader reader, Bullet bullet,
            List<(string Requirement, int Row)> validManualRequirements, string newRevisionLink)
        {
            //open link and read the lines

            var spopLink = reader.GetDMC(bullet.NavigateToSPOPName);
            if (string.IsNullOrEmpty(spopLink)) return new List<string>();

            var web = new HtmlWeb();
            var doc = web.Load($"{newRevisionLink}/{spopLink}.html");

            var stepLevel0 = doc.DocumentNode.SelectSingleNode("//ol[@class='steplevel0']");
            var secondLi = stepLevel0?.SelectSingleNode("./li[2]");

            string websiteText = HtmlEntity.DeEntitize(doc.DocumentNode.InnerText);

            var allMatchingSentences = new List<string>();

            (var acqCleanerSolnTemperatureList, var soakTimeList, var allInOne) = GetChildBullets(doc, bullet.ProcedureText);

            foreach (var req in validManualRequirements)
            {
                var temperatureSentences = new List<string>();
                var durationSentences = new List<string>();

                var isTemperature = Regex.IsMatch(req.Requirement, @"℃|°\s*C|degrees\s*C", RegexOptions.IgnoreCase);

                if (isTemperature)
                {
                    if (bullet.ProcedureText.ToLower().Contains("immersion"))
                    {
                        if (bullet.HasSPSinManualRequirement)
                        {
                            var spsCompliant = acqCleanerSolnTemperatureList.Where
                                 (p => p.ToLower().Contains(bullet.SPSinManualRequirement.ToLower())).Select(p => p).ToList();

                            temperatureSentences = spsCompliant.Where(p => Regex.IsMatch(p, @"°\s*[FC]|℃|℉|degrees\s*[FC]", RegexOptions.IgnoreCase)).
                                  Select(p => p).ToList();
                        }
                        else
                        {
                            temperatureSentences = acqCleanerSolnTemperatureList.Where(p => Regex.IsMatch(p, @"°\s*[FC]|℃|℉|degrees\s*[FC]", RegexOptions.IgnoreCase)).
                                  Select(p => p).ToList();
                        }
                    }
                    else
                    {
                        temperatureSentences = allInOne.Where(p => Regex.IsMatch(p, @"°\s*[FC]|℃|℉|degrees\s*[FC]", RegexOptions.IgnoreCase)).
                              Select(p => p).ToList();
                    }

                    allMatchingSentences.AddRange(temperatureSentences);
                    continue;
                }

                var isDuration = ContainsTimeUnit(req.Requirement);

                if (isDuration)
                {
                    if (bullet.ProcedureText.ToLower().Contains("immersion"))
                    {
                        if (bullet.HasSPSinManualRequirement)
                        {
                            var spsCompliant = soakTimeList
    .Where(p => p.ToLower().Contains(bullet.SPSinManualRequirement.ToLower()))
    .ToList();
                            if (!spsCompliant.Any())
                            {
                                durationSentences = soakTimeList.Where(p => Regex.IsMatch(p,
    @"For all other solutions", RegexOptions.IgnoreCase)).ToList();
                            }
                        }
                        else
                        {
                            durationSentences = soakTimeList.Where(p => Regex.IsMatch(p, @"°\s*[FC]|℃|℉|degrees\s*[FC]", RegexOptions.IgnoreCase)).
                                                                  Select(p => p).ToList();
                        }
                    }
                    else
                    {
                        durationSentences = allInOne.Where(p => Regex.IsMatch(p, @"°\s*[FC]|℃|℉|degrees\s*[FC]", RegexOptions.IgnoreCase)).
                              Select(p => p).ToList();
                    }
                    allMatchingSentences.AddRange(durationSentences);
                    continue;
                }

                var isPressure = req.Requirement.ToLower().Contains("mpa") ? true : false;

                if (isPressure)
                {
                    var pressureSentences = allInOne
.Where(p => Regex.IsMatch(
    p,
    @"\bpsi\b|\bkpa\b",
    RegexOptions.IgnoreCase))
.Select(p => p)
.ToList();
                    allMatchingSentences.AddRange(pressureSentences);
                    continue;
                }

                var isAngle = Regex.IsMatch(
req.Requirement,
@"\d+\s*(?:-\s*\d+)?\s*°");

                if (isAngle)
                {
                    var angleSentences = allInOne
.Where(p => Regex.IsMatch(
    p,
    @"\d+\s*(?:-\s*\d+)?\s*°|\d+\s*(?:-\s*\d+)?\s*degrees?",
    RegexOptions.IgnoreCase))
.Select(p => p)
.ToList();
                    allMatchingSentences.AddRange(angleSentences);
                    continue;
                }

                var isLength = Regex.IsMatch( req.Requirement, @"\d+(?:\.\d+)?\s*(?:-\s*\d+(?:\.\d+)?)?\s*mm\b", RegexOptions.IgnoreCase);

                if (isLength)
                {
                    var lengthSentences = allInOne
                        .Where(p => Regex.IsMatch(
                            p,
                            @"\b(mm|inch|inches)\b",
                            RegexOptions.IgnoreCase))
                        .ToList();

                    allMatchingSentences.AddRange(lengthSentences);
                    continue;
                }
            }
            return allMatchingSentences;
        }

        private static int GetDataStartRow(ExcelWorksheet testWorkSheet)
        {
            for (int row = 1; row <= testWorkSheet.Dimension.Rows; row++)
            {
                string value = testWorkSheet.Cells[row, 2].Text;

                if (!string.IsNullOrEmpty(value))
                {
                    value = value.Replace("<br>", " ")
                                 .Replace("<br/>", " ")
                                 .Replace("<br />", " ");

                    if (value.Contains("Para.") && value.Contains("No."))
                    {
                        return row + 1;
                    }
                }
            }

            return 0;
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


        private SPMReader GetReader(string oldRevisionLink, string newRevisionLink, List<string> lstSPOPs)
        {
            var reader = new SPMReader();
            reader.ReadSPM(oldRevisionLink, newRevisionLink, lstSPOPs);
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

            // Add .0 variants for whole numbers
            foreach (var variant in variants.ToList())
            {
                variants.Add(Regex.Replace(
                    variant,
                    @"\b(\d+)\b",
                    "$1.0"));
            }

            foreach (var variant in variants.ToList())
            {
                variants.Add(Regex.Replace(
                    variant,
                    @"\b(\d+)\.0+\b",
                    "$1"));
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
                if (Regex.IsMatch(
        sentence,
        $@"\b{Regex.Escape(min.ToString())}(?:\.0+)?\s*-\s*{Regex.Escape(max.ToString())}(?:\.0+)?\s*{Regex.Escape(unit)}\b",
        RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static (List<string> acqCleanerSolnTemperatureList, List<string> soakTimeList, List<string> allInOne)
            GetChildBullets(HtmlAgilityPack.HtmlDocument doc, string bulletProcedure)
        {
            var acqCleanerSolnTemperatureList = new List<string>();
            var soakTimeList = new List<string>();
            var allInOne = new List<string>();

            if (bulletProcedure.ToLower().Contains("immersion"))
            {
                var matchingLi = doc.DocumentNode.SelectSingleNode(
        "//li[contains(., 'aqueous cleaner solutions')]");

                var table = matchingLi?.SelectSingleNode(
                    ".//table[contains(@class,'noneFrame') and contains(@class,'hoverRowBackEnabled')]");

                foreach (var row in table?.SelectNodes(".//tr") ?? Enumerable.Empty<HtmlNode>())
                {
                    acqCleanerSolnTemperatureList.Add(HtmlEntity.DeEntitize(row.InnerText).Trim());
                }

                var ol = doc.DocumentNode.SelectSingleNode(
        "//ol[@class='steplevel4' and div/p[contains(text(),'Let the part soak in the solution for the applicable time:')]]");

                foreach (var li in ol?.SelectNodes("./li") ?? Enumerable.Empty<HtmlNode>())
                {
                    var text = li.SelectSingleNode(".//p")?.InnerText.Trim();

                    if (!string.IsNullOrEmpty(text))
                    {
                        soakTimeList.Add(text);
                    }
                }
            }
            else
            {
                foreach (var li in doc.DocumentNode.SelectNodes("//ol[@class='steplevel0']//li")
                         ?? Enumerable.Empty<HtmlNode>())
                {
                    allInOne.Add(HtmlEntity.DeEntitize(li.InnerText).Trim());
                }
            }

            return (acqCleanerSolnTemperatureList, soakTimeList, allInOne);
        }
    }
}
