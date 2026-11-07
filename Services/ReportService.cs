using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using Topo.Model.AdditionalAwards;
using Topo.Model.Approvals;
using Topo.Model.Program;
using Topo.Model.Logbook;
using Topo.Model.Members;
using Topo.Model.Milestone;
using Topo.Model.OAS;
using System.Globalization;
using Topo.Model.SIA;
using Topo.Model.SiaProgress;
using Topo.Model.Wallchart;
using Topo.Model.Progress;
using Topo.Model.History;
using Topo.Model.UnitReport;
using Topo.Model.SectionReport;
using System.Globalization;
using Syncfusion.EJ2.Spreadsheet;
using System;
using Syncfusion.EJ2.Linq;

namespace Topo.Services
{
    public interface IReportService
    {
        public IWorkbook GenerateAdditionalAwardsWorkbook(List<AdditionalAwardSpecificationListModel> awardSpecificationsList, List<AdditionalAwardListModel> sortedAdditionalAwardsList, List<string>? distinctAwards, string groupName, string section, string unitName);
        public IWorkbook GeneratePatrolListWorkbook(List<MemberListModel> sortedPatrolList, string groupName, string section, string unitName, bool includeLeaders);
        public IWorkbook GenerateMemberListWorkbook(List<MemberListModel> sortedMemberList, string groupName, string section, string unitName);
        public IWorkbook GeneratePatrolSheetsWorkbook(List<MemberListModel> sortedPatrolList, string section);
        public IWorkbook GenerateSignInSheetWorkbook(List<MemberListModel> memberListModel, string groupName, string section, string unitName, string eventName, SignInSheetOptions options);
        public IWorkbook GenerateEventAttendanceWorkbook(EventListModel eventListModel, string groupName, string section, string unitName);
        public IWorkbook GenerateAttendanceReportWorkbook(AttendanceReportModel attendanceReportData, string groupName, string section, string unitName, DateTime fromDate, DateTime toDate, bool forPdfOutput);
        public IWorkbook GenerateOASWorksheetWorkbook(List<OASWorksheetAnswers> worksheetAnswers, string groupName, string section, string unitName, bool forPdfOutput, bool formatLikeTerrain, bool showByPatrol = false);
        public IWorkbook GenerateSIAWorkbook(List<SIAProjectListModel> siaProjects, string groupName, string section, string unitName, bool forPdfOutput);
        public IWorkbook GenerateMilestoneWorkbook(List<MilestoneSummaryListModel> milestoneSummaries, string groupName, string section, string unitName, bool forPdfOutput, bool breakByPatrol = false);
        public IWorkbook GenerateLogbookWorkbook(List<MemberLogbookReportViewModel> logbookEntries, string groupName, string section, string unitName, bool forPdfOutput);
        public IWorkbook GenerateWallchartWorkbook(List<WallchartItemModel> wallchartEntries, string groupName, string section, string unitName, bool forPdfOutput, bool breakByPatrol = false);
        public IWorkbook GenerateApprovalsWorkbook(List<ApprovalsListModel> selectedApprovals, string groupName, string section, string unitName, DateTime approvalSearchFromDate, DateTime approvalSearchToDate, bool groupByMember, bool forPdfOutput);
        public IWorkbook GenerateProgressWorkbook(ProgressDetailsPageViewModel progressEntries, string groupName, string section, string unitName);
        public IWorkbook GenerateScoutingHistoryWorkbook(ScoutingHistoryModel history, string groupName, string section, string unitName);
        public IWorkbook GenerateUnitReportWorkbook(UnitReportModel report, string groupName, string section, string unitName);
        public IWorkbook GenerateSectionReportWorkbook(SectionReportSet report, string groupName, string section, string unitName);
        public IWorkbook GenerateTermProgramWorkbook(List<EventListModel> wallchartEntries, string groupName, string section, string unitName, bool forPdfOutput);
        public IWorkbook GenerateProgramIdeasWorkbook(List<ProgramIdeaItemModel> programIdeas, string groupName, string section, string unitName, bool forPdfOutput);
        public IWorkbook GenerateSiaProgressWorkbook(SiaProgressReport report, string groupName, string section);
        public IWorkbook CreateWorkbookWithSheets(int sheetsToCreate);
    }
    public class ReportService : IReportService
    {
        public TimeSpan CurrentUtcOffset { get; set; } = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);

        private enum participateAssistLead
        {
            participate,
            assist,
            lead
        }

        private enum wallchartGroups
        {
            intro,
            participate,
            assist,
            lead,
            oasCore,
            oasLand,
            oasWater,
            oasProgression,
            siaOdd,
            siaEven,
            leadershipCourse,
            adventurousJourney,
            personalReflection,
            peakAward
        }

        private Color[] MilestoneColours1 = new Color[]
        {
            Color.FromArgb(236, 82, 82),
            Color.FromArgb(156, 242, 80)
        };

        private Color[] MilestoneColours2 = new Color[]
        {
            Color.FromArgb(236, 82, 82),
            Color.FromArgb(253, 238, 0),
            Color.FromArgb(156, 242, 80)
        };

        private Color[] MilestoneColours4 = new Color[]
        {
            Color.FromArgb(236, 82, 82),
            Color.FromArgb(244, 157, 43),
            Color.FromArgb(253, 238, 0),
            Color.FromArgb(206, 240, 38),
            Color.FromArgb(156, 242, 80)
        };

        private Color[] MilestoneColours5 = new Color[]
        {
            Color.FromArgb(236, 82, 82),
            Color.FromArgb(241, 131, 57),
            Color.FromArgb(248, 189, 26),
            Color.FromArgb(221, 239, 26),
            Color.FromArgb(184, 241, 57),
            Color.FromArgb(156, 242, 80)
        };

        private Color[] MilestoneColours6 = new Color[]
        {
            Color.FromArgb(236, 82, 82),
            Color.FromArgb(241, 129, 57),
            Color.FromArgb(248, 194, 23),
            Color.FromArgb(253, 238, 0),
            Color.FromArgb(218, 240, 29),
            Color.FromArgb(183, 241, 58),
            Color.FromArgb(156, 242, 80)
         };


        private bool RunningInUtc { get; set; }

        public ReportService()
        {
            RunningInUtc = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now) == TimeSpan.Zero;
        }

        // A summary of where the Scouts are up to (when asked for), then one sheet for each Special Interest Area project: where
        // it is up to, the plan and the review as question and answer, the files attached, and room for the leader's notes.
        // A project that is too long for a page runs onto a second. For double-sided printing each Scout (and the summary)
        // takes whole sheets of paper, with a blank page added wherever the next one would start on the back of a sheet.
        public IWorkbook GenerateSiaProgressWorkbook(SiaProgressReport report, string groupName, string section)
        {
            const string dateFormat = "d MMM yyyy";
            // what an A4 page holds with the margins used, in points
            const double printWidth = 559;
            const double printHeight = 806;
            // a page may be shrunk this far to keep a project on one page before it is given a second
            const double smallestScale = 0.75;
            var culture = CultureInfo.InvariantCulture;
            var workbook = CreateWorkbookWithSheets(Math.Max(1, (report.IncludeSummary ? 1 : 0) + report.Projects.Count));
            var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // each sheet in order, whose it is and how many pages it prints on
            var printed = new List<(string owner, IWorksheet sheet, int pages)>();

            // the logo hangs down from the top left corner, so leave room for it before the lines beside it begin
            int RoomForLogo(IWorksheet sheet, int rowNumber, double usedPoints)
            {
                var logoPoints = sheet.Pictures.Count > 0 ? sheet.Pictures[0].Height * 0.75 : 0.0;
                if (logoPoints + 6 <= usedPoints)
                    return rowNumber;
                rowNumber++;
                sheet.SetRowHeight(rowNumber, logoPoints + 6 - usedPoints);
                return rowNumber;
            }

            // Rows have the heights set for them, so how many pages a sheet needs is known: one if it fits when shrunk a
            // little, otherwise as many as its length at full size needs.
            int SetUpPage(IWorksheet sheet, int lastRow, int lastColumn)
            {
                double height = 0, width = 0;
                for (var row = 1; row <= lastRow; row++)
                    height += sheet.GetRowHeight(row);
                for (var column = 1; column <= lastColumn; column++)
                    width += sheet.GetColumnWidthInPixels(column) * 0.75;
                var pagesLong = height * Math.Min(1.0, printWidth / width) / printHeight;
                var pages = pagesLong <= 1 / smallestScale ? 1 : (int)Math.Ceiling(pagesLong);

                sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
                sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
                sheet.PageSetup.BottomMargin = 0.25;
                sheet.PageSetup.TopMargin = 0.25;
                sheet.PageSetup.LeftMargin = 0.25;
                sheet.PageSetup.RightMargin = 0.25;
                sheet.PageSetup.HeaderMargin = 0;
                sheet.PageSetup.FooterMargin = 0;
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = pages;
                return pages;
            }

            // Excel sheet names are at most 31 characters, without some symbols, and unique
            string SheetName(string wanted)
            {
                var clean = new string(wanted.Where(c => !"[]:*?/\\".Contains(c)).ToArray()).Trim();
                clean = clean.Length == 0 ? "Project" : clean.Length > 31 ? clean.Substring(0, 31) : clean;
                var name = clean;
                for (var n = 2; !sheetNames.Add(name); n++)
                {
                    var suffix = $" ({n})";
                    name = clean.Substring(0, Math.Min(clean.Length, 31 - suffix.Length)) + suffix;
                }
                return name;
            }

            var sheetIndex = 0;
            if (report.IncludeSummary)
            {
                var sheet = AddLogoToSheet(workbook, sheetIndex++, groupName, section, 4);
                sheet.Name = SheetName("Summary");
                IStyle headingStyle = workbook.Styles["headingStyle"];
                int rowNumber = 1;

                rowNumber++;
                sheet.Range[rowNumber, 2].Text = "Special Interest Area Progress";
                sheet.Range[rowNumber, 2, rowNumber, 4].Merge();
                sheet.Range[rowNumber, 2].CellStyle = headingStyle;
                sheet.SetRowHeight(rowNumber, 40);

                rowNumber++;
                sheet.Range[rowNumber, 2].Text = report.UnitName;
                sheet.Range[rowNumber, 2, rowNumber, 4].Merge();
                sheet.Range[rowNumber, 2].CellStyle = headingStyle;
                sheet.SetRowHeight(rowNumber, 30);

                rowNumber = RoomForLogo(sheet, rowNumber, 100);

                rowNumber += 2;
                var headings = new[] { "Name", "Completed", "In progress", "Areas completed" };
                for (var c = 0; c < headings.Length; c++)
                {
                    var cell = sheet.Range[rowNumber, c + 1];
                    cell.Text = headings[c];
                    cell.CellStyle.Font.Bold = true;
                    cell.CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    cell.CellStyle.HorizontalAlignment = c == 0 ? ExcelHAlign.HAlignLeft : ExcelHAlign.HAlignCenter;
                }
                sheet.SetRowHeight(rowNumber, 20);
                var headingRow = rowNumber;

                foreach (var row in report.Summary)
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = row.MemberName;
                    sheet.Range[rowNumber, 2].Text = report.SiaRequired > 0 ? $"{row.Completed}/{report.SiaRequired}" : row.Completed.ToString();
                    sheet.Range[rowNumber, 3].Text = row.InProgress.ToString();
                    sheet.Range[rowNumber, 4].Text = report.AreasRequired > 0 ? $"{row.AreasCompleted}/{report.AreasRequired}" : row.AreasCompleted.ToString();
                    sheet.Range[rowNumber, 2, rowNumber, 4].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                }
                if (report.Summary.Count == 0)
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = "No members to show";
                }

                rowNumber += 2;
                sheet.Range[rowNumber, 1, rowNumber, 4].Merge();
                sheet.Range[rowNumber, 1].Text = report.SiaRequired > 0
                    ? $"The {report.AwardName} asks for {report.SiaRequired} Special Interest Area projects over at least {report.AreasRequired} areas. Completed projects are awarded ones; in progress is every other project started in this section."
                    : "Completed projects are awarded ones; in progress is every other project started in this section.";
                sheet.Range[rowNumber, 1].CellStyle.WrapText = true;
                sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                sheet.SetRowHeight(rowNumber, 45);

                sheet.SetColumnWidth(1, 36);
                sheet.SetColumnWidth(2, 16);
                sheet.SetColumnWidth(3, 16);
                sheet.SetColumnWidth(4, 20);
                var pages = SetUpPage(sheet, rowNumber, 4);
                // a second page starts with the headings again
                if (pages > 1)
                    sheet.PageSetup.PrintTitleRows = $"${headingRow}:${headingRow}";
                printed.Add(("summary", sheet, pages));
            }

            foreach (var project in report.Projects)
            {
                const int questionWidth = 44;
                const int answerWidth = 62;
                var sheet = AddLogoToSheet(workbook, sheetIndex++, groupName, section, 2);
                sheet.Name = SheetName($"{project.MemberName} - {project.ProjectName}");
                IStyle headingStyle = workbook.Styles["headingStyle"];
                int rowNumber = 1;

                rowNumber++;
                sheet.Range[rowNumber, 2].Text = "Special Interest Area Project";
                sheet.Range[rowNumber, 2].CellStyle = headingStyle;
                sheet.SetRowHeight(rowNumber, 40);

                rowNumber++;
                sheet.Range[rowNumber, 2].Text = project.MemberName;
                sheet.Range[rowNumber, 2].CellStyle = headingStyle;
                sheet.SetRowHeight(rowNumber, 30);

                // rows 1 to 3 are 100 points tall
                rowNumber = RoomForLogo(sheet, rowNumber, 100);
                var headerRows = rowNumber;

                void Block(string heading)
                {
                    rowNumber += 2;
                    sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
                    sheet.Range[rowNumber, 1].Text = heading;
                    sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.Font.Size = 12;
                    sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    sheet.SetRowHeight(rowNumber, 20);
                }

                // A row can't be taller than Excel allows (409 points), so a very long answer carries on in the rows below it
                void Row(string label, string text, bool bold = true, double minHeight = 15)
                {
                    const int charsPerLine = answerWidth - 2;
                    const int linesPerRow = 25;
                    var chunks = new List<string>();
                    var current = new List<string>();
                    var currentLines = 0;
                    foreach (var paragraph in text.Split('\n'))
                    {
                        // a paragraph is cut at a space when it is longer than a row holds
                        var rest = paragraph;
                        do
                        {
                            var room = (linesPerRow - currentLines) * charsPerLine;
                            var piece = rest;
                            if (rest.Length > room && currentLines < linesPerRow)
                            {
                                var cut = rest.LastIndexOf(' ', Math.Min(room, rest.Length - 1));
                                piece = rest.Substring(0, cut > 0 ? cut : room);
                            }
                            var lines = Math.Max(1, (int)Math.Ceiling(piece.Length / (double)charsPerLine));
                            if (currentLines + lines > linesPerRow && current.Count > 0)
                            {
                                chunks.Add(string.Join("\n", current));
                                current = new List<string>();
                                currentLines = 0;
                                continue;
                            }
                            current.Add(piece);
                            currentLines += lines;
                            rest = rest.Substring(piece.Length).TrimStart();
                        } while (rest.Length > 0);
                    }
                    chunks.Add(string.Join("\n", current));

                    for (var i = 0; i < chunks.Count; i++)
                    {
                        var chunk = chunks[i];
                        var rowLabel = i == 0 ? label : "";
                        rowNumber++;
                        sheet.Range[rowNumber, 1].Text = rowLabel;
                        sheet.Range[rowNumber, 1].CellStyle.Font.Bold = bold;
                        sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                        sheet.Range[rowNumber, 1].CellStyle.WrapText = true;
                        sheet.Range[rowNumber, 2].Text = chunk;
                        sheet.Range[rowNumber, 2].CellStyle.WrapText = true;
                        sheet.Range[rowNumber, 2].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                        var lines = Math.Max(
                            (int)Math.Ceiling(rowLabel.Length / (double)(questionWidth - 12)),
                            chunk.Split('\n').Sum(l => Math.Max(1, (int)Math.Ceiling(l.Length / (double)charsPerLine))));
                        sheet.SetRowHeight(rowNumber, Math.Min(400, Math.Max(minHeight, lines * 15)));
                    }
                }

                void Questions(List<SiaAnswer> answers)
                {
                    // an unanswered question is left as room to write the answer in by hand
                    foreach (var item in answers)
                        Row(item.Question, item.Answer?.Trim() ?? "", bold: false, minHeight: string.IsNullOrWhiteSpace(item.Answer) ? 30 : 15);
                }

                Block("Project");
                Row("Name", project.ProjectName);
                Row("Special Interest Area", project.Area);
                Row("Where it is up to", project.Status);
                Row("Last updated", project.StatusUpdated.ToString(dateFormat, culture));
                if (project.DateAwarded.HasValue)
                    Row("Awarded", project.DateAwarded.Value.ToString(dateFormat, culture));

                Block("Plan");
                Questions(project.Plan);

                Block("Review");
                Questions(project.Review);

                Block("Files attached");
                Row("Files", project.Files.Count == 0 ? "None" : string.Join("\n", project.Files));

                Block("Notes to talk through");
                Row("Next steps", "", minHeight: 60);
                Row("Leader comments", "", minHeight: 60);

                sheet.SetColumnWidth(1, questionWidth);
                sheet.SetColumnWidth(2, answerWidth);
                var pages = SetUpPage(sheet, rowNumber, 2);
                // a second page says whose it is; the rows include the one the logo sits in, or its top is cut off
                if (pages > 1)
                    sheet.PageSetup.PrintTitleRows = $"$1:${headerRows}";
                printed.Add((project.MemberId, sheet, pages));
            }

            if (report.DoubleSided)
            {
                // a blank page goes before whoever would start on the back of a sheet
                var pagesSoFar = 0;
                var previousOwner = "";
                foreach (var (owner, sheet, pages) in printed)
                {
                    if (owner != previousOwner && pagesSoFar % 2 == 1)
                    {
                        var blank = workbook.Worksheets.Create(SheetName("Blank"));
                        // a space is enough for the page to be printed
                        blank.Range[1, 1].Text = " ";
                        blank.Move(sheet.Index);
                        pagesSoFar++;
                    }
                    pagesSoFar += pages;
                    previousOwner = owner;
                }
            }

            if (printed.Count == 0)
            {
                var sheet = AddLogoToSheet(workbook, 0, groupName, section, 2);
                sheet.Range[3, 2].Text = "No Special Interest Area projects to show";
            }

            return workbook;
        }

        public IWorkbook CreateWorkbookWithSheets(int sheetsToCreate)
        {
            //Step 1 : Instantiate the spreadsheet creation engine.
            ExcelEngine excelEngine = new ExcelEngine();
            //Step 2 : Instantiate the excel application object.
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;

            //Initializes the SubstituteFont event to perform font substitution in Excel-to-PDF conversion.
            application.SubstituteFont += new SubstituteFontEventHandler(SubstituteFont);

            // Creating new workbook
            IWorkbook workbook = application.Workbooks.Create(sheetsToCreate);

            //Adding cell style.               
            IStyle headingStyle = workbook.Styles.Add("headingStyle");
            headingStyle.Font.Bold = true;
            headingStyle.Font.Size = 14;
            headingStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            headingStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

            return workbook;
        }

        private static void SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            var directory = Directory.GetCurrentDirectory();
            args.AlternateFontStream = File.OpenRead($@"{directory}/Fonts/carlito.regular.ttf");
        }

        private IWorksheet AddLogoToSheet(IWorkbook workbook, int worksheetIndex, string groupName, string section, int lastHeadingCol)
        {
            IWorksheet sheet = workbook.Worksheets[worksheetIndex];

            // Add Logo
            var directory = Directory.GetCurrentDirectory();
            var logoName = GetLogoForSection(section);
            var logoFullPathName = $@"{directory}/Images/{logoName}";
            if (File.Exists(logoFullPathName))
            {
                FileStream imageStream = new FileStream(logoFullPathName, FileMode.Open, FileAccess.Read);
                IPictureShape logo = sheet.Pictures.AddPicture(1, 1, imageStream);
                var aspectRatio = (double)logo.Height / logo.Width;
                logo.Width = 100;
                logo.Height = (int)(100 * aspectRatio);
            }

            // Add Group Name
            var groupNameCell = sheet.Range[1, 2];
            groupNameCell.Text = groupName;
            groupNameCell.CellStyle = workbook.Styles["headingStyle"];
            sheet.Range[1, 2, 1, lastHeadingCol].Merge();
            sheet.SetRowHeight(1, 30);

            return sheet;
        }

        public string GetLogoForSection(string section)
        {
            var logoName = "";
            switch (section)
            {
                case "joey":
                    logoName = "Joey Scouts Full Col Vertical.jpg";
                    break;
                case "cub":
                    logoName = "Cub Scouts Full Col Vertical.png";
                    break;
                case "scout":
                    logoName = "Scouts Full Col Vertical.jpg";
                    break;
                case "venturer":
                    logoName = "Venturer Scouts Full Col Vertical.jpg";
                    break;
                case "rover":
                    logoName = "Rover Scouts Full Col Vertical.jpg";
                    break;
            }
            return logoName;
        }

        private IWorkbook CreateWorkbookWithLogo(string groupName, string section, int lastHeadingCol)
        {
            var workbook = CreateWorkbookWithSheets(1);
            AddLogoToSheet(workbook, 0, groupName, section, lastHeadingCol);
            return workbook;
        }
        public IWorkbook GenerateAdditionalAwardsWorkbook(List<AdditionalAwardSpecificationListModel> awardSpecificationsList, List<AdditionalAwardListModel> sortedAdditionalAwardsList, List<string>? distinctAwards, string groupName, string section, string unitName)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 16);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;
            int columnOffset = 3;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"Additional Badges Awarded as at {DateTime.Now.ToShortDateString()}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 16].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 16].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            //Add Heading Row
            rowNumber++;
            var columnNumber = 1;
            sheet.Range[rowNumber, 1].Text = "Scout";
            sheet.Range[rowNumber, 1].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            sheet.Range[rowNumber, 2].Text = "KMs";
            sheet.Range[rowNumber, 2].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            sheet.Range[rowNumber, 3].Text = "Nights";
            sheet.Range[rowNumber, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            var usedAwards = awardSpecificationsList.Where(x => distinctAwards.Contains(x.id));
            foreach (var award in usedAwards.OrderBy(x => x.additionalAwardSortIndex))
            {
                columnNumber = distinctAwards.IndexOf(award.id) + columnOffset;
                var cell = sheet.Range[rowNumber, columnNumber];
                cell.Text = " " + award.name;
                cell.CellStyle.Rotation = 90;
                cell.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                cell.CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom;
                cell.BorderAround();
            }
            rowNumber++;
            // Add detail rows
            foreach (var additionalAward in sortedAdditionalAwardsList.GroupBy(x => x.MemberName))
            {
                sheet.SetRowHeight(rowNumber, 15);
                sheet.Range[rowNumber, 1].Text = additionalAward.Key;
                sheet.Range[rowNumber, 1].BorderAround();
                sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                sheet.Range[rowNumber, 2].Text = additionalAward.FirstOrDefault().KMsHiked.ToString();
                sheet.Range[rowNumber, 2].BorderAround();
                sheet.Range[rowNumber, 2].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                sheet.Range[rowNumber, 3].Text = additionalAward.FirstOrDefault().NightsCamped.ToString();
                sheet.Range[rowNumber, 3].BorderAround();
                sheet.Range[rowNumber, 3].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                // Set style for date boxes
                for (int i = 0; i < distinctAwards.Count(); i++)
                {
                    var cell = sheet.Range[rowNumber, i + columnOffset];
                    cell.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    cell.CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                    cell.BorderAround();
                }
                // Populate award dates
                foreach (var x in additionalAward)
                {
                    columnNumber = distinctAwards.IndexOf(x.AwardId) + columnOffset;
                    var cell = sheet.Range[rowNumber, columnNumber];
                    if (x.AwardDate.HasValue)
                    {
                        cell.DateTime = x.AwardDate.Value;
                        cell.CellStyle.Color = Color.Orange;
                        if (x.PresentedDate.HasValue)
                        {
                            cell.DateTime = x.PresentedDate.Value;
                            cell.CellStyle.Color = Color.LightGreen;
                        }
                        cell.NumberFormat = "dd/MM/yy";
                    }
                    else
                    {
                        if (x.AwardId.StartsWith("walkabout"))
                        {
                            if (x.KMsHiked >= x.AwardUnits)
                            {
                                cell.CellStyle.Color = Color.Red;
                                sheet.Range[rowNumber, 2].CellStyle.Color = Color.Red;
                            }
                        }
                        if (x.AwardId.StartsWith("camper"))
                        {
                            if (x.NightsCamped >= x.AwardUnits)
                            {
                                cell.CellStyle.Color = Color.Red;
                                sheet.Range[rowNumber, 3].CellStyle.Color = Color.Red;
                            }
                        }
                    }
                }
                rowNumber++;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Not awarded";
            sheet.Range[rowNumber, 1].BorderAround();
            sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            sheet.Range[rowNumber, 2].BorderAround();
            sheet.Range[rowNumber, 2].CellStyle.Color = Color.Red;

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Presented";
            sheet.Range[rowNumber, 1].BorderAround();
            sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            sheet.Range[rowNumber, 2].BorderAround();
            sheet.Range[rowNumber, 2].CellStyle.Color = Color.LightGreen;


            sheet.UsedRange.AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA3;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.IsFitToPage = true;

            return workbook;
        }

        public IWorkbook GeneratePatrolListWorkbook(List<MemberListModel> sortedPatrolList, string groupName, string section, string unitName, bool includeLeaders)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 9);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"Patrols as at {DateTime.Now.ToShortDateString()}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 9].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 9].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            var groupedMembersList = sortedPatrolList.Where(pl => pl.isAdultLeader == 0).GroupBy(pl => pl.patrol_name).ToList();
            foreach (var patrol in groupedMembersList)
            {
                // Add Patrol Name
                var patrolName = string.IsNullOrEmpty(patrol.Key) ? "Unassigned" : patrol.Key;
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = $"{patrolName} ({patrol.Count()})";
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1].CellStyle.Font.Size = 14;
                sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                sheet.Range[rowNumber, 1, rowNumber, 3].Merge();
                sheet.SetRowHeight(rowNumber, 20);

                // Add Patrol members
                foreach (var member in patrol.OrderByDescending(p => p.patrol_duty))
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = member.first_name;
                    sheet.Range[rowNumber, 1].BorderAround();
                    sheet.Range[rowNumber, 2].Text = member.last_name;
                    sheet.Range[rowNumber, 2].BorderAround();
                    sheet.Range[rowNumber, 3].Text = member.patrol_duty;
                    sheet.Range[rowNumber, 3].BorderAround();
                    if (!string.IsNullOrEmpty(member.patrol_duty))
                    {
                        sheet.Range[rowNumber, 1, rowNumber, 3].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    }
                }
                rowNumber++;
            }

            rowNumber = 3;
            var unitCouncilList = sortedPatrolList.Where(pl => pl.unit_council)
                .OrderBy(uc => uc.isAdultLeader)
                .ThenBy(uc => uc.first_name)
                .ThenBy(uc => uc.last_name)
                .ToList();
            // Add Unit Council Name
            rowNumber++;
            sheet.Range[rowNumber, 6].Text = "Unit Council";
            sheet.Range[rowNumber, 6].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 6].CellStyle.Font.Size = 14;
            sheet.Range[rowNumber, 6].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            sheet.Range[rowNumber, 6, rowNumber, 9].Merge();
            sheet.SetRowHeight(rowNumber, 20);

            foreach (var member in unitCouncilList)
            {
                // Add Unit Council members
                rowNumber++;
                sheet.Range[rowNumber, 6].Text = member.first_name;
                sheet.Range[rowNumber, 6].BorderAround();
                sheet.Range[rowNumber, 7].Text = member.last_name;
                sheet.Range[rowNumber, 7].BorderAround();
                sheet.Range[rowNumber, 8].Text = member.patrol_duty;
                sheet.Range[rowNumber, 8].BorderAround();
            }
            rowNumber++;

            if (includeLeaders)
            {
                var adultLeaderList = sortedPatrolList
                    .Where(uc => uc.isAdultLeader == 1)
                    .OrderBy(uc => uc.first_name)
                    .ThenBy(uc => uc.last_name)
                    .ToList();
                // Add Adult Leaders Name
                rowNumber++;
                sheet.Range[rowNumber, 6].Text = "Adult Leaders";
                sheet.Range[rowNumber, 6].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 6].CellStyle.Font.Size = 14;
                sheet.Range[rowNumber, 6].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                sheet.Range[rowNumber, 6, rowNumber, 9].Merge();
                sheet.SetRowHeight(rowNumber, 20);

                foreach (var member in adultLeaderList)
                {
                    // Add Unit Council members
                    rowNumber++;
                    sheet.Range[rowNumber, 6].Text = member.first_name;
                    sheet.Range[rowNumber, 6].BorderAround();
                    sheet.Range[rowNumber, 7].Text = member.last_name;
                    sheet.Range[rowNumber, 7].BorderAround();
                    sheet.Range[rowNumber, 8].Text = member.patrol_duty;
                    sheet.Range[rowNumber, 8].BorderAround();
                }
            }

            sheet.UsedRange.AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.IsFitToPage = true;

            return workbook;
        }

        public IWorkbook GenerateMemberListWorkbook(List<MemberListModel> sortedMemberList, string groupName, string section, string unitName)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 7);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"Members as at {DateTime.Now.ToShortDateString()}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            // Add Header
            rowNumber++;
            sheet.Range[rowNumber, 2].Text = "First Name";
            sheet.Range[rowNumber, 2].BorderAround();
            sheet.Range[rowNumber, 3].Text = "Last Name";
            sheet.Range[rowNumber, 3].BorderAround();
            sheet.Range[rowNumber, 4].Text = "Age";
            sheet.Range[rowNumber, 4].BorderAround();
            sheet.Range[rowNumber, 5].Text = "Member";
            sheet.Range[rowNumber, 5].BorderAround();
            sheet.Range[rowNumber, 6].Text = "Duty";
            sheet.Range[rowNumber, 6].BorderAround();
            sheet.Range[rowNumber, 7].Text = "Patrol";
            sheet.Range[rowNumber, 7].BorderAround();
            sheet.Range[rowNumber, 2, rowNumber, 7].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;

            foreach (var member in sortedMemberList)
            {
                rowNumber++;
                sheet.Range[rowNumber, 2].Text = member.first_name;
                sheet.Range[rowNumber, 2].BorderAround();
                sheet.Range[rowNumber, 3].Text = member.last_name;
                sheet.Range[rowNumber, 3].BorderAround();
                sheet.Range[rowNumber, 4].Text = member.age;
                sheet.Range[rowNumber, 4].BorderAround();
                sheet.Range[rowNumber, 5].Text = member.member_number;
                sheet.Range[rowNumber, 5].BorderAround();
                sheet.Range[rowNumber, 6].Text = member.patrol_duty;
                sheet.Range[rowNumber, 6].BorderAround();
                sheet.Range[rowNumber, 7].Text = member.patrol_name.Trim();
                sheet.Range[rowNumber, 7].BorderAround();

                decimal.TryParse(member.age, out decimal approxYears);
                if (member.age.StartsWith((UnitMaxAge(section) - 2).ToString()))
                    sheet.Range[rowNumber, 4].CellStyle.ColorIndex = ExcelKnownColors.Yellow;
                if (member.age.StartsWith(UnitMaxAge(section).ToString()) || member.age.StartsWith((UnitMaxAge(section) - 1).ToString()))
                    sheet.Range[rowNumber, 4].CellStyle.ColorIndex = ExcelKnownColors.Rose;
            }

            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 2].Text = "Second last year in section";
            sheet.Range[rowNumber, 2, rowNumber, 3].Merge();
            sheet.Range[rowNumber, 4].Text = "";
            sheet.Range[rowNumber, 4].BorderAround();
            sheet.Range[rowNumber, 4].CellStyle.ColorIndex = ExcelKnownColors.Yellow;
            rowNumber++;
            sheet.Range[rowNumber, 2].Text = "Last year in section";
            sheet.Range[rowNumber, 2, rowNumber, 3].Merge();
            sheet.Range[rowNumber, 4].Text = "";
            sheet.Range[rowNumber, 4].BorderAround();
            sheet.Range[rowNumber, 4].CellStyle.ColorIndex = ExcelKnownColors.Rose;

            sheet.UsedRange.AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.FitToPagesWide = 1;

            return workbook;
        }

        public IWorkbook GeneratePatrolSheetsWorkbook(List<MemberListModel> sortedPatrolList, string section)
        {
            //Step 1 : Instantiate the spreadsheet creation engine.
            ExcelEngine excelEngine = new ExcelEngine();
            //Step 2 : Instantiate the excel application object.
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;

            //Initializes the SubstituteFont event to perform font substitution in Excel-to-PDF conversion.
            application.SubstituteFont += new SubstituteFontEventHandler(SubstituteFont);

            var groupedMembersList = sortedPatrolList.Where(pl => pl.isAdultLeader == 0).GroupBy(pl => pl.patrol_name).ToList();

            // Creating new workbook
            IWorkbook workbook = application.Workbooks.Create(groupedMembersList.Count);

            //Adding cell style.               
            IStyle headingStyle = workbook.Styles.Add("headingStyle");
            headingStyle.Font.Bold = true;
            headingStyle.Font.Size = 30;
            headingStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            headingStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            headingStyle.WrapText = true;

            IStyle rowStyle = workbook.Styles.Add("rowStyle");
            rowStyle.Font.Bold = true;
            rowStyle.Font.Size = 20;
            rowStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

            var sheetIndex = 0;
            foreach (var patrol in groupedMembersList)
            {
                IWorksheet sheet = workbook.Worksheets[sheetIndex];

                int rowNumber = 1;
                int columnNumber = 1;

                // Add Logo
                var directory = Directory.GetCurrentDirectory();
                var logoName = GetLogoForSection(section);
                var logoFullPathName = $@"{directory}/Images/{logoName}";
                if (File.Exists(logoFullPathName))
                {
                    FileStream imageStream = new FileStream(logoFullPathName, FileMode.Open, FileAccess.Read);
                    IPictureShape logo = sheet.Pictures.AddPicture(rowNumber, 1, imageStream);
                    var aspectRatio = (double)logo.Height / logo.Width;
                    logo.Width = 100;
                    logo.Height = (int)(100 * aspectRatio);
                }

                // Add Patrol Name
                sheet.Range[rowNumber, 2].Text = patrol.Key;
                sheet.Range[rowNumber, 2].CellStyle = headingStyle;
                sheet.Range[rowNumber, 2, rowNumber, 6].Merge();
                sheet.SetRowHeight(rowNumber, 100);

                // Add Patrol members
                foreach (var member in patrol.OrderByDescending(p => p.patrol_duty))
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 1, rowNumber, 5].CellStyle = rowStyle;
                    sheet.Range[rowNumber, 2].Text = member.first_name;
                    sheet.Range[rowNumber, 2].BorderAround();
                    sheet.Range[rowNumber, 3].Text = member.last_name;
                    sheet.Range[rowNumber, 3].BorderAround();
                    sheet.Range[rowNumber, 4].Text = member.patrol_duty;
                    sheet.Range[rowNumber, 4].BorderAround();
                    if (!string.IsNullOrEmpty(member.patrol_duty))
                    {
                        sheet.Range[rowNumber, 2, rowNumber, 4].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    }
                }
                rowNumber++;

                sheet.UsedRange.AutofitColumns();

                sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
                sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
                sheet.PageSetup.BottomMargin = 0.25;
                sheet.PageSetup.TopMargin = 0.25;
                sheet.PageSetup.LeftMargin = 0.25;
                sheet.PageSetup.RightMargin = 0.25;
                sheet.PageSetup.HeaderMargin = 0;
                sheet.PageSetup.FooterMargin = 0;
                sheet.PageSetup.IsFitToPage = true;

                sheetIndex++;
            }


            return workbook;

        }

        public IWorkbook GenerateSignInSheetWorkbook(List<MemberListModel> memberListModel, string groupName, string section, string unitName, string eventName, SignInSheetOptions options)
        {
            // Optional narrow columns with sideways headings straight after Role, then the attendance columns:
            // time in / signature / time out / signature, or a single tick box.
            var infoHeadings = new List<string>();
            if (options.IncludeRegisteredColumn)
                infoHeadings.Add("Registered");
            if (options.IncludePaidColumn)
                infoHeadings.Add("Paid");
            var attendanceHeadings = options.UseTimeInOutSignature
                ? new[] { "Time In", "Signature", "Time Out", "Signature" }
                : new[] { "Attended" };
            const int textColumns = 3; // Name, Patrol, Role
            var infoFirstColumn = textColumns + 1;
            var attendanceFirstColumn = infoFirstColumn + infoHeadings.Count;
            var lastColumn = attendanceFirstColumn + attendanceHeadings.Length - 1;

            // Optional second block beside the main table (after a narrow spacer column): the person's
            // name again plus tick columns for how they took part in the activity.
            var participationHeadings = new[] { "Participate", "Assist", "Lead" };
            var participationNameColumn = lastColumn + 2;
            var participationLastColumn = participationNameColumn + participationHeadings.Length;

            var workbook = CreateWorkbookWithLogo(groupName, section, lastColumn);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, lastColumn].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            // the row stays when no event was chosen (blank title) so the table still starts below the logo
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = eventName;
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, lastColumn].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            var displayNames = SignInSheetLayout.GetDisplayNames(memberListModel, options.UseFullNames);

            void WriteHeader()
            {
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = "Name";
                sheet.Range[rowNumber, 2].Text = "Patrol";
                sheet.Range[rowNumber, 3].Text = "Role";
                for (var i = 0; i < infoHeadings.Count; i++)
                {
                    var heading = sheet.Range[rowNumber, infoFirstColumn + i];
                    heading.Text = infoHeadings[i];
                    heading.CellStyle.Rotation = 90;
                }
                for (var i = 0; i < attendanceHeadings.Length; i++)
                    sheet.Range[rowNumber, attendanceFirstColumn + i].Text = attendanceHeadings[i];
                sheet.Range[rowNumber, 1, rowNumber, lastColumn].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1, rowNumber, lastColumn].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                for (var c = 1; c <= lastColumn; c++)
                    sheet.Range[rowNumber, c].BorderAround();

                if (options.IncludeParticipationColumns)
                {
                    sheet.Range[rowNumber, participationNameColumn].Text = "Name";
                    for (var i = 0; i < participationHeadings.Length; i++)
                    {
                        var heading = sheet.Range[rowNumber, participationNameColumn + 1 + i];
                        heading.Text = participationHeadings[i];
                        heading.CellStyle.Rotation = 90;
                    }
                    sheet.Range[rowNumber, participationNameColumn, rowNumber, participationLastColumn].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, participationNameColumn, rowNumber, participationLastColumn].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    for (var c = participationNameColumn; c <= participationLastColumn; c++)
                        sheet.Range[rowNumber, c].BorderAround();
                }

                // tall enough for the sideways headings
                if (infoHeadings.Count > 0 || options.IncludeParticipationColumns)
                    sheet.SetRowHeight(rowNumber, 62);
            }

            void WriteMembers(IEnumerable<MemberListModel> members)
            {
                foreach (var member in members)
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = displayNames[member.id];
                    sheet.Range[rowNumber, 2].Text = member.patrol_name;
                    sheet.Range[rowNumber, 3].Text = member.patrol_duty;
                    for (var c = 1; c <= lastColumn; c++)
                        sheet.Range[rowNumber, c].BorderAround();
                    if (options.IncludeParticipationColumns)
                    {
                        sheet.Range[rowNumber, participationNameColumn].Text = displayNames[member.id];
                        for (var c = participationNameColumn; c <= participationLastColumn; c++)
                            sheet.Range[rowNumber, c].BorderAround();
                    }
                    // taller rows so there is room to write a signature
                    sheet.SetRowHeight(rowNumber, options.UseTimeInOutSignature ? 24 : 18);
                }
            }

            // Youth by section then leaders (see SignInSheetLayout); a labelled heading row above each group.
            var groups = SignInSheetLayout.GetGroups(memberListModel, displayNames);
            if (!groups.Any())
                WriteHeader();
            foreach (var group in groups)
            {
                if (group.Heading != null)
                {
                    if (group != groups.First())
                        rowNumber++; // a blank row between sections
                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = group.Heading;
                    sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 1].CellStyle.Font.Size = 12;
                }
                WriteHeader();
                WriteMembers(group.Members);
            }

            // Columns that hold text fit their content. The blank write-in columns autofit to nothing, so they
            // start at a writable minimum. The page is portrait when the columns fit across it, otherwise
            // landscape, and whatever page width is left over (in column width units) is shared out so the
            // sheet fills the page instead of leaving a gap on the right.
            sheet.UsedRange.AutofitColumns();
            const double portraitPageWidth = 97;
            const double landscapePageWidth = 140;
            const double tickColumnWidth = 5; // the narrow columns with sideways headings
            for (var c = infoFirstColumn; c < attendanceFirstColumn; c++)
                sheet.SetColumnWidth(c, tickColumnWidth);
            if (options.IncludeParticipationColumns)
            {
                // autofit sizes the spacer and the sideways-headed tick columns badly
                sheet.SetColumnWidth(lastColumn + 1, 3);
                for (var c = participationNameColumn + 1; c <= participationLastColumn; c++)
                    sheet.SetColumnWidth(c, tickColumnWidth);
            }
            if (options.UseTimeInOutSignature)
            {
                for (var c = attendanceFirstColumn; c <= lastColumn; c++)
                    sheet.SetColumnWidth(c, (c - attendanceFirstColumn) % 2 == 0 ? 10 : 22);
            }
            else
            {
                sheet.SetColumnWidth(attendanceFirstColumn, 10);
            }

            var usedWidth = 0.0;
            var lastUsedColumn = options.IncludeParticipationColumns ? participationLastColumn : lastColumn;
            for (var c = 1; c <= lastUsedColumn; c++)
                usedWidth += sheet.GetColumnWidth(c);
            var portrait = usedWidth <= portraitPageWidth;
            var spare = (portrait ? portraitPageWidth : landscapePageWidth) - usedWidth;
            if (spare > 0)
            {
                if (options.UseTimeInOutSignature)
                {
                    // signatures need the most room: weight them over the times
                    var weights = Enumerable.Range(0, attendanceHeadings.Length).Select(i => i % 2 == 0 ? 1.0 : 2.2).ToList();
                    for (var i = 0; i < weights.Count; i++)
                    {
                        var c = attendanceFirstColumn + i;
                        sheet.SetColumnWidth(c, sheet.GetColumnWidth(c) + spare * weights[i] / weights.Sum());
                    }
                }
                else
                {
                    // a single tick box: spread the extra over the name/patrol/role and tick box columns
                    var growing = Enumerable.Range(1, textColumns).Append(attendanceFirstColumn).ToList();
                    var growingWidth = growing.Sum(c => sheet.GetColumnWidth(c));
                    foreach (var c in growing)
                        sheet.SetColumnWidth(c, sheet.GetColumnWidth(c) * (1 + spare / growingWidth));
                }
            }

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = portrait ? ExcelPageOrientation.Portrait : ExcelPageOrientation.Landscape;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            return workbook;
        }

        public IWorkbook GenerateEventAttendanceWorkbook(EventListModel eventListModel, string groupName, string section, string unitName)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 7);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 7].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = eventListModel.EventDisplay;
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 7].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            // Add Header
            rowNumber++;
            sheet.Range[rowNumber, 2, rowNumber, 5].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
            sheet.Range[rowNumber, 2, rowNumber, 5].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = "First Name";
            sheet.Range[rowNumber, 2].BorderAround();
            sheet.Range[rowNumber, 3].Text = "Last Name";
            sheet.Range[rowNumber, 3].BorderAround();
            sheet.Range[rowNumber, 4].Text = "Patrol";
            sheet.Range[rowNumber, 4].BorderAround();
            sheet.Range[rowNumber, 5].Text = "Attended";
            sheet.Range[rowNumber, 5].BorderAround();

            foreach (var attendee in eventListModel.attendees.Where(a => !a.isAdultMember).OrderBy(a => a.last_name))
            {
                rowNumber++;
                sheet.Range[rowNumber, 2].Text = attendee.first_name;
                sheet.Range[rowNumber, 2].BorderAround();
                sheet.Range[rowNumber, 3].Text = attendee.last_name;
                sheet.Range[rowNumber, 3].BorderAround();
                sheet.Range[rowNumber, 4].Text = attendee.patrol_name;
                sheet.Range[rowNumber, 4].BorderAround();
                sheet.Range[rowNumber, 5].Text = attendee.attended ? "Y" : "";
                sheet.Range[rowNumber, 5].BorderAround();
                sheet.Range[rowNumber, 5].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            }

            rowNumber++;

            foreach (var attendee in eventListModel.attendees.Where(a => a.isAdultMember).OrderBy(a => a.last_name))
            {
                rowNumber++;
                sheet.Range[rowNumber, 2].Text = attendee.first_name;
                sheet.Range[rowNumber, 2].BorderAround();
                sheet.Range[rowNumber, 3].Text = attendee.last_name;
                sheet.Range[rowNumber, 3].BorderAround();
                sheet.Range[rowNumber, 4].Text = attendee.patrol_name;
                sheet.Range[rowNumber, 4].BorderAround();
                sheet.Range[rowNumber, 5].Text = attendee.attended ? "Y" : "";
                sheet.Range[rowNumber, 5].BorderAround();
                sheet.Range[rowNumber, 5].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            }

            sheet.UsedRange.AutofitColumns();

            return workbook;
        }

        public IWorkbook GenerateAttendanceReportWorkbook(AttendanceReportModel attendanceReportData, string groupName, string section, string unitName, DateTime fromDate, DateTime toDate, bool forPdfOutput)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 10);
            IWorksheet sheet = workbook.Worksheets[0];
            sheet.EnableSheetCalculations();

            int rowNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 10].Merge();
            sheet.SetRowHeight(rowNumber, 25);

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"Attendance between {fromDate.ToShortDateString()} and {toDate.ToShortDateString()}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 10].Merge();
            sheet.SetRowHeight(rowNumber, 25);

            var totalEvents = attendanceReportData.attendanceReportChallengeAreaSummaries.FirstOrDefault()?.TotalEvents ?? 0;
            var communityEvents = attendanceReportData.attendanceReportChallengeAreaSummaries.Where(c => c.ChallengeArea == "Community").FirstOrDefault()?.EventCount ?? 0;
            var creativeEvents = attendanceReportData.attendanceReportChallengeAreaSummaries.Where(c => c.ChallengeArea == "Creative").FirstOrDefault()?.EventCount ?? 0;
            var outdoorsEvents = attendanceReportData.attendanceReportChallengeAreaSummaries.Where(c => c.ChallengeArea == "Outdoors").FirstOrDefault()?.EventCount ?? 0;
            var personalGrowthEvents = attendanceReportData.attendanceReportChallengeAreaSummaries.Where(c => c.ChallengeArea == "Growth").FirstOrDefault()?.EventCount ?? 0;

            // Add Challenge Area Summary
            rowNumber++;
            var columnNumber = forPdfOutput ? 0 : 1;
            sheet.Range[rowNumber, columnNumber + 2].Text = $"Community {communityEvents} / {totalEvents}";
            sheet.Range[rowNumber, columnNumber + 2, rowNumber, columnNumber + 3].Merge();
            sheet.Range[rowNumber, columnNumber + 2, rowNumber, columnNumber + 3].BorderAround();
            sheet.Range[rowNumber, columnNumber + 2, rowNumber, columnNumber + 3].CellStyle.ColorIndex = GetChallengeAreaColour("Community");
            sheet.Range[rowNumber, columnNumber + 4].Text = $"Creative {creativeEvents} / {totalEvents}";
            sheet.Range[rowNumber, columnNumber + 4, rowNumber, columnNumber + 5].Merge();
            sheet.Range[rowNumber, columnNumber + 4, rowNumber, columnNumber + 5].BorderAround();
            sheet.Range[rowNumber, columnNumber + 4, rowNumber, columnNumber + 5].CellStyle.ColorIndex = GetChallengeAreaColour("Creative");
            sheet.Range[rowNumber, columnNumber + 6].Text = $"Outdoors {outdoorsEvents} / {totalEvents}";
            sheet.Range[rowNumber, columnNumber + 6, rowNumber, columnNumber + 7].Merge();
            sheet.Range[rowNumber, columnNumber + 6, rowNumber, columnNumber + 7].BorderAround();
            sheet.Range[rowNumber, columnNumber + 6, rowNumber, columnNumber + 7].CellStyle.ColorIndex = GetChallengeAreaColour("Outdoors");
            sheet.Range[rowNumber, columnNumber + 8].Text = $"Personal Growth {personalGrowthEvents} / {totalEvents}";
            sheet.Range[rowNumber, columnNumber + 8, rowNumber, columnNumber + 10].Merge();
            sheet.Range[rowNumber, columnNumber + 8, rowNumber, columnNumber + 10].BorderAround();
            sheet.Range[rowNumber, columnNumber + 8, rowNumber, columnNumber + 10].CellStyle.ColorIndex = GetChallengeAreaColour("Growth");
            sheet.Range[rowNumber, columnNumber + 2, rowNumber, columnNumber + 10].CellStyle.Font.Bold = true;
            rowNumber++;

            var allEvents = attendanceReportData.attendanceReportItems.DistinctBy(i => i.EventNameDisplay).OrderBy(i => i.EventStartDate).ToList();

            // Add Event Details
            columnNumber = forPdfOutput ? 1 : 2;
            rowNumber++;
            foreach (var eventAttendance in allEvents)
            {
                columnNumber++;
                if (forPdfOutput)
                {
                    // Name
                    sheet.Range[rowNumber, columnNumber].Text = $"{eventAttendance.EventNameDisplay} ({eventAttendance.EventStatus})";
                    sheet.Range[rowNumber, columnNumber].CellStyle.ColorIndex = GetChallengeAreaColour(eventAttendance.EventChallengeArea);
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                    sheet.Range[rowNumber, columnNumber].CellStyle.Rotation = 90;
                    sheet.Range[rowNumber, columnNumber].CellStyle.WrapText = true;
                    sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
                    sheet.SetColumnWidth(columnNumber, 8);
                }
                else
                {
                    // Challenge Area
                    sheet.Range[rowNumber, columnNumber].Text = eventAttendance.EventChallengeArea;
                    sheet.Range[rowNumber, columnNumber].CellStyle.ColorIndex = GetChallengeAreaColour(eventAttendance.EventChallengeArea);
                    sheet.Range[rowNumber, columnNumber].BorderAround();

                    // Date
                    sheet.Range[rowNumber + 1, columnNumber].DateTime = eventAttendance.EventStartDate;
                    sheet.Range[rowNumber + 1, columnNumber].CellStyle.ColorIndex = GetChallengeAreaColour(eventAttendance.EventChallengeArea);
                    sheet.Range[rowNumber + 1, columnNumber].BorderAround();

                    // State
                    sheet.Range[rowNumber + 2, columnNumber].Text = eventAttendance.EventStatus;
                    sheet.Range[rowNumber + 2, columnNumber].CellStyle.ColorIndex = GetChallengeAreaColour(eventAttendance.EventChallengeArea);
                    sheet.Range[rowNumber + 2, columnNumber].BorderAround();

                    // Name
                    sheet.Range[rowNumber + 3, columnNumber].Text = eventAttendance.EventName;
                    sheet.Range[rowNumber + 3, columnNumber].CellStyle.ColorIndex = GetChallengeAreaColour(eventAttendance.EventChallengeArea);
                    sheet.Range[rowNumber + 3, columnNumber].BorderAround();
                }
            }

            if (forPdfOutput)
            {
                sheet.SetRowHeight(rowNumber, 175);
            }
            else
            {
                rowNumber = rowNumber + 3;
            }
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Total";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
            sheet.SetColumnWidth(columnNumber, 8);


            // Add youth member rows
            // Group attendance by member for youth
            var groupedAttendances = attendanceReportData.attendanceReportItems.Where(m => m.IsAdultMember == 0).GroupBy(wa => wa.MemberName).ToList();
            var sumStartRow = rowNumber + 1;
            foreach (var groupedAttendance in groupedAttendances.OrderBy(a => a.Key))
            {
                // Name
                rowNumber++;
                columnNumber = 1;
                if (forPdfOutput)
                {
                    sheet.Range[rowNumber, columnNumber].Text = groupedAttendance.FirstOrDefault().MemberNameAndRate;
                    sheet.Range[rowNumber, columnNumber].BorderAround();

                }
                else
                {
                    sheet.Range[rowNumber, columnNumber].Text = groupedAttendance.FirstOrDefault().MemberFirstName;
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                    columnNumber++;
                    sheet.Range[rowNumber, columnNumber].Text = groupedAttendance.FirstOrDefault().MemberLastName;
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                }
                // Event Attendance
                foreach (var events in allEvents)
                {
                    var eventAttendance = groupedAttendance.Where(a => a.EventNameDisplay == events.EventNameDisplay).FirstOrDefault();
                    columnNumber++;
                    sheet.Range[rowNumber, columnNumber].Text = eventAttendance?.Pal;
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                    sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                }
                // Row total
                var sumRange = sheet.Range[rowNumber, forPdfOutput ? 2 : 3, rowNumber, columnNumber].AddressLocal;
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Formula = @$"=COUNTIFS({sumRange}, ""Y"")+COUNTIFS({sumRange}, ""P"")+COUNTIFS({sumRange}, ""A"")+COUNTIFS({sumRange}, ""L"")";
                sheet.Range[rowNumber, columnNumber].BorderAround();
                sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, columnNumber].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            }

            // Add total row
            var sumEndRow = rowNumber;
            rowNumber++;
            var youthTotalRow = rowNumber;
            sheet.Range[rowNumber, 1].Text = "Youth Total";
            sheet.Range[rowNumber, 1].BorderAround();
            var startCol = forPdfOutput ? 2 : 3;
            // A unit with no youth (an adults-only team) has no rows to add up, so its totals are zero
            var anyYouthRows = sumEndRow >= sumStartRow;
            for (int i = startCol; i <= columnNumber - 1; i++)
            {
                if (anyYouthRows)
                {
                    var sumRange = sheet.Range[sumStartRow, i, sumEndRow, i].AddressLocal;
                    sheet.Range[rowNumber, i].Formula = @$"=COUNTIFS({sumRange}, ""Y"")+COUNTIFS({sumRange}, ""P"")+COUNTIFS({sumRange}, ""A"")+COUNTIFS({sumRange}, ""L"")";
                }
                else
                {
                    sheet.Range[rowNumber, i].Number = 0;
                }
                sheet.Range[rowNumber, i].BorderAround();
                sheet.Range[rowNumber, i].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, i].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            }
            if (anyYouthRows)
            {
                sheet.Range[rowNumber, columnNumber].Formula = $"SUM({sheet.Range[sumStartRow, columnNumber, sumEndRow, columnNumber].AddressLocal})";
                sheet.Range[rowNumber, columnNumber].Number = sheet.Range[rowNumber, columnNumber].FormulaNumberValue;
            }
            else
            {
                sheet.Range[rowNumber, columnNumber].Number = 0;
            }
            sheet.Range[rowNumber, columnNumber].BorderAround();
            sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;

            // Group attendance by member for adults
            groupedAttendances = attendanceReportData.attendanceReportItems.Where(m => m.IsAdultMember == 1).GroupBy(wa => wa.MemberName).ToList();

            // Add adult member rows
            sumStartRow = rowNumber + 1;
            foreach (var groupedAttendance in groupedAttendances.OrderBy(a => a.Key))
            {
                rowNumber++;
                columnNumber = 1;
                if (forPdfOutput)
                {
                    sheet.Range[rowNumber, columnNumber].Text = groupedAttendance.FirstOrDefault().MemberName;
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                }
                else
                {
                    sheet.Range[rowNumber, columnNumber].Text = groupedAttendance.FirstOrDefault().MemberFirstName;
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                    columnNumber++;
                    sheet.Range[rowNumber, columnNumber].Text = groupedAttendance.FirstOrDefault().MemberLastName;
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                }
                foreach (var events in allEvents)
                {
                    var eventAttendance = groupedAttendance.Where(a => a.EventNameDisplay == events.EventNameDisplay).FirstOrDefault();
                    columnNumber++;
                    sheet.Range[rowNumber, columnNumber].Text = eventAttendance.Attended > 0 ? "Y" : "";
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                    sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                }
                // Row total
                var sumRange = sheet.Range[rowNumber, forPdfOutput ? 2 : 3, rowNumber, columnNumber].AddressLocal;
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Formula = @$"=COUNTIFS({sumRange}, ""Y"")";
                sheet.Range[rowNumber, columnNumber].BorderAround();
                sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, columnNumber].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            }

            // Add total row
            sumEndRow = rowNumber;
            rowNumber++;
            var adultTotalRow = rowNumber;
            sheet.Range[rowNumber, 1].Text = "Adult Total";
            sheet.Range[rowNumber, 1].BorderAround();
            if (!forPdfOutput)
                sheet.Range[rowNumber, 2].BorderAround();
            var anyAdultRows = sumEndRow >= sumStartRow;
            for (int i = forPdfOutput ? 2 : 3; i <= columnNumber - 1; i++)
            {
                if (anyAdultRows)
                {
                    var sumRange = sheet.Range[sumStartRow, i, sumEndRow, i].AddressLocal;
                    sheet.Range[rowNumber, i].Formula = @$"=COUNTIFS({sumRange}, ""Y"")";
                }
                else
                {
                    sheet.Range[rowNumber, i].Number = 0;
                }
                sheet.Range[rowNumber, i].BorderAround();
                sheet.Range[rowNumber, i].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, i].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            }
            if (anyAdultRows)
            {
                sheet.Range[rowNumber, columnNumber].Formula = $"SUM({sheet.Range[sumStartRow, columnNumber, sumEndRow, columnNumber].AddressLocal})";
                sheet.Range[rowNumber, columnNumber].Number = sheet.Range[rowNumber, columnNumber].FormulaNumberValue;
            }
            else
            {
                sheet.Range[rowNumber, columnNumber].Number = 0;
            }
            sheet.Range[rowNumber, columnNumber].BorderAround();
            sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;

            // Add Grand Total row
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Event Total";
            sheet.Range[rowNumber, 1].BorderAround();
            if (!forPdfOutput)
                sheet.Range[rowNumber, 2].BorderAround();
            for (int i = forPdfOutput ? 2 : 3; i <= columnNumber; i++)
            {
                var youthTotalCell = int.Parse(sheet.Range[youthTotalRow, i].CalculatedValue);
                var adultTotalCell = int.Parse(sheet.Range[adultTotalRow, i].CalculatedValue);
                sheet.Range[rowNumber, i].Number = youthTotalCell + adultTotalCell;
                sheet.Range[rowNumber, i].BorderAround();
                sheet.Range[rowNumber, i].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, i].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            }
            sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;

            sheet.Range[1, 1, rowNumber, 1].AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA3;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
            sheet.PageSetup.CenterHorizontally = true;
            sheet.PageSetup.CenterVertically = true;
            sheet.PageSetup.BottomMargin = 0.2;
            sheet.PageSetup.TopMargin = 0.2;
            sheet.PageSetup.LeftMargin = 0.2;
            sheet.PageSetup.RightMargin = 0.2;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.IsFitToPage = true;

            return workbook;
        }

        private void GenerateOASWorksheetBodyLikeTerrain(IWorksheet sheet, IList<IGrouping<string, OASWorksheetAnswers>> groupedAnswers, ref int rowNumber, ref int columnNumber)
        {
            var pdrText = "";
            var titleRow = rowNumber;
            int pdrStartRow = 99;
            foreach (var plan in groupedAnswers.FirstOrDefault().OrderBy(ga => ga.InputTitleSortIndex).ThenBy(ga => ga.InputSortIndex).ToList())
            {
                // Plan Do Review
                rowNumber++;
                if (pdrText != plan.InputTitle)
                {
                    pdrText = plan.InputTitle;
                    sheet.Range[rowNumber, columnNumber].Text = plan.InputTitle;
                    sheet.Range[rowNumber, columnNumber].CellStyle.Color = GetInputTitleColour(plan.InputTitle);
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                    sheet.Range[rowNumber, columnNumber].CellStyle.Rotation = 0;
                    sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, columnNumber].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                    if (pdrStartRow > rowNumber)
                        pdrStartRow = rowNumber;
                    else
                    {
                        sheet.Range[pdrStartRow, columnNumber, rowNumber - 1, columnNumber].Merge();
                        sheet.Range[pdrStartRow, columnNumber, rowNumber - 1, columnNumber].BorderAround();
                        pdrStartRow = rowNumber;
                    }
                }

                // iStatement
                sheet.Range[rowNumber, columnNumber + 1].Text = plan.InputLabel;
                sheet.Range[rowNumber, columnNumber + 1].BorderAround();
                sheet.Range[rowNumber, columnNumber + 1].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, columnNumber + 1].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                sheet.Range[rowNumber, columnNumber + 1].CellStyle.Rotation = 0;
                sheet.Range[rowNumber, columnNumber + 1].CellStyle.WrapText = true;
                sheet.SetColumnWidth(columnNumber + 1, 50);
                sheet.AutofitRow(rowNumber);
            }

            // Member Name and answers
            columnNumber++;
            sheet.Range[titleRow + 1, columnNumber + 1].FreezePanes();
            rowNumber = titleRow;
            foreach (var groupedAnswer in groupedAnswers.OrderBy(ga => ga.Key))
            {
                columnNumber++;
                string memberName;
                if (groupedAnswer.Key.Contains('|'))
                {
                    memberName = groupedAnswer.Key.Split("|")[0];
                }
                else
                {
                    memberName = groupedAnswer.Key;
                }
                sheet.Range[rowNumber, columnNumber].Text = memberName;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, columnNumber].CellStyle.Rotation = 90;
                sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

                foreach (var answer in groupedAnswer.OrderBy(ga => ga.InputTitleSortIndex).ThenBy(ga => ga.InputSortIndex))
                {
                    rowNumber++;
                    if (answer.Answered)
                    {
                        if (answer.Awarded)
                        {
                            if (answer.InputId == "logbook_up_to_date")
                            {
                                sheet.Range[rowNumber, columnNumber].Text = answer.MemberAnswer;

                            }
                            sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                            sheet.Range[rowNumber, columnNumber].CellStyle.Color = Color.DarkSeaGreen;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(answer.MemberAnswer))
                            {
                                sheet.Range[rowNumber, columnNumber].Text = answer.MemberAnswer;

                            }
                            else
                            {
                                sheet.Range[rowNumber, columnNumber].Text = "Y";
                            }
                            sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                            sheet.Range[rowNumber, columnNumber].CellStyle.Color = Color.Orange;
                        }
                    }
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                }
                rowNumber = titleRow;
                sheet.AutofitColumn(columnNumber);
            }


        }
        private void GenerateOASWorksheetBodyOriginal(IWorksheet sheet, IList<IGrouping<string, OASWorksheetAnswers>> groupedAnswers, ref int rowNumber, ref int columnNumber)
        {
            var pdrText = "";
            int pdrStartCol = 99;
            foreach (var plan in groupedAnswers.FirstOrDefault().OrderBy(ga => ga.InputTitleSortIndex).ThenBy(ga => ga.InputSortIndex).ToList())
            {
                // Plan Do Review
                columnNumber++;
                if (pdrText != plan.InputTitle)
                {
                    pdrText = plan.InputTitle;
                    sheet.Range[rowNumber, columnNumber].Text = plan.InputTitle;
                    sheet.Range[rowNumber, columnNumber].CellStyle.Color = GetInputTitleColour(plan.InputTitle);
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                    sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    if (pdrStartCol > columnNumber)
                        pdrStartCol = columnNumber;
                    else
                    {
                        sheet.Range[rowNumber, pdrStartCol, rowNumber, columnNumber - 1].Merge();
                        sheet.Range[rowNumber, pdrStartCol, rowNumber, columnNumber - 1].BorderAround();
                        pdrStartCol = columnNumber;
                    }
                }

                // iStatement
                sheet.Range[rowNumber + 1, columnNumber].Text = plan.InputLabel;
                sheet.Range[rowNumber + 1, columnNumber].BorderAround();
                sheet.Range[rowNumber + 1, columnNumber].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber + 1, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                sheet.Range[rowNumber + 1, columnNumber].CellStyle.Rotation = 90;
                sheet.Range[rowNumber + 1, columnNumber].CellStyle.WrapText = true;
                sheet.SetRowHeight(rowNumber + 1, 200);
                sheet.SetColumnWidth(columnNumber, 10);
            }

            // Member Name and answers
            rowNumber++;
            rowNumber++;
            columnNumber = 0;
            foreach (var groupedAnswer in groupedAnswers.OrderBy(ga => ga.Key))
            {
                columnNumber++;
                string memberName;
                if (groupedAnswer.Key.Contains('|'))
                {
                    memberName = groupedAnswer.Key.Split("|")[0];
                }
                else
                {
                    memberName = groupedAnswer.Key;
                }
                sheet.Range[rowNumber, columnNumber].Text = memberName;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                sheet.Range[rowNumber, columnNumber].CellStyle.Font.Bold = true;

                foreach (var answer in groupedAnswer.OrderBy(ga => ga.InputTitleSortIndex).ThenBy(ga => ga.InputSortIndex))
                {
                    columnNumber++;
                    if (answer.Answered)
                    {
                        if (answer.Awarded)
                        {
                            if (answer.InputId == "logbook_up_to_date")
                            {
                                sheet.Range[rowNumber, columnNumber].Text = answer.MemberAnswer;
                            }
                            sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                            sheet.Range[rowNumber, columnNumber].CellStyle.Color = Color.DarkSeaGreen;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(answer.MemberAnswer))
                            {
                                sheet.Range[rowNumber, columnNumber].Text = answer.MemberAnswer;
                            }
                            else
                            {
                                sheet.Range[rowNumber, columnNumber].Text = "Y";
                            }
                            sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                            sheet.Range[rowNumber, columnNumber].CellStyle.Color = Color.Orange;
                        }
                    }
                    sheet.Range[rowNumber, columnNumber].BorderAround();
                }
                rowNumber++;
                columnNumber = 0;
            }

            sheet.Range[1, 1, rowNumber, 1].AutofitColumns();
        }

        public IWorkbook GenerateOASWorksheetWorkbook(List<OASWorksheetAnswers> worksheetAnswers, string groupName, string section, string unitName, bool forPdfOutput, bool formatLikeTerrain, bool breakByPatrol = false)
        {
            var worksheetAnswersGroupedByTemplate = worksheetAnswers.GroupBy(wa => wa.TemplateTitle + (breakByPatrol ? " - " + wa.MemberPatrol : ""));
            var workbook = CreateWorkbookWithSheets(worksheetAnswersGroupedByTemplate.Count());
            var worksheetIndex = 0;
            foreach (var templatAnswerGroup in worksheetAnswersGroupedByTemplate.OrderBy(a => a.Key))
            {
                IWorksheet sheet = AddLogoToSheet(workbook, worksheetIndex, groupName, section, 10);
                sheet.Name = templatAnswerGroup.Key;
                int rowNumber = 1;
                int columnNumber = 1;

                IStyle headingStyle = workbook.Styles["headingStyle"];

                // Add Unit name
                rowNumber++;
                var unit = sheet.Range[rowNumber, 2];
                unit.Text = unitName;
                unit.CellStyle = headingStyle;
                sheet.Range[rowNumber, 2, rowNumber, 10].Merge();
                sheet.SetRowHeight(rowNumber, 25);

                // Add Title
                string templateTitle = templatAnswerGroup.Key ?? "";
                rowNumber++;
                var title = sheet.Range[rowNumber, 2];
                title.Text = $"{templateTitle} as at {DateTime.Now.ToShortDateString()}";
                title.CellStyle = headingStyle;
                sheet.Range[rowNumber, 2, rowNumber, 10].Merge();
                sheet.SetRowHeight(rowNumber, 25);

                rowNumber++;
                IList<IGrouping<string, OASWorksheetAnswers>> groupedAnswers = templatAnswerGroup.GroupBy(x => x.MemberName + "|" + x.MemberId).ToList();

                if (formatLikeTerrain)
                    GenerateOASWorksheetBodyLikeTerrain(sheet, groupedAnswers, ref rowNumber, ref columnNumber);
                else
                    GenerateOASWorksheetBodyOriginal(sheet, groupedAnswers, ref rowNumber, ref columnNumber);

                sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA3;
                sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
                sheet.PageSetup.BottomMargin = 0.25;
                sheet.PageSetup.TopMargin = 0.25;
                sheet.PageSetup.LeftMargin = 0.25;
                sheet.PageSetup.RightMargin = 0.25;
                sheet.PageSetup.HeaderMargin = 0;
                sheet.PageSetup.FooterMargin = 0;
                sheet.PageSetup.IsFitToPage = true;

                worksheetIndex++;
            }

            return workbook;
        }


        public IWorkbook GenerateSIAWorkbook(List<SIAProjectListModel> siaProjects, string groupName, string section, string unitName, bool forPdfOutput)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 6);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 6].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"SIA Project List as at {DateTime.Now.ToShortDateString()}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 6].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            foreach (var memberSiaProject in siaProjects.GroupBy(sp => sp.memberName))
            {
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = memberSiaProject.Key;
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
                sheet.SetRowHeight(rowNumber, 20);
                foreach (var siaProject in memberSiaProject.OrderBy(sia => sia.statusUpdated))
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 2].Text = siaProject.area;
                    sheet.Range[rowNumber, 3].Text = siaProject.projectName;
                    sheet.Range[rowNumber, 3].CellStyle.WrapText = true;
                    sheet.Range[rowNumber, 4].Text = siaProject.status;
                    sheet.Range[rowNumber, 5].DateTime = siaProject.statusUpdated;
                }
            }

            sheet.SetColumnWidth(1, 10);
            sheet.SetColumnWidth(2, 21);
            sheet.SetColumnWidth(3, 65);
            sheet.SetColumnWidth(4, 20);
            sheet.SetColumnWidth(5, 15);

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.IsFitToPage = true;

            return workbook;
        }

        public IWorkbook GenerateMilestoneWorkbook(List<MilestoneSummaryListModel> milestoneSummaries, string groupName, string section, string unitName, bool forPdfOutput, bool breakByPatrol = false)
        {
            var milestoneSummariesGroupedByPatrol = milestoneSummaries.GroupBy(ms => breakByPatrol ? ms.memberPatrol : "");
            var workbook = CreateWorkbookWithSheets(milestoneSummariesGroupedByPatrol.Count());
            var worksheetIndex = 0;

            foreach (var patrolGroup in milestoneSummariesGroupedByPatrol.OrderBy(g => g.Key))
            {
                IWorksheet sheet = AddLogoToSheet(workbook, worksheetIndex, groupName, section, 8);
                sheet.Name = string.IsNullOrEmpty(patrolGroup.Key) ? "Unit" : patrolGroup.Key;
                int rowNumber = 1;
                int averageStartRow = 0;
                int averageEndRow = 0;

                IStyle headingStyle = workbook.Styles["headingStyle"];

                // Add Unit name
                rowNumber++;
                var unit = sheet.Range[rowNumber, 2];
                unit.Text = unitName;
                unit.CellStyle = headingStyle;
                sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
                sheet.SetRowHeight(rowNumber, 40);

                // Add Title
                rowNumber++;
                var title = sheet.Range[rowNumber, 2];
                title.Text = $"{(string.IsNullOrEmpty(patrolGroup.Key) ? "" : patrolGroup.Key + " ")}Milestone Summary as at {DateTime.Now.ToShortDateString()}";
                title.CellStyle = headingStyle;
                sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
                sheet.SetRowHeight(rowNumber, 30);

                // Group key 4 is "Milestone 3 complete" - there's no Milestone 4, so a member whose
                // Milestone 3 is awarded/approved stays at currentLevel 3 (see MilestoneService), but
                // gets its own section here rather than being mixed in with those still working on it.
                foreach (var milestoneSummary in patrolGroup.GroupBy(ms => ms.currentLevel == 3 && ms.milestone3Awarded ? 4 : ms.currentLevel).OrderBy(ms => ms.Key))
                {
                    // Headings
                    rowNumber++;
                    sheet.Range[rowNumber, 3].Text = milestoneSummary.Key == 4 ? "Milestone 3 - Completed" : $"Milestone {milestoneSummary.Key}";
                    sheet.Range[rowNumber, 3].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 3].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                    sheet.Range[rowNumber, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    sheet.Range[rowNumber, 3, rowNumber, 8].Merge();
                    sheet.Range[rowNumber, 3, rowNumber, 8].BorderAround();
                    sheet.Range[rowNumber, 3, rowNumber, 8].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    sheet.SetRowHeight(rowNumber, 20);

                    rowNumber++;
                    sheet.Range[rowNumber, 3].Text = GetParticipateHeadingText(milestoneSummary.Key);
                    sheet.Range[rowNumber, 3].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 3].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                    sheet.Range[rowNumber, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    sheet.Range[rowNumber, 3, rowNumber, 6].Merge();
                    sheet.Range[rowNumber, 3, rowNumber, 6].BorderAround();
                    sheet.Range[rowNumber, 3, rowNumber, 6].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    sheet.SetRowHeight(rowNumber, 20);

                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = "Name";
                    sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 1].BorderAround();
                    sheet.Range[rowNumber, 2].Text = "Percent Complete";
                    sheet.Range[rowNumber, 2].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 2].BorderAround();
                    sheet.Range[rowNumber, 2].CellStyle.WrapText = true;
                    sheet.Range[rowNumber, 3].Text = "Community";
                    sheet.Range[rowNumber, 3].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 3].BorderAround();
                    sheet.Range[rowNumber, 4].Text = "Outdoors";
                    sheet.Range[rowNumber, 4].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 4].BorderAround();
                    sheet.Range[rowNumber, 5].Text = "Creative";
                    sheet.Range[rowNumber, 5].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 5].BorderAround();
                    sheet.Range[rowNumber, 6].Text = "Personal Growth";
                    sheet.Range[rowNumber, 6].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 6].BorderAround();
                    sheet.Range[rowNumber, 6].CellStyle.WrapText = true;
                    sheet.Range[rowNumber, 7].Text = GetLeadHeadingText(milestoneSummary.Key);
                    sheet.Range[rowNumber, 7].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    sheet.Range[rowNumber - 1, 7, rowNumber, 7].Merge();
                    sheet.Range[rowNumber - 1, 7, rowNumber, 7].BorderAround();
                    sheet.Range[rowNumber - 1, 7, rowNumber, 7].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    sheet.Range[rowNumber, 8].Text = GetLeadHeadingText(milestoneSummary.Key) + " Areas";
                    sheet.Range[rowNumber, 8].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 8].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    sheet.Range[rowNumber, 8].CellStyle.WrapText = true;
                    sheet.Range[rowNumber - 1, 8, rowNumber, 8].Merge();
                    sheet.Range[rowNumber - 1, 8, rowNumber, 8].BorderAround();
                    sheet.Range[rowNumber - 1, 8, rowNumber, 8].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;

                    sheet.Range[rowNumber, 1, rowNumber, 8].CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom;
                    sheet.Range[rowNumber, 1, rowNumber, 8].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    sheet.Range[rowNumber, 1, rowNumber, 8].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    sheet.SetRowHeight(rowNumber, 35);

                    averageStartRow = rowNumber + 1;
                    foreach (var memberSummary in milestoneSummary.OrderBy(ms => ms.memberName))
                    {
                        rowNumber++;
                        sheet.Range[rowNumber, 1].Text = memberSummary.memberName;
                        sheet.Range[rowNumber, 1].BorderAround();
                        sheet.Range[rowNumber, 2].Number = memberSummary.percentComplete / 100.00;
                        sheet.Range[rowNumber, 2].NumberFormat = "0%";
                        sheet.Range[rowNumber, 2].BorderAround();
                        sheet.Range[rowNumber, 2].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                        switch (milestoneSummary.Key)
                        {
                            case 1:
                                SetMilestoneCell(sheet.Range[rowNumber, 3], 1, participateAssistLead.participate, memberSummary.milestone1ParticipateCommunity);
                                SetMilestoneCell(sheet.Range[rowNumber, 4], 1, participateAssistLead.participate, memberSummary.milestone1ParticipateOutdoors);
                                SetMilestoneCell(sheet.Range[rowNumber, 5], 1, participateAssistLead.participate, memberSummary.milestone1ParticipateCreative);
                                SetMilestoneCell(sheet.Range[rowNumber, 6], 1, participateAssistLead.participate, memberSummary.milestone1ParticipatePersonalGrowth);
                                SetMilestoneCell(sheet.Range[rowNumber, 7], 1, participateAssistLead.lead, memberSummary.milestone1Lead);
                                SetMilestoneCell(sheet.Range[rowNumber, 8], 1, participateAssistLead.lead, memberSummary.milestone1Lead, memberSummary.milestone1LeadArea);
                                break;
                            case 2:
                                SetMilestoneCell(sheet.Range[rowNumber, 3], 2, participateAssistLead.participate, memberSummary.milestone2ParticipateCommunity);
                                SetMilestoneCell(sheet.Range[rowNumber, 4], 2, participateAssistLead.participate, memberSummary.milestone2ParticipateOutdoors);
                                SetMilestoneCell(sheet.Range[rowNumber, 5], 2, participateAssistLead.participate, memberSummary.milestone2ParticipateCreative);
                                SetMilestoneCell(sheet.Range[rowNumber, 6], 2, participateAssistLead.participate, memberSummary.milestone2ParticipatePersonalGrowth);
                                SetMilestoneCell(sheet.Range[rowNumber, 7], 2, participateAssistLead.lead, memberSummary.milestone2Lead);
                                SetMilestoneCell(sheet.Range[rowNumber, 8], 2, participateAssistLead.lead, memberSummary.milestone2Lead, memberSummary.milestone2LeadArea);
                                break;
                            // Key 4 ("Milestone 3 - Completed") reuses Milestone 3's own fields and
                            // level-3 colouring - it's the same milestone, just split into its own section.
                            case 3:
                            case 4:
                                SetMilestoneCell(sheet.Range[rowNumber, 3], 3, participateAssistLead.participate, memberSummary.milestone3ParticipateCommunity);
                                SetMilestoneCell(sheet.Range[rowNumber, 4], 3, participateAssistLead.participate, memberSummary.milestone3ParticipateOutdoors);
                                SetMilestoneCell(sheet.Range[rowNumber, 5], 3, participateAssistLead.participate, memberSummary.milestone3ParticipateCreative);
                                SetMilestoneCell(sheet.Range[rowNumber, 6], 3, participateAssistLead.participate, memberSummary.milestone3ParticipatePersonalGrowth);
                                SetMilestoneCell(sheet.Range[rowNumber, 7], 3, participateAssistLead.lead, memberSummary.milestone3Lead);
                                SetMilestoneCell(sheet.Range[rowNumber, 8], 3, participateAssistLead.lead, memberSummary.milestone3Lead, memberSummary.milestone3LeadArea);
                                break;
                        }
                        sheet.Range[rowNumber, 3].BorderAround();
                        sheet.Range[rowNumber, 4].BorderAround();
                        sheet.Range[rowNumber, 5].BorderAround();
                        sheet.Range[rowNumber, 6].BorderAround();
                        sheet.Range[rowNumber, 7].BorderAround();
                        sheet.Range[rowNumber, 8].BorderAround();
                        sheet.Range[rowNumber, 2, rowNumber, 8].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    }

                    // Add average row
                    averageEndRow = rowNumber;
                    rowNumber++;
                    sheet.Range[rowNumber, 2].Text = "Average";
                    sheet.Range[rowNumber, 2].BorderAround();
                    for (int i = 3; i <= 8; i++)
                    {
                        if (!(i == 8)) // Don't total area text cols
                        {
                            var avgRange = sheet.Range[averageStartRow, i, averageEndRow, i].AddressLocal;
                            sheet.Range[rowNumber, i].Formula = $"=AVERAGE({avgRange})";
                            sheet.Range[rowNumber, i].NumberFormat = "0.0";
                        }
                        sheet.Range[rowNumber, i].BorderAround();
                        sheet.Range[rowNumber, i].CellStyle.Font.Bold = true;
                        sheet.Range[rowNumber, i].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    }
                    sheet.Range[rowNumber, 2, rowNumber, 8].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;

                    rowNumber++;
                    rowNumber++;
                }

                sheet.Range[1, 1, rowNumber, 1].AutofitColumns();
                sheet.SetColumnWidth(2, 10);
                sheet.SetColumnWidth(3, 10);
                sheet.SetColumnWidth(4, 10);
                sheet.SetColumnWidth(5, 10);
                sheet.SetColumnWidth(6, 10);
                sheet.SetColumnWidth(7, 10);
                sheet.SetColumnWidth(8, 15);

                sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
                sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
                sheet.PageSetup.BottomMargin = 0.25;
                sheet.PageSetup.TopMargin = 0.25;
                sheet.PageSetup.LeftMargin = 0.25;
                sheet.PageSetup.RightMargin = 0.25;
                sheet.PageSetup.HeaderMargin = 0;
                sheet.PageSetup.FooterMargin = 0;
                sheet.PageSetup.IsFitToPage = true;

                worksheetIndex++;
            }

            return workbook;
        }

        public IWorkbook GenerateLogbookWorkbook(List<MemberLogbookReportViewModel> logbookEntries, string groupName, string section, string unitName, bool forPdfOutput)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 8);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"Logbook entries as at {DateTime.Now.ToShortDateString()}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            foreach (var logbookGroupEntry in logbookEntries.GroupBy(lb => lb.MemberName).OrderBy(lb => lb.Key))
            {
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = logbookGroupEntry.Key;
                sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
                sheet.Range[rowNumber, 3].Text = $"{logbookGroupEntry.FirstOrDefault()?.TotalKilometersHiked ?? 0} KMs";
                sheet.Range[rowNumber, 4].Text = $"{logbookGroupEntry.FirstOrDefault()?.TotalNightsCamped ?? 0} Nights";
                sheet.Range[rowNumber, 1, rowNumber, 4].CellStyle.Font.Bold = true;

                sheet.SetRowHeight(rowNumber, 20);

                rowNumber++;
                sheet.Range[rowNumber, 2].Text = "Activity Name";
                sheet.Range[rowNumber, 3].Text = "Activity Area";
                sheet.Range[rowNumber, 4].Text = "Activity Date";
                sheet.Range[rowNumber, 5].Text = "Member Role";
                sheet.Range[rowNumber, 6].Text = "Kilometers Hiked";
                sheet.Range[rowNumber, 7].Text = "Nights Camped";
                sheet.Range[rowNumber, 8].Text = "Weather";
                sheet.Range[rowNumber, 9].Text = "Location";
                sheet.Range[rowNumber, 10].Text = "Grade";
                sheet.Range[rowNumber, 11].Text = "Other Participants";
                sheet.Range[rowNumber, 12].Text = "Purpose";
                sheet.Range[rowNumber, 13].Text = "Verifier Name";
                sheet.Range[rowNumber, 14].Text = "Verifier Contact";
                sheet.Range[rowNumber, 2, rowNumber, 14].CellStyle.Font.Italic = true;

                foreach (var logbookEntry in logbookGroupEntry)
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 2].Text = logbookEntry.ActivityName;
                    sheet.Range[rowNumber, 3].Text = logbookEntry.ActivityArea;
                    sheet.Range[rowNumber, 4].DateTime = logbookEntry.ActivityDate;
                    sheet.Range[rowNumber, 5].Text = logbookEntry.MemberRole;
                    sheet.Range[rowNumber, 6].Text = $"{logbookEntry.KilometersHiked} KMs";
                    sheet.Range[rowNumber, 7].Text = $"{logbookEntry.NightsCamped} Nights";
                    sheet.Range[rowNumber, 8].Text = logbookEntry.Weather;
                    sheet.Range[rowNumber, 9].Text = logbookEntry.Location;
                    sheet.Range[rowNumber, 10].Text = logbookEntry.Grade;
                    sheet.Range[rowNumber, 11].Text = logbookEntry.OtherParticipants;
                    sheet.Range[rowNumber, 12].Text = logbookEntry.Purpose;
                    sheet.Range[rowNumber, 13].Text = logbookEntry.VerifierName;
                    sheet.Range[rowNumber, 14].Text = logbookEntry.VerifierContact;
                }
            }

            sheet.SetColumnWidth(1, 10);
            sheet.SetColumnWidth(2, 40);
            sheet.Range[1, 2, rowNumber, 2].CellStyle.WrapText = true;
            sheet.SetColumnWidth(3, 15);
            sheet.SetColumnWidth(4, 15);
            sheet.SetColumnWidth(5, 15);
            sheet.SetColumnWidth(6, 15);
            sheet.SetColumnWidth(7, 15);
            sheet.SetColumnWidth(8, 15);
            sheet.SetColumnWidth(9, 20);
            sheet.Range[1, 9, rowNumber, 9].CellStyle.WrapText = true;
            sheet.SetColumnWidth(10, 10);
            sheet.SetColumnWidth(11, 40);
            sheet.Range[1, 11, rowNumber, 11].CellStyle.WrapText = true;
            sheet.SetColumnWidth(12, 10);
            sheet.SetColumnWidth(13, 15);
            sheet.SetColumnWidth(14, 15);
            sheet.Range[1, 14, rowNumber, 14].AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
            sheet.PageSetup.BottomMargin = 0.5;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0.25;
            sheet.PageSetup.RightFooter = "&P of &N";

            return workbook;

        }

        public IWorkbook GenerateWallchartWorkbook(List<WallchartItemModel> wallchartEntries, string groupName, string section, string unitName, bool forPdfOutput, bool breakByPatrol = false)
        {
            var worksheetAnswersGroupedByTemplate = wallchartEntries.GroupBy(wa => breakByPatrol ? wa.MemberPatrol : "");
            var workbook = CreateWorkbookWithSheets(worksheetAnswersGroupedByTemplate.Count());
            var worksheetIndex = 0;

            foreach (var templatAnswerGroup in worksheetAnswersGroupedByTemplate.OrderBy(a => a.Key))
            {
                IWorksheet sheet = AddLogoToSheet(workbook, worksheetIndex, groupName, section, 42);
                sheet.Name = string.IsNullOrEmpty(templatAnswerGroup.Key) ? "Unit" : templatAnswerGroup.Key;
                int rowNumber = 1;
                int columnNumber = 1;

                IStyle headingStyle = workbook.Styles["headingStyle"];

                // Add Unit name
                rowNumber++;
                var unit = sheet.Range[rowNumber, 2];
                unit.Text = unitName;
                unit.CellStyle = headingStyle;
                sheet.Range[rowNumber, 2, rowNumber, 42].Merge();
                sheet.SetRowHeight(rowNumber, 25);

                // Add Title
                rowNumber++;
                var title = sheet.Range[rowNumber, 2];
                title.Text = $"{(string.IsNullOrEmpty(templatAnswerGroup.Key) ? "" : templatAnswerGroup.Key) + " "}Group Life Wallchart as at {DateTime.Now.ToShortDateString()}";
                title.CellStyle = headingStyle;
                sheet.Range[rowNumber, 2, rowNumber, 42].Merge();
                sheet.SetRowHeight(rowNumber, 25);

                // Add Heading 1
                rowNumber++;
                sheet.Range[rowNumber, 2, rowNumber, 4].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 2, rowNumber, 4].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 5].Text = "MILESTONE 1";
                sheet.Range[rowNumber, 5, rowNumber, 9].Merge();
                sheet.Range[rowNumber, 5, rowNumber, 9].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 5, rowNumber, 9].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 10].Text = "MILESTONE 2";
                sheet.Range[rowNumber, 10, rowNumber, 14].Merge();
                sheet.Range[rowNumber, 10, rowNumber, 14].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 10, rowNumber, 14].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 15].Text = "MILESTONE 3";
                sheet.Range[rowNumber, 15, rowNumber, 19].Merge();
                sheet.Range[rowNumber, 15, rowNumber, 19].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 15, rowNumber, 19].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 20].Text = "OUTDOOR ADVENTURE SKILLS";
                sheet.Range[rowNumber, 20, rowNumber, 31].Merge();
                sheet.Range[rowNumber, 20, rowNumber, 31].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 20, rowNumber, 31].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 32].Text = "SPECIAL INTEREST AREAS";
                sheet.Range[rowNumber, 32, rowNumber, 37].Merge();
                sheet.Range[rowNumber, 32, rowNumber, 37].CellStyle.WrapText = true;
                sheet.Range[rowNumber, 32, rowNumber, 37].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 32, rowNumber, 37].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 38, rowNumber, 41].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 38, rowNumber, 41].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.Font.Bold = true;
                sheet.SetRowHeight(rowNumber, 30);

                // Add Heading 2
                rowNumber++;
                sheet.Range[rowNumber, 2, rowNumber, 19].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 2, rowNumber, 19].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 20].Text = "CORE";
                sheet.Range[rowNumber, 20, rowNumber, 22].Merge();
                sheet.Range[rowNumber, 20, rowNumber, 22].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 20, rowNumber, 22].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 23].Text = "LAND";
                sheet.Range[rowNumber, 23, rowNumber, 25].Merge();
                sheet.Range[rowNumber, 23, rowNumber, 25].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 23, rowNumber, 25].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 26].Text = "WATER";
                sheet.Range[rowNumber, 26, rowNumber, 28].Merge();
                sheet.Range[rowNumber, 26, rowNumber, 28].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 26, rowNumber, 28].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 29, rowNumber, 41].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 29, rowNumber, 41].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.Font.Bold = true;

                // Add Heading 3
                rowNumber++;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom;
                sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.Rotation = 90;
                sheet.Range[rowNumber, 1, rowNumber, 41].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1].Text = " Name";
                sheet.Range[rowNumber, 2].Text = " Intro To Scouting";
                sheet.Range[rowNumber, 3].Text = " Intro To Section";
                sheet.Range[rowNumber, 4].Text = " Years of Service";
                sheet.Range[rowNumber, 5].Text = " Community (6)";
                sheet.Range[rowNumber, 6].Text = " Outdoors (6)";
                sheet.Range[rowNumber, 7].Text = " Creative (6)";
                sheet.Range[rowNumber, 8].Text = " Personal Growth (6)";
                sheet.Range[rowNumber, 9].Text = " Lead (2)";
                sheet.Range[rowNumber, 10].Text = " Community (5)";
                sheet.Range[rowNumber, 11].Text = " Outdoors (5)";
                sheet.Range[rowNumber, 12].Text = " Creative (5)";
                sheet.Range[rowNumber, 13].Text = " Personal Growth (5)";
                sheet.Range[rowNumber, 14].Text = " Lead (4)";
                sheet.Range[rowNumber, 15].Text = " Community (4)";
                sheet.Range[rowNumber, 16].Text = " Outdoors (4)";
                sheet.Range[rowNumber, 17].Text = " Creative (4)";
                sheet.Range[rowNumber, 18].Text = " Personal Growth (4)";
                sheet.Range[rowNumber, 19].Text = " Lead (6)";
                sheet.Range[rowNumber, 20].Text = " Bushcraft";
                sheet.Range[rowNumber, 21].Text = " Bushwalking";
                sheet.Range[rowNumber, 22].Text = " Camping";
                sheet.Range[rowNumber, 23].Text = " Alpine";
                sheet.Range[rowNumber, 24].Text = " Cycling";
                sheet.Range[rowNumber, 25].Text = " Vertical";
                sheet.Range[rowNumber, 26].Text = " Aquatics";
                sheet.Range[rowNumber, 27].Text = " Boating";
                sheet.Range[rowNumber, 28].Text = " Paddling";
                sheet.Range[rowNumber, 29].Text = " Total Progressions";
                sheet.Range[rowNumber, 30].Text = " Total Nights Camped";
                sheet.Range[rowNumber, 31].Text = " Total KMs Hiked";
                sheet.Range[rowNumber, 32].Text = " Adventure & Sport";
                sheet.Range[rowNumber, 33].Text = " Arts & Literature";
                sheet.Range[rowNumber, 34].Text = " Environment";
                sheet.Range[rowNumber, 35].Text = " STEM & Innovation";
                sheet.Range[rowNumber, 36].Text = " Growth & Development";
                sheet.Range[rowNumber, 37].Text = " Creating a Better World";
                sheet.Range[rowNumber, 38].Text = " Leadership Course";
                sheet.Range[rowNumber, 39].Text = " Adventurous Journey";
                sheet.Range[rowNumber, 40].Text = " Personal Reflection";
                sheet.Range[rowNumber, 41].Text = " PEAK AWARD";
                sheet.SetRowHeight(rowNumber, 120);

                foreach (var wallchartEntry in templatAnswerGroup.OrderBy(x => x.MemberName))
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = wallchartEntry.MemberName;
                    SetWallchartCell(sheet.Range[rowNumber, 2], 0, wallchartGroups.intro, wallchartEntry.IntroToScouting);
                    SetWallchartCell(sheet.Range[rowNumber, 3], 0, wallchartGroups.intro, wallchartEntry.IntroToSection);
                    sheet.Range[rowNumber, 4].Number = wallchartEntry.YearsOfService ?? 0;
                    if (wallchartEntry.Milestone1Presented.HasValue)
                    {
                        sheet.Range[rowNumber, 5].DateTime = wallchartEntry.Milestone1Presented.Value;
                        sheet.Range[rowNumber, 5, rowNumber, 9].Merge();
                        sheet.Range[rowNumber, 5, rowNumber, 9].CellStyle.Color = MilestoneColours1[1];
                    }
                    else if (wallchartEntry.Milestone1Awarded.HasValue)
                    {
                        sheet.Range[rowNumber, 5].Text = $"{wallchartEntry.Milestone1Awarded.Value.ToShortDateString()}*";
                        sheet.Range[rowNumber, 5, rowNumber, 9].Merge();
                        sheet.Range[rowNumber, 5, rowNumber, 9].CellStyle.Color = MilestoneColours1[1];
                    }
                    else
                    {
                        SetWallchartCell(sheet.Range[rowNumber, 5], 1, wallchartGroups.participate, wallchartEntry.Milestone1Community);
                        SetWallchartCell(sheet.Range[rowNumber, 6], 1, wallchartGroups.participate, wallchartEntry.Milestone1Outdoors);
                        SetWallchartCell(sheet.Range[rowNumber, 7], 1, wallchartGroups.participate, wallchartEntry.Milestone1Creative);
                        SetWallchartCell(sheet.Range[rowNumber, 8], 1, wallchartGroups.participate, wallchartEntry.Milestone1PersonalGrowth);
                        SetWallchartCell(sheet.Range[rowNumber, 9], 1, wallchartGroups.lead, wallchartEntry.Milestone1Lead);
                    }

                    if (wallchartEntry.Milestone2Presented.HasValue)
                    {
                        sheet.Range[rowNumber, 10].DateTime = wallchartEntry.Milestone2Presented.Value;
                        sheet.Range[rowNumber, 10, rowNumber, 14].Merge();
                        sheet.Range[rowNumber, 10, rowNumber, 14].CellStyle.Color = MilestoneColours1[1];
                    }
                    else if (wallchartEntry.Milestone2Awarded.HasValue)
                    {
                        sheet.Range[rowNumber, 10].Text = $"{wallchartEntry.Milestone2Awarded.Value.ToShortDateString()}*";
                        sheet.Range[rowNumber, 10, rowNumber, 14].Merge();
                        sheet.Range[rowNumber, 10, rowNumber, 14].CellStyle.Color = MilestoneColours1[1];
                    }
                    else
                    {
                        SetWallchartCell(sheet.Range[rowNumber, 10], 2, wallchartGroups.participate, wallchartEntry.Milestone2Community);
                        SetWallchartCell(sheet.Range[rowNumber, 11], 2, wallchartGroups.participate, wallchartEntry.Milestone2Outdoors);
                        SetWallchartCell(sheet.Range[rowNumber, 12], 2, wallchartGroups.participate, wallchartEntry.Milestone2Creative);
                        SetWallchartCell(sheet.Range[rowNumber, 13], 2, wallchartGroups.participate, wallchartEntry.Milestone2PersonalGrowth);
                        SetWallchartCell(sheet.Range[rowNumber, 14], 2, wallchartGroups.lead, wallchartEntry.Milestone2Lead);
                    }

                    if (wallchartEntry.Milestone3Presented.HasValue)
                    {
                        sheet.Range[rowNumber, 15].DateTime = wallchartEntry.Milestone3Presented.Value;
                        sheet.Range[rowNumber, 15, rowNumber, 19].Merge();
                        sheet.Range[rowNumber, 15, rowNumber, 19].CellStyle.Color = MilestoneColours1[1];
                    }
                    else if (wallchartEntry.Milestone3Awarded.HasValue)
                    {
                        sheet.Range[rowNumber, 15].Text = $"{wallchartEntry.Milestone3Awarded.Value.ToShortDateString()}*";
                        sheet.Range[rowNumber, 15, rowNumber, 19].Merge();
                        sheet.Range[rowNumber, 15, rowNumber, 19].CellStyle.Color = MilestoneColours1[1];
                    }
                    else
                    {
                        SetWallchartCell(sheet.Range[rowNumber, 15], 3, wallchartGroups.participate, wallchartEntry.Milestone3Community);
                        SetWallchartCell(sheet.Range[rowNumber, 16], 3, wallchartGroups.participate, wallchartEntry.Milestone3Outdoors);
                        SetWallchartCell(sheet.Range[rowNumber, 17], 3, wallchartGroups.participate, wallchartEntry.Milestone3Creative);
                        SetWallchartCell(sheet.Range[rowNumber, 18], 3, wallchartGroups.participate, wallchartEntry.Milestone3PersonalGrowth);
                        SetWallchartCell(sheet.Range[rowNumber, 19], 3, wallchartGroups.lead, wallchartEntry.Milestone3Lead);
                    }
                    SetWallchartCell(sheet.Range[rowNumber, 20], 0, wallchartGroups.oasCore, wallchartEntry.OASBushcraftStage);
                    SetWallchartCell(sheet.Range[rowNumber, 21], 0, wallchartGroups.oasCore, wallchartEntry.OASBushwalkingStage);
                    SetWallchartCell(sheet.Range[rowNumber, 22], 0, wallchartGroups.oasCore, wallchartEntry.OASCampingStage);
                    SetWallchartCell(sheet.Range[rowNumber, 23], 0, wallchartGroups.oasLand, wallchartEntry.OASAlpineStage);
                    SetWallchartCell(sheet.Range[rowNumber, 24], 0, wallchartGroups.oasLand, wallchartEntry.OASCyclingStage);
                    SetWallchartCell(sheet.Range[rowNumber, 25], 0, wallchartGroups.oasLand, wallchartEntry.OASVerticalStage);
                    SetWallchartCell(sheet.Range[rowNumber, 26], 0, wallchartGroups.oasWater, wallchartEntry.OASAquaticsStage);
                    SetWallchartCell(sheet.Range[rowNumber, 27], 0, wallchartGroups.oasWater, wallchartEntry.OASBoatingStage);
                    SetWallchartCell(sheet.Range[rowNumber, 28], 0, wallchartGroups.oasWater, wallchartEntry.OASPaddlingStage);
                    SetWallchartCell(sheet.Range[rowNumber, 29], 0, wallchartGroups.oasProgression, wallchartEntry.OASStageProgressions);
                    SetWallchartCell(sheet.Range[rowNumber, 30], 0, wallchartGroups.oasProgression, wallchartEntry.NightsCamped);
                    SetWallchartCell(sheet.Range[rowNumber, 31], 0, wallchartGroups.oasProgression, wallchartEntry.KMsHiked);
                    SetWallchartCell(sheet.Range[rowNumber, 32], 0, wallchartGroups.siaOdd, wallchartEntry.SIAAdventureSport);
                    SetWallchartCell(sheet.Range[rowNumber, 33], 0, wallchartGroups.siaEven, wallchartEntry.SIAArtsLiterature);
                    SetWallchartCell(sheet.Range[rowNumber, 34], 0, wallchartGroups.siaOdd, wallchartEntry.SIAEnvironment);
                    SetWallchartCell(sheet.Range[rowNumber, 35], 0, wallchartGroups.siaEven, wallchartEntry.SIAStemInnovation);
                    SetWallchartCell(sheet.Range[rowNumber, 36], 0, wallchartGroups.siaOdd, wallchartEntry.SIAGrowthDevelopment);
                    SetWallchartCell(sheet.Range[rowNumber, 37], 0, wallchartGroups.siaEven, wallchartEntry.SIACreatingABetterWorld);
                    SetWallchartCell(sheet.Range[rowNumber, 38], 0, wallchartGroups.leadershipCourse, wallchartEntry.LeadershipCourse);
                    SetWallchartCell(sheet.Range[rowNumber, 39], 0, wallchartGroups.adventurousJourney, wallchartEntry.AdventurousJourney);
                    SetWallchartCell(sheet.Range[rowNumber, 40], 0, wallchartGroups.personalReflection, wallchartEntry.PersonalReflection);
                    // Percentage complete until the award itself is granted, then the awarded date.
                    if (wallchartEntry.PeakAwardAwarded.HasValue)
                        SetWallchartCell(sheet.Range[rowNumber, 41], 0, wallchartGroups.peakAward, wallchartEntry.PeakAwardAwarded);
                    else
                        SetWallchartCell(sheet.Range[rowNumber, 41], 0, wallchartGroups.peakAward, wallchartEntry.PeakAward);
                    sheet.Range[rowNumber, 1, rowNumber, 41].BorderAround();
                    sheet.Range[rowNumber, 1, rowNumber, 41].BorderInside();
                    sheet.Range[rowNumber, 2, rowNumber, 41].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                }


                sheet.Range[1, 1, rowNumber, 41].AutofitColumns();

                sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
                sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
                sheet.PageSetup.BottomMargin = 0.25;
                sheet.PageSetup.TopMargin = 0.25;
                sheet.PageSetup.LeftMargin = 0.25;
                sheet.PageSetup.RightMargin = 0.25;
                sheet.PageSetup.HeaderMargin = 0;
                sheet.PageSetup.FooterMargin = 0;
                sheet.PageSetup.IsFitToPage = true;

                worksheetIndex++;
            }
            return workbook;
        }


        public IWorkbook GenerateApprovalsWorkbook(List<ApprovalsListModel> selectedApprovals, string groupName, string section, string unitName, DateTime approvalSearchFromDate, DateTime approvalSearchToDate, bool groupByMember, bool forPdfOutput)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 8);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"Award Approvals between {approvalSearchFromDate.ToShortDateString()} and {approvalSearchToDate.ToShortDateString()}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            var groupedApprovals = new List<IGrouping<string, ApprovalsListModel>>();
            if (groupByMember)
                groupedApprovals = selectedApprovals.GroupBy(a => a.member_display_name).ToList();
            else
                groupedApprovals = selectedApprovals.GroupBy(a => a.achievement_name).ToList();

            foreach (var approvalGroup in groupedApprovals.OrderBy(a => a.Key))
            {
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = approvalGroup.Key;
                sheet.Range[rowNumber, 1, rowNumber, 4].Merge();
                if (rowNumber == 4)
                {
                    sheet.Range[rowNumber, 6].Text = "Awarded";
                    sheet.Range[rowNumber, 6].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                    sheet.Range[rowNumber, 7].Text = "Presented";
                    sheet.Range[rowNumber, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                }
                sheet.Range[rowNumber, 1, rowNumber, 7].CellStyle.Font.Bold = true;

                sheet.SetRowHeight(rowNumber, 20);

                foreach (var approval in approvalGroup.OrderBy(a => a.member_display_name).ThenBy(a => a.achievement_name))
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 2].Text = groupByMember ? approval.achievement_name : approval.member_display_name;
                    sheet.Range[rowNumber, 3].Text = approval.submission_type;
                    sheet.Range[rowNumber, 4].Text = approval.submission_status;
                    sheet.Range[rowNumber, 5].Text = approval.submission_outcome;
                    if (approval.awarded_date.HasValue)
                        sheet.Range[rowNumber, 6].DateTime = approval.awarded_date.Value;
                    if (approval.presented_date.HasValue)
                        sheet.Range[rowNumber, 7].DateTime = approval.presented_date.Value;
                }
            }

            // nothing to fit when there are no approvals
            if (rowNumber >= 4)
                sheet.Range[4, 2, rowNumber, 7].AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;

            return workbook;
        }


        // A section's report for a period, or the group's (each unit's, then totals first). Label and detail columns.
        public IWorkbook GenerateSectionReportWorkbook(SectionReportSet report, string groupName, string section, string unitName)
        {
            const int labelWidth = 28;
            const int detailWidth = 80;
            const string dateFormat = "d MMM yyyy";
            var culture = CultureInfo.InvariantCulture;

            var workbook = CreateWorkbookWithLogo(groupName, section, 2);
            IWorksheet sheet = workbook.Worksheets[0];
            IStyle headingStyle = workbook.Styles["headingStyle"];
            int rowNumber = 1;

            rowNumber++;
            sheet.Range[rowNumber, 2].Text = report.IsGroup ? "Group Report" : "Section Report";
            sheet.Range[rowNumber, 2].CellStyle = headingStyle;
            sheet.SetRowHeight(rowNumber, 40);

            if (!report.IsGroup && report.Sections.Count == 1)
            {
                rowNumber++;
                sheet.Range[rowNumber, 2].Text = report.Sections[0].UnitName;
                sheet.Range[rowNumber, 2].CellStyle = headingStyle;
                sheet.SetRowHeight(rowNumber, 30);
            }

            // the logo hangs down from the top left corner, so leave room for it before the lines beside it begin
            var usedPoints = 30.0 + 40.0 + (!report.IsGroup && report.Sections.Count == 1 ? 30.0 : 0.0);
            var logoPoints = sheet.Pictures.Count > 0 ? sheet.Pictures[0].Height * 0.75 : 0.0;
            if (logoPoints + 6 > usedPoints)
            {
                rowNumber++;
                sheet.SetRowHeight(rowNumber, logoPoints + 6 - usedPoints);
            }

            rowNumber++;
            sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
            sheet.Range[rowNumber, 1].Text = $"Generated {report.GeneratedAt.ToString("d MMM yyyy 'at' h:mm tt", culture)} by {report.GeneratedBy}";
            rowNumber++;
            sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
            sheet.Range[rowNumber, 1].Text = $"For the period {report.From.ToString(dateFormat, culture)} to {report.To.ToString(dateFormat, culture)}";

            void Block(string heading)
            {
                rowNumber += 2;
                sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
                sheet.Range[rowNumber, 1].Text = heading;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.Font.Size = 12;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                sheet.SetRowHeight(rowNumber, 20);
            }

            void Row(string label, IEnumerable<string> lines)
            {
                var text = lines.ToList();
                if (text.Count == 0)
                    text.Add("None");
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = label;
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                sheet.Range[rowNumber, 1].CellStyle.WrapText = true;
                sheet.Range[rowNumber, 2].Text = string.Join("\n", text);
                sheet.Range[rowNumber, 2].CellStyle.WrapText = true;
                sheet.Range[rowNumber, 2].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                var lineCount = text.Sum(l => Math.Max(1, (int)Math.Ceiling(l.Length / (double)(detailWidth - 8))));
                sheet.SetRowHeight(rowNumber, Math.Max(15, lineCount * 15));
            }

            string Names(IEnumerable<string> names)
            {
                var list = names.ToList();
                return list.Count == 0 ? "" : " - " + string.Join(", ", list);
            }
            string Day(DateTime date) => date.ToString(dateFormat, culture);
            string Section(string key) => culture.TextInfo.ToTitleCase(key) + "s";
            IEnumerable<string> Counts(IEnumerable<UnitBadgeCount> counts) => counts.Select(c => $"{c.Name}: {c.Count}");

            void Badges(SectionBadges badges, string unitPrefix)
            {
                Block(unitPrefix == "" ? "Awarded in the period" : $"{unitPrefix} awarded in the period");
                Row("Totals", new[]
                {
                    $"Peak Awards: {badges.PeakAwards}",
                    $"Milestones: {badges.MilestonesTotal}",
                    $"Special Interest Areas: {badges.SiaTotal}",
                    $"Outdoor Adventure Skill progressions: {badges.OasProgressions}",
                    $"Additional Awards: {badges.AdditionalAwardsTotal}"
                });
                Block(unitPrefix == "" ? "Badges used" : $"{unitPrefix} badges used");
                Row("Peak Award", new[] { badges.PeakAwards.ToString() });
                Row("Milestones", new[] { $"Total: {badges.MilestonesTotal}" }.Concat(Counts(badges.Milestones)));
                Row("Outdoor Adventure Skills", new[] { $"Badges: {badges.OasBadgesTotal}" }
                    .Concat(badges.OasStages.Select(o => $"{o.Stream} stage {o.Stage}: {o.Count}")));
                Row("Special Interest Areas", new[] { $"Total: {badges.SiaTotal}" }.Concat(Counts(badges.SiaCategories)));
                Row("Other badges", new[] { $"Total: {badges.AdditionalAwardsTotal}" }.Concat(Counts(badges.AdditionalAwards)));
                Row("For new members", Counts(badges.InvestitureItems));
            }

            if (report.IsGroup)
            {
                Block("Group totals");
                Row("Youth", new[]
                {
                    $"{report.TotalYouth} current youth {(report.TotalYouth == 1 ? "member" : "members")} (as at {Day(report.GeneratedAt)})",
                    $"{report.TotalInvested} invested, {report.TotalInvestedInPeriod} invested in this period"
                });
                Row("New members", new[]
                {
                    $"From a younger section: {report.TotalNewFromYounger}",
                    $"Off the street: {report.TotalNewOffTheStreet}",
                    $"From another group: {report.TotalNewFromAnotherGroup}"
                });
                Row("Members leaving", new[]
                {
                    $"Moving up: {report.TotalMovedUp}",
                    $"Moved groups: {report.TotalMovedGroup}",
                    $"Left Scouting: {report.TotalLeftScouting}"
                });
                Row("Youth helpers", new[] { $"{report.TotalYouthHelpers}" });
                Row("Leaders", new[] { $"{report.TotalLeaders} current {(report.TotalLeaders == 1 ? "leader" : "leaders")}" });
                Row("Awarded in the period", new[]
                {
                    $"Peak Awards: {report.TotalPeakAwards}",
                    $"Milestones: {report.TotalMilestones}",
                    $"Special Interest Areas: {report.TotalSia}",
                    $"Outdoor Adventure Skill progressions: {report.TotalOasProgressions}",
                    $"Additional Awards: {report.TotalAdditionalAwards}"
                });
                Row("Badges needed", new[]
                {
                    $"Outdoor Adventure Skill badges: {report.TotalOasBadges}"
                });
                Row("For new members", Counts(report.TotalInvestitureItems));
            }

            foreach (var unit in report.Sections)
            {
                Block(report.IsGroup ? unit.UnitName : "Members");
                Row("Youth", new[]
                {
                    $"{unit.CurrentYouth} current youth {(unit.CurrentYouth == 1 ? "member" : "members")} (as at {Day(report.GeneratedAt)})",
                    $"{unit.InvestedYouth} invested, {unit.InvestedInPeriod} invested in this period"
                });
                if (unit.Ages.Counted > 0)
                {
                    Row("Ages", new[]
                    {
                        $"Average {AgeProfile.Format(unit.Ages.AverageYears)} (youngest {AgeProfile.Format(unit.Ages.YoungestYears)}, oldest {AgeProfile.Format(unit.Ages.OldestYears)}) as at {Day(report.GeneratedAt)}",
                        $"Spread: {unit.Ages.Spread}",
                        "By age: " + string.Join(", ", unit.Ages.Buckets.Select(b => $"{b.Age}: {b.Count}"))
                    });
                }
                var newMembers = new List<string>();
                if (unit.Section != "joey")
                    newMembers.Add($"From the {Section(SectionReportBuilder.PreviousSectionOf(unit.Section))}: {unit.NewFromYoungerSection}{Names(unit.NewMembers.Where(n => n.Origin == NewMemberOrigin.YoungerSection).Select(n => n.Name))}");
                newMembers.Add($"Off the street: {unit.NewOffTheStreet}{Names(unit.NewMembers.Where(n => n.Origin == NewMemberOrigin.OffTheStreet).Select(n => n.Name))}");
                newMembers.Add($"From another group: {unit.NewFromAnotherGroup}{Names(unit.NewMembers.Where(n => n.Origin == NewMemberOrigin.AnotherGroup).Select(n => n.Name))}");
                Row("New members", newMembers);

                var leaving = new List<string>();
                if (unit.NextSection != "")
                    leaving.Add($"Moving up to the {Section(unit.NextSection)}: {unit.MovedUp}{Names(unit.Departed.Where(d => d.Status == DepartureStatus.MovedUp).Select(d => d.Name))}");
                leaving.Add($"Moved groups: {unit.MovedGroup}{Names(unit.Departed.Where(d => d.Status == DepartureStatus.MovedGroup).Select(d => d.Name))}");
                leaving.Add($"Left Scouting: {unit.LeftScouting}{Names(unit.Departed.Where(d => d.Status == DepartureStatus.LeftScouting).Select(d => d.Name))}");
                if (unit.DepartedUnset > 0)
                    leaving.Add($"Not yet sorted: {unit.DepartedUnset}{Names(unit.Departed.Where(d => d.Status == DepartureStatus.Unset).Select(d => d.Name))}");
                Row("Members leaving", leaving);

                var leaders = new List<string> { $"{unit.CurrentLeaders} current {(unit.CurrentLeaders == 1 ? "leader" : "leaders")}" };
                if (unit.SavedRosterDate == null)
                {
                    leaders.Add("New leaders can't be named: no earlier report's roster has been saved");
                }
                else
                {
                    leaders.Add(unit.NewLeaders.Count == 0 ? "New: none" : "New: " + string.Join(", ", unit.NewLeaders));
                }
                leaders.Add(unit.DepartedLeaders.Count == 0 ? "Departed: none" : "Departed: " + string.Join(", ", unit.DepartedLeaders));
                Row("Youth helpers", new[] { $"{unit.YouthHelpers.Count} (under 18 with a leader role in the unit)" });
                Row("Leaders", leaders);

                Block(report.IsGroup ? $"{unit.UnitName}: activities" : "Activities");
                Row("Past events", unit.PastEvents.Select(e => $"{Day(e.Start)} - {e.Title} ({e.ChallengeArea}) - {e.YouthAttended} youth, {e.LeadersAttended} {(e.LeadersAttended == 1 ? "leader" : "leaders")}"));
                Row("Upcoming events", unit.UpcomingEvents.Select(e => $"{Day(e.Start)} - {e.Title} ({e.ChallengeArea})"));
                Badges(unit.Badges, report.IsGroup ? unit.UnitName + ":" : "");
            }

            Block("About these numbers");
            Row("Notes", new[]
            {
                "Youth are the unit's current youth members. A new member is one invested (Introduction to Section) in the period.",
                "Terrain only lists current members, so people who left are found from the next section's unit, the roster saved by the last report and events they came to in the period.",
                "Badges are those awarded in the period to current youth in this section; a member who jumped several Outdoor Adventure Skill levels at once only needs the highest."
            });

            sheet.SetColumnWidth(1, labelWidth);
            sheet.SetColumnWidth(2, detailWidth);

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            return workbook;
        }

        // Totals for one unit over a date range. Label and detail columns.
        public IWorkbook GenerateUnitReportWorkbook(UnitReportModel report, string groupName, string section, string unitName)
        {
            const int labelWidth = 28;
            const int detailWidth = 70;
            const string dateFormat = "d MMM yyyy";
            var culture = CultureInfo.InvariantCulture;

            var workbook = CreateWorkbookWithLogo(groupName, section, 2);
            IWorksheet sheet = workbook.Worksheets[0];
            IStyle headingStyle = workbook.Styles["headingStyle"];
            int rowNumber = 1;

            rowNumber++;
            sheet.Range[rowNumber, 2].Text = string.IsNullOrEmpty(report.UnitName) ? unitName : report.UnitName;
            sheet.Range[rowNumber, 2].CellStyle = headingStyle;
            sheet.SetRowHeight(rowNumber, 40);

            rowNumber++;
            sheet.Range[rowNumber, 2].Text = "Unit History";
            sheet.Range[rowNumber, 2].CellStyle = headingStyle;
            sheet.SetRowHeight(rowNumber, 30);

            rowNumber++;
            sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
            sheet.Range[rowNumber, 1].Text = report.AllTime
                    ? $"All time (to {report.To.ToString(dateFormat, culture)})"
                    : $"{report.From.ToString(dateFormat, culture)} to {report.To.ToString(dateFormat, culture)}";

            void Block(string heading)
            {
                rowNumber += 2;
                sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
                sheet.Range[rowNumber, 1].Text = heading;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.Font.Size = 12;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                sheet.SetRowHeight(rowNumber, 20);
            }

            void Row(string label, IEnumerable<string> lines)
            {
                var text = lines.ToList();
                if (text.Count == 0)
                    text.Add("None");
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = label;
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                sheet.Range[rowNumber, 1].CellStyle.WrapText = true;
                sheet.Range[rowNumber, 2].Text = string.Join("\n", text);
                sheet.Range[rowNumber, 2].CellStyle.WrapText = true;
                sheet.Range[rowNumber, 2].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                var lineCount = text.Sum(l => Math.Max(1, (int)Math.Ceiling(l.Length / (double)(detailWidth - 8))));
                sheet.SetRowHeight(rowNumber, Math.Max(15, lineCount * 15));
            }

            string Number(double value) => value.ToString("#,0.#", culture);
            IEnumerable<string> Counts(IEnumerable<UnitBadgeCount> counts) => counts.Select(c => $"{c.Name}: {c.Count}");

            var m = report.Members;
            Block("Members");
            var currentMembers = new List<string>
            {
                $"{m.CurrentYouth} youth (as at {report.AsAt.ToString(dateFormat, culture)})",
                $"{m.CurrentLeaders} {(m.CurrentLeaders == 1 ? "leader" : "leaders")}"
            };
            if (m.YouthHelpers > 0)
                currentMembers.Insert(1, $"{m.YouthHelpers} youth {(m.YouthHelpers == 1 ? "helper" : "helpers")}");
            Row("Current members", currentMembers);
            Row("Former members", new[] { $"{m.FormerAttendees} who attended the unit's events" });
            Row("Been part of the unit", new[] { m.Total.ToString() });

            var b = report.Badges;
            Block("Badges earned");
            Row("Peak Award", new[] { b.PeakAwards.ToString() });
            Row("Milestones", new[] { $"Total: {b.MilestonesTotal}" }.Concat(Counts(b.Milestones)));
            Row("Outdoor Adventure Skills", new[] { $"Progressions: {b.OasProgressions}" }.Concat(Counts(b.OasStreams)));
            Row("Special Interest Areas", new[] { $"Total: {b.SiaTotal}" }.Concat(Counts(b.SiaCategories)));
            Row("Additional Awards", new[] { $"Total: {b.AdditionalAwardsTotal}" }.Concat(Counts(b.AdditionalAwards)));

            Block("Nights away");
            Row("Nights run", new[] { report.NightsRun.ToString() });
            Row("Person nights camped", new[] { $"{report.PersonNightsCamped} ({report.YouthPersonNights} youth, {report.LeaderPersonNights} {(report.LeaderPersonNights == 1 ? "leader" : "leaders")})" });

            Block("Hiking");
            Row("Km hiked", new[]
            {
                $"{Number(report.HikeKm)} km, {(report.KmPerPerson ? "added up for everyone on each hike" : "each hike counted once")}",
                $"{report.HikeEvents} hike{(report.HikeEvents == 1 ? "" : "s")}, {report.HikeEventsWithDistance} with a distance recorded"
            });

            Block("Hours");
            Row("Leaders volunteered", new[]
            {
                Number(report.LeaderHours) + " hours",
                "Only the hours leaders were with the youth at events. It doesn't include planning or admin time."
            });
            Row("Youth participated", new[] { Number(report.YouthHours) + " hours" });

            Block("About these numbers");
            Row("Events", new[]
            {
                $"{report.EventsCounted} events, {report.EventsWithoutAttendance} with no attendance recorded",
                "Nights, hours and km come from the events' dates, attendance and logbook data.",
                "Badges are for current youth members, earned in this section."
            });

            sheet.SetColumnWidth(1, labelWidth);
            sheet.SetColumnWidth(2, detailWidth);

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            return workbook;
        }

        // One member's history: a block per section, then a summary of the lot. Label and detail columns.
        public IWorkbook GenerateScoutingHistoryWorkbook(ScoutingHistoryModel history, string groupName, string section, string unitName)
        {
            const int labelWidth = 26;
            const int detailWidth = 80;
            const string dateFormat = "d MMM yyyy";
            var culture = CultureInfo.InvariantCulture;

            var workbook = CreateWorkbookWithLogo(groupName, section, 2);
            IWorksheet sheet = workbook.Worksheets[0];
            IStyle headingStyle = workbook.Styles["headingStyle"];
            int rowNumber = 1;

            rowNumber++;
            sheet.Range[rowNumber, 2].Text = "Scouting History";
            sheet.Range[rowNumber, 2].CellStyle = headingStyle;
            sheet.SetRowHeight(rowNumber, 40);

            rowNumber++;
            sheet.Range[rowNumber, 2].Text = history.MemberName;
            sheet.Range[rowNumber, 2].CellStyle = headingStyle;
            sheet.SetRowHeight(rowNumber, 30);

            rowNumber++;
            if (history.ScoutingStart != null)
            {
                sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
                sheet.Range[rowNumber, 1].Text = $"In Scouting since {history.ScoutingStart.Value.ToString(dateFormat, culture)} ({history.TimeInScouting})";
            }

            void Block(string heading)
            {
                rowNumber += 2;
                sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
                sheet.Range[rowNumber, 1].Text = heading;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.Font.Size = 12;
                sheet.Range[rowNumber, 1, rowNumber, 2].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                sheet.SetRowHeight(rowNumber, 20);
            }

            void Row(string label, IEnumerable<string> lines)
            {
                var text = lines.ToList();
                if (text.Count == 0)
                    text.Add("None");
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = label;
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                sheet.Range[rowNumber, 1].CellStyle.WrapText = true;
                sheet.Range[rowNumber, 2].Text = string.Join("\n", text);
                sheet.Range[rowNumber, 2].CellStyle.WrapText = true;
                sheet.Range[rowNumber, 2].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                // room for every line, including ones that wrap
                var lineCount = text.Sum(l => Math.Max(1, (int)Math.Ceiling(l.Length / (double)(detailWidth - 8))));
                sheet.SetRowHeight(rowNumber, Math.Max(15, lineCount * 15));
            }

            string Plural(int count, string word) => $"{count} {word}{(count == 1 ? "" : "s")}";
            string Date(DateTime? date) => date == null ? "" : date.Value.ToString(dateFormat, culture);

            foreach (var s in history.Sections)
            {
                var title = s.SectionName + (string.IsNullOrEmpty(s.Units) ? "" : " - " + s.Units);
                if (s.From != null)
                    title += $"  ({Date(s.From)} to {Date(s.To)}{(string.IsNullOrEmpty(s.TimeInSection) ? "" : ", " + s.TimeInSection)})";
                Block(title);

                Row("Milestones", s.Milestones.Count == 0
                    ? new string[0]
                    : new[] { $"Final level: {s.FinalMilestone}" }
                        .Concat(s.Milestones.Select(m => $"Milestone {m.Stage} - {m.Status}{(m.Date != null ? " " + Date(m.Date) : "")}")));
                Row("Outdoor Adventure Skills", s.Oas.Count == 0
                    ? new string[0]
                    : new[] { $"{Plural(s.OasProgressions, "progression")} in this section" }
                        .Concat(s.Oas.Select(o => $"{o.Stream} - final stage {o.FinalStage} ({Plural(o.Stages.Count, "stage")} in this section: {string.Join(", ", o.Stages.Select(st => st.Stage))})")));
                var sia = s.Sia.Select(a => $"{a.Area}{(string.IsNullOrEmpty(a.Project) ? "" : " - " + a.Project)}{(a.Date != null ? " " + Date(a.Date) : "")}").ToList();
                if (s.SiaInProgress > 0 && sia.Count > 0)
                    sia.Add($"{s.SiaInProgress} not yet awarded");
                Row("Special Interest Areas", sia);
                Row("Additional Awards", s.AdditionalAwards.Select(a => $"{a.Name}{(a.Date != null ? " " + Date(a.Date) : "")}"));
                Row("Activities", new[]
                {
                    $"{Plural(s.Nights, "night")} camped over {Plural(s.Camps, "camp")}",
                    $"{Plural(s.Hikes, "hike")} ({s.HikeKm.ToString("0.#", culture)} km)"
                });
                var roles = new List<string> { $"Led {s.Led}, assisted {s.Assisted}, participated in {s.Participated} logged activities" };
                if (!string.IsNullOrEmpty(s.CurrentRole))
                    roles.Add($"Current role: {s.CurrentRole}");
                Row("Roles", roles);
            }

            var t = history.Totals;
            Block("Summary");
            Row("Time in Scouting", new[] { history.TimeInScouting });
            Row("Badges", new[]
            {
                Plural(t.Milestones, "milestone"),
                $"{Plural(t.OasProgressions, "Outdoor Adventure Skill progression")}",
                Plural(t.Sia, "Special Interest Area"),
                Plural(t.AdditionalAwards, "additional award")
            });
            Row("Activities", new[]
            {
                $"{Plural(t.Nights, "night")} camped over {Plural(t.Camps, "camp")}",
                $"{Plural(t.Hikes, "hike")} ({t.HikeKm.ToString("0.#", culture)} km)"
            });
            Row("Roles", new[] { $"Led {t.Led}, assisted {t.Assisted}, participated in {t.Participated} logged activities" });

            sheet.SetColumnWidth(1, labelWidth);
            sheet.SetColumnWidth(2, detailWidth);

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            return workbook;
        }

        public IWorkbook GenerateProgressWorkbook(ProgressDetailsPageViewModel progressEntries, string groupName, string section, string unitName)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 7);
            IWorksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Milestones";
            int rowNumber = 1;
            int cellNumber = 1;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 7].Merge();
            sheet.SetRowHeight(rowNumber, 40);

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"Progress Details for {progressEntries.Member.first_name} {progressEntries.Member.last_name}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 7].Merge();
            sheet.SetRowHeight(rowNumber, 30);

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Member Number:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = progressEntries.Member.member_number;

            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Intro to Scouting:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = progressEntries.IntroToScoutingDate;

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Intro to Section:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = progressEntries.IntroToSectionDate;

            // Milestones
            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Milestones";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;

            foreach (var milestone in progressEntries.Milestones)
            {
                rowNumber++;
                rowNumber++;
                var statusAndDate = milestone.Awarded == DateTime.MinValue ? milestone.Status : milestone.Status + " " + milestone.Awarded.ToString("dd/MM/yy");
                sheet.Range[rowNumber, 1].Text = $"Milestone {milestone.Milestone}: {statusAndDate}";
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1, rowNumber, 2].Merge();
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = "Community:";
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                cellNumber = 1;
                foreach (var milestoneEvent in milestone.ParticipateLogs.Where(e => e.ChallengeArea == "community"))
                {
                    cellNumber++;
                    sheet.Range[rowNumber, cellNumber].Text = milestoneEvent.EventName;
                    sheet.Range[rowNumber + 1, cellNumber].DateTime = milestoneEvent.EventDate;
                    sheet.Range[rowNumber + 1, cellNumber].HorizontalAlignment = ExcelHAlign.HAlignLeft;
                }
                rowNumber++;
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = "Outdoors:";
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                cellNumber = 1;
                foreach (var milestoneEvent in milestone.ParticipateLogs.Where(e => e.ChallengeArea == "outdoors"))
                {
                    cellNumber++;
                    sheet.Range[rowNumber, cellNumber].Text = milestoneEvent.EventName;
                    sheet.Range[rowNumber + 1, cellNumber].DateTime = milestoneEvent.EventDate;
                    sheet.Range[rowNumber + 1, cellNumber].HorizontalAlignment = ExcelHAlign.HAlignLeft;
                }
                rowNumber++;
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = "Creative:";
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                cellNumber = 1;
                foreach (var milestoneEvent in milestone.ParticipateLogs.Where(e => e.ChallengeArea == "creative"))
                {
                    cellNumber++;
                    sheet.Range[rowNumber, cellNumber].Text = milestoneEvent.EventName;
                    sheet.Range[rowNumber + 1, cellNumber].DateTime = milestoneEvent.EventDate;
                    sheet.Range[rowNumber + 1, cellNumber].HorizontalAlignment = ExcelHAlign.HAlignLeft;
                }
                rowNumber++;
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = "Personal Growth:";
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                cellNumber = 1;
                foreach (var milestoneEvent in milestone.ParticipateLogs.Where(e => e.ChallengeArea == "personal_growth"))
                {
                    cellNumber++;
                    sheet.Range[rowNumber, cellNumber].Text = milestoneEvent.EventName;
                    sheet.Range[rowNumber + 1, cellNumber].DateTime = milestoneEvent.EventDate;
                    sheet.Range[rowNumber + 1, cellNumber].HorizontalAlignment = ExcelHAlign.HAlignLeft;
                }
                rowNumber++;
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = "Lead:";
                sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
                cellNumber = 1;
                foreach (var milestoneEvent in milestone.LeadLogs)
                {
                    cellNumber++;
                    sheet.Range[rowNumber, cellNumber].Text = $"{milestoneEvent.EventName} {milestoneEvent.ChallengeAreaAbbrev}";
                    sheet.Range[rowNumber + 1, cellNumber].DateTime = milestoneEvent.EventDate;
                    sheet.Range[rowNumber + 1, cellNumber].HorizontalAlignment = ExcelHAlign.HAlignLeft;
                }
                rowNumber++;
            }

            sheet.Range[4, 1, rowNumber, 7].AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.IsFitToPage = true;

            sheet = workbook.Worksheets.Create("Sheet2");
            sheet.Name = "OAS - SIA";
            rowNumber = 0;

            // Outdoor Adventure Skills
            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Outdoor Adventure Skills";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;

            rowNumber++;
            rowNumber++;
            for (int i = 1; i < 10; i++)
            {
                sheet.Range[rowNumber, i + 1].Text = $"Stage {i}";
                sheet.Range[rowNumber, i + 1].CellStyle.Font.Bold = true;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Bushcraft:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "bushcraft" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Bushwalking:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "bushwalking" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Camping:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "camping" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Alpine:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "alpine" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Cycling:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "cycling" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Vertical:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "vertical" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Aquatics:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "aquatics" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Boating:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "boating" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Paddling:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            for (int i = 1; i < 10; i++)
            {
                var oas = progressEntries.OASSummaries.Where(o => o.Stream == "paddling" && o.Stage == i).OrderByDescending(o => o.Awarded).FirstOrDefault();
                var oasText = "";
                oasText = setOasText(oas, oasText);
                sheet.Range[rowNumber, i + 1].Text = oasText;
            }

            // Statistics
            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Statistics";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;

            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Progressions:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = progressEntries.Stats.OasProgressions.ToString();
            sheet.Range[rowNumber, 3].Text = "KMs Hiked:";
            sheet.Range[rowNumber, 3].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 4].Text = progressEntries.Stats.KmsHiked.ToString();
            sheet.Range[rowNumber, 5].Text = "Nights Camped:";
            sheet.Range[rowNumber, 5].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 6].Text = progressEntries.Stats.NightsCamped.ToString();
            sheet.Range[rowNumber, 7].Text = "Camped in Section:";
            sheet.Range[rowNumber, 7].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 8].Text = progressEntries.Stats.NightsCampedInSection.ToString();

            // Special Interest Areas
            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Special Interest Areas";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;

            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Area";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = "Project";
            sheet.Range[rowNumber, 2].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
            sheet.Range[rowNumber, 9].Text = "Status";
            sheet.Range[rowNumber, 9].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 10].Text = "Date";
            sheet.Range[rowNumber, 10].CellStyle.Font.Bold = true;
            foreach (var siaSummary in progressEntries.SIASummaries.OrderByDescending(s => s.Status).ThenBy(s => s.Area))
            {
                rowNumber++;
                sheet.Range[rowNumber, 1].Text = siaSummary.Area;
                sheet.Range[rowNumber, 2].Text = siaSummary.Project;
                sheet.Range[rowNumber, 2, rowNumber, 8].Merge();
                sheet.Range[rowNumber, 9].Text = siaSummary.Status;
                sheet.Range[rowNumber, 10].DateTime = siaSummary.Awarded;
                sheet.Range[rowNumber, 10].NumberFormat = "dd/MM/yy";
            }

            // Peak Award
            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = $"Peak Award: {(string.IsNullOrEmpty(progressEntries.PeakAward.Awarded) ? "" : "Awarded " + progressEntries.PeakAward.Awarded)}";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;

            rowNumber++;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Personal Development Course:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = progressEntries.PeakAward.PersonalDevelopmentCourse;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Adventurous Journey:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = progressEntries.PeakAward.AdventurousJourney;
            rowNumber++;
            sheet.Range[rowNumber, 1].Text = "Personal Reflection:";
            sheet.Range[rowNumber, 1].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 2].Text = progressEntries.PeakAward.PersonalReflection;


            sheet.Range[1, 1, rowNumber, 10].AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.FitToPagesTall = 1;
            sheet.PageSetup.FitToPagesWide = 1;

            return workbook;

        }

        private static string setOasText(OASSummary? oas, string oasText)
        {
            if (oas != null)
            {
                oasText = "Started";
                if (oas.Approved > new DateTime(2000, 1, 1))
                {
                    oasText = "Approved";
                }
                if (oas.Awarded > new DateTime(2000, 1, 1))
                {
                    oasText = $"{oas.Awarded.ToString("dd/MM/yy")} {oas.Section}";
                }
            }

            return oasText;
        }

        public IWorkbook GenerateTermProgramWorkbook(List<EventListModel> eventEntries, string groupName, string section, string unitName, bool forPdfOutput)
        {
            var workbook = CreateWorkbookWithLogo(groupName, section, 10);
            IWorksheet sheet = workbook.Worksheets[0];
            int rowNumber = 1;
            int columnNumber = 0;

            IStyle headingStyle = workbook.Styles["headingStyle"];

            // Add Unit name
            rowNumber++;
            var unit = sheet.Range[rowNumber, 2];
            unit.Text = unitName;
            unit.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 10].Merge();
            sheet.SetRowHeight(rowNumber, 25);

            // Add Title
            rowNumber++;
            var title = sheet.Range[rowNumber, 2];
            title.Text = $"Program Term x {DateTime.Now.Year}";
            title.CellStyle = headingStyle;
            sheet.Range[rowNumber, 2, rowNumber, 10].Merge();
            sheet.SetRowHeight(rowNumber, 25);

            //Headings
            rowNumber++;
            rowNumber++;
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Start";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "End";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Date";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Week";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Start and Finish";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Meet at";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Activity";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Organiser";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Lead";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Assist";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Duty Patrol";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Uniform";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Operoo/E1";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            columnNumber++;
            sheet.Range[rowNumber, columnNumber].Text = "Challenge";
            sheet.Range[rowNumber, columnNumber].BorderAround();
            sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.Font.Italic = true;
            sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.Font.Bold = true;
            sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.ColorIndex = ExcelKnownColors.Grey_40_percent;

            //Events
            var currentMonth = 0;
            foreach (var eventEntry in eventEntries)
            {
                if (RunningInUtc)
                {
                    // Adjust time to users local time.
                    // To Do: Check for daylight savings correction
                    eventEntry.StartDateTime = ConvertToClientLocalTime(eventEntry.StartDateTime);
                    eventEntry.EndDateTime = ConvertToClientLocalTime(eventEntry.EndDateTime);
                }

                if (currentMonth != eventEntry.StartDateTime.Month)
                {
                    rowNumber++;
                    currentMonth = eventEntry.StartDateTime.Month;
                    sheet.Range[rowNumber, 1].Text = eventEntry.StartDateTime.ToString("MMMM");
                    sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.Font.Bold = true;
                    sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.ColorIndex = ExcelKnownColors.Grey_25_percent;
                    sheet.Range[rowNumber, 1, rowNumber, columnNumber].BorderAround();
                }
                rowNumber++;
                columnNumber = 1;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.StartDateTimeDisplay;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.EndDateTimeDisplay;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.DateDisplay;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = "";
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.StartFinishDisplay;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.Location;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.EventName;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.Organiser;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.Lead;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.Assist;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = "";
                sheet.Range[rowNumber, columnNumber].BorderAround();
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = "Y";
                sheet.Range[rowNumber, columnNumber].BorderAround();
                sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = "";
                sheet.Range[rowNumber, columnNumber].BorderAround();
                sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                columnNumber++;
                sheet.Range[rowNumber, columnNumber].Text = eventEntry.ChallengeArea;
                sheet.Range[rowNumber, columnNumber].BorderAround();
                sheet.Range[rowNumber, columnNumber].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

                sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.Font.Bold = true;
                if (eventEntry.StartDateTime.Date != eventEntry.EndDateTime.Date
                    || eventEntry.StartDateTime.DayOfWeek == DayOfWeek.Saturday
                    || eventEntry.StartDateTime.DayOfWeek == DayOfWeek.Sunday)
                    sheet.Range[rowNumber, 1, rowNumber, columnNumber].CellStyle.Font.Color = ExcelKnownColors.Red;
            }

            sheet.Range[1, 1, rowNumber, columnNumber].AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            sheet.PageSetup.FitToPagesTall = 1;
            sheet.PageSetup.FitToPagesWide = 1;

            return workbook;

        }

        public IWorkbook GenerateProgramIdeasWorkbook(List<ProgramIdeaItemModel> programIdeas, string groupName, string section, string unitName, bool forPdfOutput)
        {
            var areas = new[] { "Community", "Creative", "Outdoors", "Personal Growth" };

            if (forPdfOutput)
                return GenerateProgramIdeasPdfWorkbook(programIdeas, groupName, section, unitName, areas);

            var workbook = CreateWorkbookWithSheets(areas.Length);

            for (int sheetIndex = 0; sheetIndex < areas.Length; sheetIndex++)
            {
                var area = areas[sheetIndex];
                // Merged through column 6 (not just 3, the last data column) so the heading text has
                // enough width to avoid overlapping the logo - the data table itself stays 3 columns.
                var sheet = AddLogoToSheet(workbook, sheetIndex, groupName, section, 6);
                sheet.Name = area;

                int rowNumber = 2;
                var unitCell = sheet.Range[rowNumber, 2];
                unitCell.Text = unitName;
                unitCell.CellStyle = workbook.Styles["headingStyle"];
                sheet.Range[rowNumber, 2, rowNumber, 6].Merge();
                sheet.SetRowHeight(rowNumber, 25);

                rowNumber++;
                var titleCell = sheet.Range[rowNumber, 2];
                titleCell.Text = $"Program Ideas - {area}";
                titleCell.CellStyle = workbook.Styles["headingStyle"];
                sheet.Range[rowNumber, 2, rowNumber, 6].Merge();
                sheet.SetRowHeight(rowNumber, 25);

                rowNumber += 2;
                var headerRowNumber = rowNumber;
                sheet.Range[rowNumber, 1].Text = "Activity";
                sheet.Range[rowNumber, 2].Text = "Last Run";
                sheet.Range[rowNumber, 3].Text = "Times Run";
                sheet.Range[rowNumber, 1, rowNumber, 3].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1, rowNumber, 3].CellStyle.ColorIndex = ExcelKnownColors.Grey_40_percent;
                sheet.Range[rowNumber, 1].BorderAround();
                sheet.Range[rowNumber, 2].BorderAround();
                sheet.Range[rowNumber, 3].BorderAround();

                // Preserves the order programIdeas arrives in (the caller's current on-page sort),
                // rather than re-sorting here - the report should match what the user was looking at.
                var areaIdeas = programIdeas.Where(i => i.ChallengeArea == area);
                foreach (var idea in areaIdeas)
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = idea.ActivityName;
                    sheet.Range[rowNumber, 1].BorderAround();
                    sheet.Range[rowNumber, 2].Text = idea.LastRunDate.ToString("dd/MM/yyyy");
                    sheet.Range[rowNumber, 2].BorderAround();
                    sheet.Range[rowNumber, 3].Number = idea.TimesRun;
                    sheet.Range[rowNumber, 3].BorderAround();

                    if (idea.Highlight)
                    {
                        sheet.Range[rowNumber, 1, rowNumber, 3].CellStyle.Color = Color.Yellow;
                    }
                }

                if (rowNumber > headerRowNumber)
                {
                    sheet.AutoFilters.FilterRange = sheet.Range[headerRowNumber, 1, rowNumber, 3];
                }

                sheet.Range[1, 1, rowNumber, 3].AutofitColumns();

                sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
                sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
                sheet.PageSetup.BottomMargin = 0.25;
                sheet.PageSetup.TopMargin = 0.25;
                sheet.PageSetup.LeftMargin = 0.25;
                sheet.PageSetup.RightMargin = 0.25;
                sheet.PageSetup.HeaderMargin = 0;
                sheet.PageSetup.FooterMargin = 0;
                sheet.PageSetup.FitToPagesWide = 1;
            }

            workbook.Worksheets[0].Activate();
            return workbook;
        }

        // PDF gets its own single-sheet layout (rather than the xlsx's one-sheet-per-area, which the
        // Excel-to-PDF renderer always starts on a fresh page) so short categories can share a page:
        // a category only forces a page break before the next one if it used more than half a page
        // itself, otherwise the next category continues right below it.
        private IWorkbook GenerateProgramIdeasPdfWorkbook(List<ProgramIdeaItemModel> programIdeas, string groupName, string section, string unitName, string[] areas)
        {
            // Calibrated empirically: a single-category A4 portrait page at these margins/default row
            // height fits ~47-50 rows (header + logo overhead included) before naturally paginating.
            const int RowsPerPage = 48;
            const int HalfPageRows = RowsPerPage / 2;

            var workbook = CreateWorkbookWithSheets(1);
            var sheet = AddLogoToSheet(workbook, 0, groupName, section, 6);
            sheet.Name = "Program Ideas";

            int rowNumber = 2;
            var unitCell = sheet.Range[rowNumber, 2];
            unitCell.Text = unitName;
            unitCell.CellStyle = workbook.Styles["headingStyle"];
            sheet.Range[rowNumber, 2, rowNumber, 6].Merge();
            sheet.SetRowHeight(rowNumber, 25);

            rowNumber++;
            var titleCell = sheet.Range[rowNumber, 2];
            titleCell.Text = "Program Ideas";
            titleCell.CellStyle = workbook.Styles["headingStyle"];
            sheet.Range[rowNumber, 2, rowNumber, 6].Merge();
            sheet.SetRowHeight(rowNumber, 25);

            rowNumber += 2;

            for (int areaIndex = 0; areaIndex < areas.Length; areaIndex++)
            {
                var area = areas[areaIndex];
                var categoryStartRow = rowNumber;

                var areaHeadingCell = sheet.Range[rowNumber, 1];
                areaHeadingCell.Text = area;
                areaHeadingCell.CellStyle.Font.Bold = true;
                areaHeadingCell.CellStyle.Font.Size = 12;
                rowNumber++;

                sheet.Range[rowNumber, 1].Text = "Activity";
                sheet.Range[rowNumber, 2].Text = "Last Run";
                sheet.Range[rowNumber, 3].Text = "Times Run";
                sheet.Range[rowNumber, 1, rowNumber, 3].CellStyle.Font.Bold = true;
                sheet.Range[rowNumber, 1, rowNumber, 3].CellStyle.ColorIndex = ExcelKnownColors.Grey_40_percent;
                sheet.Range[rowNumber, 1].BorderAround();
                sheet.Range[rowNumber, 2].BorderAround();
                sheet.Range[rowNumber, 3].BorderAround();

                var areaIdeas = programIdeas.Where(i => i.ChallengeArea == area);
                foreach (var idea in areaIdeas)
                {
                    rowNumber++;
                    sheet.Range[rowNumber, 1].Text = idea.ActivityName;
                    sheet.Range[rowNumber, 1].BorderAround();
                    sheet.Range[rowNumber, 2].Text = idea.LastRunDate.ToString("dd/MM/yyyy");
                    sheet.Range[rowNumber, 2].BorderAround();
                    sheet.Range[rowNumber, 3].Number = idea.TimesRun;
                    sheet.Range[rowNumber, 3].BorderAround();

                    if (idea.Highlight)
                    {
                        sheet.Range[rowNumber, 1, rowNumber, 3].CellStyle.Color = Color.Yellow;
                    }
                }

                var rowsUsedByCategory = rowNumber - categoryStartRow + 1;
                var isLastCategory = areaIndex == areas.Length - 1;
                rowNumber += 2;

                if (!isLastCategory && rowsUsedByCategory > HalfPageRows)
                {
                    sheet.HPageBreaks.Add(sheet.Range[rowNumber, 1]);
                }
            }

            sheet.Range[1, 1, rowNumber, 6].AutofitColumns();

            sheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
            sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
            sheet.PageSetup.BottomMargin = 0.25;
            sheet.PageSetup.TopMargin = 0.25;
            sheet.PageSetup.LeftMargin = 0.25;
            sheet.PageSetup.RightMargin = 0.25;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 0;
            // FitToPagesWide alone (with FitToPagesTall left at its default) shrinks everything to
            // fit on a single page vertically too, which silently defeats both natural pagination and
            // the manual HPageBreaks above - 0 means "as many pages tall as needed" instead.
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            return workbook;
        }

        private DateTime ConvertToClientLocalTime(DateTime dateTime)
        {
            return dateTime.AddHours(CurrentUtcOffset.Hours).AddMinutes(CurrentUtcOffset.Minutes);
        }

        private string FormatEventDate(DateTime startDateTime, DateTime endDateTime)
        {
            if (startDateTime.Date == endDateTime.Date)
                return startDateTime.ToString("ddd d");
            else
                return $"{startDateTime.ToString("ddd d")} - {endDateTime.ToString("ddd d")}";
        }

        private string FormatEventTime(DateTime startDateTime, DateTime endDateTime)
        {
            if (startDateTime.Date == endDateTime.Date)
                return $"{startDateTime.ToShortTimeString()} - {endDateTime.ToShortTimeString()}";
            else
                return $"{startDateTime.ToString("ddd HH:mm")} - {endDateTime.ToString("ddd HH:mm")}";
        }

        private int UnitMaxAge(string unit)
        {
            switch (unit)
            {
                case "joey":
                    return 8;
                case "cub":
                    return 11;
                case "scout":
                    return 15;
                case "venturer":
                    return 18;
                case "rover":
                default:
                    return 26;
            }
        }

        private ExcelKnownColors GetChallengeAreaColour(string challengeArea)
        {
            switch (challengeArea)
            {
                case "Community":
                    return ExcelKnownColors.Grey_25_percent;
                case "Creative":
                    return ExcelKnownColors.Rose;
                case "Outdoors":
                    return ExcelKnownColors.Sea_green;
                case "Growth":
                    return ExcelKnownColors.Gold;
                default:
                    return ExcelKnownColors.None;
            }
        }

        private Color GetInputTitleColour(string inputTitle)
        {
            switch (inputTitle)
            {
                case "Plan":
                    return Color.LightBlue;
                case "Do":
                    return Color.SkyBlue;
                case "Review":
                    return Color.CornflowerBlue;
                case "Verify":
                    return Color.RoyalBlue;
                default:
                    return Color.White;
            }
        }

        private string GetParticipateHeadingText(int currentLevel)
        {
            switch (currentLevel)
            {
                case 1:
                    return "Participate - 6";
                case 2:
                    return "Participate - 5";
                case 3:
                case 4:
                    return "Participate - 4";
                default:
                    return "Participate";
            }
        }


        private string GetLeadHeadingText(int currentLevel)
        {
            switch (currentLevel)
            {
                case 1:
                    return "Lead - 2";
                case 2:
                    return "Lead - 4";
                case 3:
                case 4:
                    return "Lead - 6";
                default:
                    return "Lead";
            }
        }

        private void SetMilestoneCell(IRange cell, int currentLevel, participateAssistLead pal, int count)
        {
            cell.Number = count;
            cell.CellStyle.Color = GetMilestoneProgressColour(currentLevel, pal, count);
        }

        private void SetMilestoneCell(IRange cell, int currentLevel, participateAssistLead pal, int count, string areas)
        {
            cell.Text = areas;
            cell.CellStyle.Color = GetMilestoneProgressColour(currentLevel, pal, count);
        }

        private Color GetMilestoneProgressColour(int currentLevel, participateAssistLead pal, int count)
        {
            // Skipped milestones have a count of -1
            if (count < 0)
                return Color.White;

            if (currentLevel == 1)
            {
                switch (pal)
                {
                    case participateAssistLead.participate:
                        return MilestoneColours6[Math.Min(count, 6)];
                    case participateAssistLead.lead:
                        return MilestoneColours2[Math.Min(count, 2)];
                    default:
                        return Color.White;
                }
            }

            if (currentLevel == 2)
            {
                switch (pal)
                {
                    case participateAssistLead.participate:
                        return MilestoneColours5[Math.Min(count, 5)];
                    case participateAssistLead.lead:
                        return MilestoneColours4[Math.Min(count, 4)];
                    default:
                        return Color.White;
                }
            }

            if (currentLevel == 3)
            {
                switch (pal)
                {
                    case participateAssistLead.participate:
                        return MilestoneColours4[Math.Min(count, 4)];
                    case participateAssistLead.lead:
                        return MilestoneColours6[Math.Min(count, 6)];
                    default:
                        return Color.White;
                }
            }

            return Color.White;
        }

        private void SetWallchartCell(IRange cell, int currentLevel, wallchartGroups wallchartGroups, int count)
        {
            if (count > 0)
                cell.Number = count;
            SetWallchartCellBackground(cell, currentLevel, wallchartGroups, count);
        }

        private void SetWallchartCell(IRange cell, int currentLevel, wallchartGroups wallchartGroups, DateTime? dateAwarded)
        {
            if (dateAwarded.HasValue)
            {
                cell.DateTime = dateAwarded.Value;
                cell.NumberFormat = "dd/MM/yy";
            }
            SetWallchartCellBackground(cell, currentLevel, wallchartGroups, 0);
        }

        private void SetWallchartCell(IRange cell, int currentLevel, wallchartGroups wallchartGroups, double percent)
        {
            cell.Number = percent;
            cell.NumberFormat = "0%";
            SetWallchartCellBackground(cell, currentLevel, wallchartGroups, 0);
        }

        private void SetWallchartCellBackground(IRange cell, int currentLevel, wallchartGroups wallchartGroup, int count)
        {
            switch (wallchartGroup)
            {
                case wallchartGroups.intro:
                    cell.CellStyle.Color = Color.Gainsboro;
                    break;
                case wallchartGroups.participate:
                    switch (currentLevel)
                    {
                        case 1:
                            if (count == 6)
                                cell.CellStyle.Color = MilestoneColours6[6];
                            break;
                        case 2:
                            if (count == 5)
                                cell.CellStyle.Color = MilestoneColours5[5];
                            break;
                        case 3:
                            if (count >= 4)
                                cell.CellStyle.Color = MilestoneColours4[4];
                            break;
                    }
                    break;
                case wallchartGroups.lead:
                    switch (currentLevel)
                    {
                        case 1:
                            if (count == 2)
                                cell.CellStyle.Color = MilestoneColours2[2];
                            else
                                cell.CellStyle.Color = Color.Gainsboro;
                            break;
                        case 2:
                            if (count == 4)
                                cell.CellStyle.Color = MilestoneColours4[4];
                            else
                                cell.CellStyle.Color = Color.Gainsboro;
                            break;
                        case 3:
                            if (count >= 6)
                                cell.CellStyle.Color = MilestoneColours6[6];
                            else
                                cell.CellStyle.Color = Color.Gainsboro;
                            break;
                    }
                    break;
                case wallchartGroups.oasCore:
                case wallchartGroups.adventurousJourney:
                    cell.CellStyle.Color = Color.DarkSeaGreen;
                    break;
                case wallchartGroups.oasLand:
                case wallchartGroups.personalReflection:
                    cell.CellStyle.Color = Color.LightSalmon;
                    break;
                case wallchartGroups.oasWater:
                case wallchartGroups.leadershipCourse:
                    cell.CellStyle.Color = Color.LightBlue;
                    break;
                case wallchartGroups.oasProgression:
                    cell.CellStyle.Color = Color.Gainsboro;
                    break;
                case wallchartGroups.siaOdd:
                    cell.CellStyle.Color = Color.MistyRose;
                    break;
                case wallchartGroups.siaEven:
                    cell.CellStyle.Color = Color.LightYellow;
                    break;
                case wallchartGroups.peakAward:
                    cell.CellStyle.Color = Color.Thistle;
                    break;
            }
        }

    }
}
