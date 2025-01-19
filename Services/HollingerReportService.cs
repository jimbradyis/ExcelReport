using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using ExcelReport.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelReport.Services
{
    public class HollingerReportService
    {
        private readonly EthicsContext _context;

        public HollingerReportService(EthicsContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Builds one Excel workbook containing:
        ///   1) A "Hollinger box Summary" tab (as before).
        ///   2) Additional tabs for each Congress (descending order),
        ///      matching the design you described: row 3 has
        ///      [Short Inquiry, Long Name, Total Count, HASC Key, ... Label4].
        /// </summary>
        /// <param name="filePath">
        ///   e.g. "C:\\HollingerReports\\HollingerBoxSummery.xlsx"
        /// </param>
        public void BuildAndSaveCompleteWorkbook(string filePath)
        {
            // -------------------------------------------------
            // 1) Gather top-level data for the summary tab
            // -------------------------------------------------
            var totalCongresses = _context.Congresses.Count();
            var totalInquiries = _context.Inquiries.Count();
            var totalArchives = _context.Archives.Count();

            // We also want all the Archive data for the summary grouping:
            var archivesForSummary = _context.Archives
                .Include(a => a.CongressNavigation)
                .Include(a => a.SubcommitteeNavigation)
                .Where(a => a.CongressNavigation != null && a.SubcommitteeNavigation != null)
                .ToList();

            // Group them (descending by CongressNo)
            var summaryGroupData = archivesForSummary
                .GroupBy(a => new
                {
                    a.CongressNavigation.CongressNo,
                    a.CongressNavigation.YearLabel,
                    a.CongressNavigation.Years,
                    a.SubcommitteeNavigation.Subcommittee,
                    a.SubcommitteeNavigation.LongName
                })
                .Select(g => new
                {
                    g.Key.CongressNo,
                    g.Key.YearLabel,
                    g.Key.Years,
                    Subcommittee = g.Key.Subcommittee,
                    LongName = g.Key.LongName,

                    TotalBoxes = g.Count(),
                    FillingCount = g.Count(a => a.Status == "Filling"),
                    AdjustCount = g.Count(a => a.Status == "Adjust"),
                    ClosedNotPrinted = g.Count(a => a.Status == "Closed" && (a.Printed ?? 0) == 0),
                    ClosedPrinted = g.Count(a => a.Status == "Closed" && (a.Printed ?? 0) == 1)
                })
                // Sort by CongressNo DESC, subcommittee ascending
                .OrderByDescending(x => x.CongressNo)
                .ThenBy(x => x.Subcommittee)
                .ToList();

            // Then group that by CongressNo (descending)
            var summaryCongressGroups = summaryGroupData
                .GroupBy(x => new { x.CongressNo, x.YearLabel, x.Years })
                .OrderByDescending(g => g.Key.CongressNo)
                .ToList();

            // -------------------------------------------------
            // 2) Gather all Congresses for the separate tabs
            //    (descending order of CongressNo)
            // -------------------------------------------------
            var allCongresses = _context.Congresses
                .OrderByDescending(c => c.CongressNo)
                .ToList();

            // -------------------------------------------------
            // 3) CREATE one workbook in memory
            // -------------------------------------------------
            using var wb = new XLWorkbook();

            // ================================
            // PART A: The "Hollinger box Summary" tab
            // ================================
            var summaryWs = wb.Worksheets.Add("Hollinger box Summary");

            int row = 1;

            // (A) TITLE ROW
            var titleRange = summaryWs.Range(row, 1, row, 7);
            titleRange.Merge();
            var titleCell = summaryWs.Cell(row, 1);
            titleCell.Value = "Hollinger box Summary";
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 16;
            titleCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            row++;

            // (B) DATE ROW
            var dateRange = summaryWs.Range(row, 1, row, 7);
            dateRange.Merge();
            var dateCell = summaryWs.Cell(row, 1);
            dateCell.Value = $"Date: {DateTime.Now:MM/dd/yyyy}";
            dateCell.Style.Font.Italic = true;
            dateCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

            row++;

            // (C) Summaries
            int summaryStartRow = row;

            summaryWs.Cell(row, 1).Value = "Congresses:";
            summaryWs.Cell(row, 2).Value = totalCongresses;
            row++;

            summaryWs.Cell(row, 1).Value = "Inquiries:";
            summaryWs.Cell(row, 2).Value = totalInquiries;
            row++;

            summaryWs.Cell(row, 1).Value = "Hollinger Boxes:";
            summaryWs.Cell(row, 2).Value = totalArchives;
            row++;

            // Border & shading
            var summaryRange = summaryWs.Range(summaryStartRow, 1, row - 1, 2);
            summaryRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            summaryRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            summaryRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            row++; // blank row

            // (D) Now fill the detail data for each Congress group
            foreach (var cg in summaryCongressGroups)
            {
                // Merge columns 1..7 for a heading row
                var headingRange = summaryWs.Range(row, 1, row, 7);
                headingRange.Merge();

                var headingCell = summaryWs.Cell(row, 1);
                headingCell.Value = $"{cg.Key.YearLabel} ({cg.Key.Years})";
                headingCell.Style.Font.Bold = true;
                headingCell.Style.Fill.BackgroundColor = XLColor.LightBlue;

                row++;

                // Column headers
                summaryWs.Cell(row, 1).Value = "Inquiry";
                summaryWs.Cell(row, 2).Value = "Long Name";
                summaryWs.Cell(row, 3).Value = "Total Boxes";
                summaryWs.Cell(row, 4).Value = "Filling";
                summaryWs.Cell(row, 5).Value = "Adjust";
                summaryWs.Cell(row, 6).Value = "Closed";
                summaryWs.Cell(row, 7).Value = "Printed";

                var hdrRange = summaryWs.Range(row, 1, row, 7);
                hdrRange.Style.Font.Bold = true;
                hdrRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                hdrRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                row++;

                // Subcommittee rows
                foreach (var item in cg)
                {
                    summaryWs.Cell(row, 1).Value = item.Subcommittee;
                    summaryWs.Cell(row, 2).Value = item.LongName;
                    summaryWs.Cell(row, 3).Value = item.TotalBoxes;
                    summaryWs.Cell(row, 4).Value = item.FillingCount;
                    summaryWs.Cell(row, 5).Value = item.AdjustCount;
                    summaryWs.Cell(row, 6).Value = item.ClosedNotPrinted;
                    summaryWs.Cell(row, 7).Value = item.ClosedPrinted;

                    // Right-align numeric columns
                    var numericCells = summaryWs.Range(row, 3, row, 7);
                    numericCells.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                    row++;
                }

                row += 2; // blank rows before next congress group
            }

            // Finally, auto-fit columns on the summary sheet
            summaryWs.Columns().AdjustToContents();

            // ================================
            // PART B: A separate tab for each Congress
            // ================================
            foreach (var c in allCongresses)
            {
                // We'll create a tab name from YearLabel (or fallback)
                // Keep it short if YearLabel is long (sheet name limit ~31 chars).
                // For example: "99th" or "99th(1985-1986)".
                var sheetName = c.YearLabel;
                if (string.IsNullOrEmpty(sheetName))
                    sheetName = $"C{c.CongressNo}";  // fallback
                if (sheetName.Length > 25)
                    sheetName = sheetName.Substring(0, 25); // trim if needed

                var ws = wb.Worksheets.Add(sheetName);

                // Row pointer
                int crow = 1;

                // Merge row1 for a "Congress: {YearLabel} ({Years})" heading
                var congressTitleRange = ws.Range(crow, 1, crow, 14);
                congressTitleRange.Merge();
                var congressTitleCell = ws.Cell(crow, 1);
                congressTitleCell.Value = $"Congress: {c.YearLabel} ({c.Years})";
                congressTitleCell.Style.Font.Bold = true;
                congressTitleCell.Style.Font.FontSize = 16;
                congressTitleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                crow += 2; // blank line

                // HEADERS in row 3 (or crow?)
                // A: "Short Inquiry"
                // B: "Long Name"
                // C: "Total Count"
                // D..N: [HASC Key, Archive No, Hollinger Box Key, Box Label w/o congress,
                //        Status, Doc, Note, Label1, Label2, Label3, Label4]

                // We'll do it at row = 3:
                // But if we've already advanced crow = 3 from above, let's set crow=3 explicitly:
                crow = 3;

                ws.Cell(crow, 1).Value = "Short Inquiry";
                ws.Cell(crow, 2).Value = "Long Name";
                ws.Cell(crow, 3).Value = "Total Count";
                ws.Cell(crow, 4).Value = "HASC Key";
                ws.Cell(crow, 5).Value = "Archive No";
                ws.Cell(crow, 6).Value = "Hollinger Box Key";
                ws.Cell(crow, 7).Value = "Box Label without congress";
                ws.Cell(crow, 8).Value = "Status";
                ws.Cell(crow, 9).Value = "Doc";
                ws.Cell(crow, 10).Value = "Note";
                ws.Cell(crow, 11).Value = "Label1";
                ws.Cell(crow, 12).Value = "Label2";
                ws.Cell(crow, 13).Value = "Label3";
                ws.Cell(crow, 14).Value = "Label4";

                var congressHeader = ws.Range(crow, 1, crow, 14);
                congressHeader.Style.Font.Bold = true;
                congressHeader.Style.Fill.BackgroundColor = XLColor.LightGray;
                congressHeader.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                crow++;

                // For the given congress, load the Archives + Inquiry + Docs
                var cArchives = _context.Archives
                    .Include(a => a.SubcommitteeNavigation)
                    .Include(a => a.Docs)
                    .Where(a => a.Congress == c.CongressNo)
                    .ToList();

                // Group by Inquiry
                var inquiryGroups = cArchives
                    .GroupBy(a => a.Subcommittee)
                    .Select(g => new
                    {
                        Subcommittee = g.Key,
                        LongName = g.First().SubcommitteeNavigation.LongName,
                        TotalCount = g.Count(), // total archives for that inquiry
                        Archives = g.OrderBy(a => a.ArchiveNo).ToList()
                    })
                    .OrderBy(x => x.Subcommittee) // alphabetical
                    .ToList();

                // Fill data
                foreach (var inq in inquiryGroups)
                {
                    bool firstArchiveForInquiry = true;

                    foreach (var arc in inq.Archives)
                    {
                        // A..C
                        ws.Cell(crow, 1).Value = inq.Subcommittee;
                        ws.Cell(crow, 2).Value = inq.LongName;

                        if (firstArchiveForInquiry)
                        {
                            ws.Cell(crow, 3).Value = inq.TotalCount;
                            firstArchiveForInquiry = false;
                        }
                        else
                        {
                            // or blank
                            ws.Cell(crow, 3).Value = "";
                        }

                        // D..N
                        ws.Cell(crow, 4).Value = arc.HascKey;
                        ws.Cell(crow, 5).Value = arc.ArchiveNo;
                        ws.Cell(crow, 6).Value = arc.HollingerBoxKey;
                        ws.Cell(crow, 7).Value = arc.BoxLabelWithoutCongress;

                        // Status cell color logic
                        var statusCell = ws.Cell(crow, 8);

                        if (arc.Status == "Closed" && (arc.Printed ?? 0) == 1)
                        {
                            // "Closed & Printed" in green
                            statusCell.Value = "Closed & Printed";
                            statusCell.Style.Font.FontColor = XLColor.Green;
                        }
                        else if (
                            (arc.Status == "Filling" || arc.Status == "Adjust" || arc.Status == "Closed")
                            && (arc.Printed ?? 0) == 0
                        )
                        {
                            statusCell.Value = arc.Status;
                            statusCell.Style.Font.FontColor = XLColor.Red;
                        }
                        else
                        {
                            statusCell.Value = arc.Status;
                        }

                        // Doc count
                        ws.Cell(crow, 9).Value = arc.Docs.Count;

                        // Note
                        ws.Cell(crow, 10).Value = arc.Note;
                        // Label1..Label4
                        ws.Cell(crow, 11).Value = arc.Label1;
                        ws.Cell(crow, 12).Value = arc.Label2;
                        ws.Cell(crow, 13).Value = arc.Label3;
                        ws.Cell(crow, 14).Value = arc.Label4;

                        crow++;
                    }

                    // maybe a blank row between inquiries
                    crow++;
                }

                // Auto-fit columns for this congress sheet
                ws.Columns().AdjustToContents();
            }

            // -------------------------------------------------
            // 4) SAVE to disk
            // -------------------------------------------------
            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            wb.SaveAs(filePath);
        }
    }
}
