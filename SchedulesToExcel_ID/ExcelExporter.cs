using OfficeOpenXml.Style;
using OfficeOpenXml;
using static NewJeans.ExportScheduleHandler;
using System.Collections.Generic;
using System.IO;

public static class ExcelExporter
{
    public static void Export(List<ScheduleDocument> scheduleDocuments, string filePath)
    {
        using (var package = new ExcelPackage())
        {
            foreach (var scheduleDocument in scheduleDocuments)
            {
                foreach (var table in scheduleDocument.Tables)
                {
                    var ws = package.Workbook.Worksheets.Add(table.Name);

                    // 헤더 작성
                    for (int c = 0; c < table.Headers.Count; c++)
                    {
                        ws.Cells[1, c + 1].Value = table.Headers[c];
                    }

                    ws.Cells[1, 1, 1, table.Headers.Count].AutoFilter = true;

                    using (var headerRange = ws.Cells[1, 1, 1, table.Headers.Count])
                    {
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                        headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        var border = headerRange.Style.Border;
                        border.Top.Style = ExcelBorderStyle.Thin;
                        border.Bottom.Style = ExcelBorderStyle.Thin;
                        border.Left.Style = ExcelBorderStyle.Thin;
                        border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    // 본문 작성
                    for (int r = 0; r < table.BodyRows.Count; r++)
                    {
                        for (int c = 0; c < table.BodyRows[r].Count; c++)
                        {
                            ws.Cells[r + 2, c + 1].Value = table.BodyRows[r][c];
                        }
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                }
            }

            package.SaveAs(new FileInfo(filePath));
        }
    }
}
