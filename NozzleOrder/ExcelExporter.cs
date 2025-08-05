using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using static NewJeans.NozzleOrderHandler;


namespace NewJeans
{
    public static class ExcelExporter
    {
        public static void Export(List<(string, string, string)> finalResults, string filePath)
        {

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Nozzle 순서검토");

                // 헤더 작성
                worksheet.Cells[1, 1].Value = "Family";
                worksheet.Cells[1, 2].Value = "Type";
                worksheet.Cells[1, 3].Value = "Nozzle 검토결과";
                
                worksheet.Cells[1, 1, 1, 3].AutoFilter = true;

                var headerRange = worksheet.Cells[1, 1, 1, 3];
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(252, 213, 180));
                    headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    var border = headerRange.Style.Border;
                    border.Top.Style = ExcelBorderStyle.Thin;
                    border.Bottom.Style = ExcelBorderStyle.Thin;
                    border.Left.Style = ExcelBorderStyle.Thin;
                    border.Right.Style = ExcelBorderStyle.Thin;
                }

                // 데이터 입력
                int row = 2;
                foreach (var result in finalResults)
                {
                    worksheet.Cells[row, 1].Value = result.Item1;
                    worksheet.Cells[row, 2].Value = result.Item2;
                    worksheet.Cells[row, 3].Value = result.Item3;
                    row++;
                }

                // 열 너비 자동 조정
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // 저장
                var file = new FileInfo(filePath);
                package.SaveAs(file);
            }
        }
    }
}
