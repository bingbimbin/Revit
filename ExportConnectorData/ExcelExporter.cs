using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using static NewJeans.ExtractParametersHandler;

namespace NewJeans
{
    public static class ExcelExporter
    {
        public static void Export(List<ConnectorInfo> classifications, string filePath)
        {

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Connector Data");

                // 헤더
                worksheet.Cells[1, 1].Value = "Category";
                worksheet.Cells[1, 2].Value = "Family";
                worksheet.Cells[1, 3].Value = "System Classification";
                worksheet.Cells[1, 4].Value = "Description";
                worksheet.Cells[1, 5].Value = "Diameter";

                worksheet.Cells[1, 1, 1, 5].AutoFilter = true;

                using (var headerRange = worksheet.Cells[1, 1, 1, 5])
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

                // 데이터
                int row = 2;
                foreach (var item in classifications)
                {
                    worksheet.Cells[row, 1].Value = item.Category;
                    worksheet.Cells[row, 2].Value = item.FamilyName;
                    worksheet.Cells[row, 3].Value = item.Classification;
                    worksheet.Cells[row, 4].Value = item.Description;
                    worksheet.Cells[row, 5].Value = item.Diameter;
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
