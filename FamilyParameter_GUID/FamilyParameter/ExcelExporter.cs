using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using static NewJeans.FamilyParameterHandler;

namespace NewJeans
{
    public static class ExcelExporter
    {
        public static void Export(List<FamilyParameterInfo> familyparameters, string filePath)
        {

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Family Parameters");

                // 헤더 작성
                worksheet.Cells[1, 1].Value = "Category";
                worksheet.Cells[1, 2].Value = "Family";
                worksheet.Cells[1, 3].Value = "Type";
                worksheet.Cells[1, 4].Value = "Parameter";
                worksheet.Cells[1, 5].Value = "Value";
                worksheet.Cells[1, 6].Value = "Formula";
                worksheet.Cells[1, 7].Value = "GUID";

                worksheet.Cells[1, 1, 1, 7].AutoFilter = true;

                using (var headerRange = worksheet.Cells[1, 1, 1, 7])
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
                foreach (var item in familyparameters)
                {
                    worksheet.Cells[row, 1].Value = item.Category;
                    worksheet.Cells[row, 2].Value = item.FamilyName;
                    worksheet.Cells[row, 3].Value = item.TypeName;
                    worksheet.Cells[row, 4].Value = item.ParamName;
                    worksheet.Cells[row, 5].Value = item.Value;
                    worksheet.Cells[row, 6].Value = item.Formula;
                    worksheet.Cells[row, 7].Value = item.Guid;
                    row++;
                }

                // 열 너비 자동 조정
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    worksheet.Column(col).Width = 25;
                }

                // 저장
                var file = new FileInfo(filePath);
                package.SaveAs(file);
            }
        }
    }
}
