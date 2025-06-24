using Autodesk.Revit.DB;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using static NewJeans.FamilyParameterHandler;

namespace NewJeans
{
    public static class ExcelExporter
    {
        public static void Export(List<FamilyParameterInfo> familyparameters, List<LabelParameterInfo> labelParameters, List<ErrorInfo> errorInfos ,string filePath)
        {
            using (var package = new ExcelPackage())
            {
                // ======================
                // 시트 1: Family Parameters
                // ======================
                var ws1 = package.Workbook.Worksheets.Add("Family Type Parameters");

                ws1.Cells[1, 1].Value = "Category";
                ws1.Cells[1, 2].Value = "Family";
                ws1.Cells[1, 3].Value = "Type";
                ws1.Cells[1, 4].Value = "Parameter";
                ws1.Cells[1, 5].Value = "Value";

                ws1.Cells[1, 1, 1, 5].AutoFilter = true;

                ApplyHeaderStyle(ws1, 1, 5);

                int row1 = 2;
                foreach (var item in familyparameters)
                {
                    ws1.Cells[row1, 1].Value = item.Category;
                    ws1.Cells[row1, 2].Value = item.FamilyName;
                    ws1.Cells[row1, 3].Value = item.TypeName;
                    ws1.Cells[row1, 4].Value = item.ParamName;
                    ws1.Cells[row1, 5].Value = item.Value;
                    row1++;
                }

                ws1.Cells[ws1.Dimension.Address].AutoFitColumns();

                // ======================
                // 시트 2: Label Parameters
                // ======================
                var ws2 = package.Workbook.Worksheets.Add("Label Parameters");

                ws2.Cells[1, 1].Value = "상위 패밀리 Category";
                ws2.Cells[1, 2].Value = "상위 패밀리";
                ws2.Cells[1, 3].Value = "하위 패밀리";
                ws2.Cells[1, 4].Value = "하위 패밀리 ID";
                ws2.Cells[1, 5].Value = "Parameter";
                ws2.Cells[1, 6].Value = "Value";

                ws2.Cells[1, 1, 1, 6].AutoFilter = true;

                ApplyHeaderStyle(ws2, 1, 6);

                int row2 = 2;
                foreach (var item in labelParameters)
                {
                    ws2.Cells[row2, 1].Value = item.Category;
                    ws2.Cells[row2, 2].Value = item.SuperFamilyName;
                    ws2.Cells[row2, 3].Value = item.SubFamilyTypeName;
                    ws2.Cells[row2, 4].Value = item.SubId;
                    ws2.Cells[row2, 5].Value = item.ParamName;
                    ws2.Cells[row2, 6].Value = item.Value;
                    row2++;
                }

                ws2.Cells[ws2.Dimension.Address].AutoFitColumns();


                var ws3 = package.Workbook.Worksheets.Add("Erro Infos");

                ws3.Cells[1, 1].Value = "Family";
                ws3.Cells[1, 2].Value = "Type";
                ws3.Cells[1, 3].Value = "Error Message";

                ApplyHeaderStyle(ws3, 1, 3);

                int row3 = 2;
                foreach (var item in errorInfos)
                {
                    ws3.Cells[row3, 1].Value = item.FamilyName;
                    ws3.Cells[row3, 2].Value = item.TypeName;
                    ws3.Cells[row3, 3].Value = item.Error;
                    row3++;
                }

                ws3.Cells[ws3.Dimension.Address].AutoFitColumns();

                var file = new FileInfo(filePath);
                package.SaveAs(file);
            }
        }

        // ✅ 공통 헤더 스타일 적용 함수
        private static void ApplyHeaderStyle(ExcelWorksheet ws, int row, int colCount)
        {
            using (var range = ws.Cells[row, 1, row, colCount])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var border = range.Style.Border;
                border.Top.Style = ExcelBorderStyle.Thin;
                border.Bottom.Style = ExcelBorderStyle.Thin;
                border.Left.Style = ExcelBorderStyle.Thin;
                border.Right.Style = ExcelBorderStyle.Thin;
            }
        }
    }
}
