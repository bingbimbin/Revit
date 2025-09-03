using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace NewJeans
{
    public static class ExcelExporter
    {
        public static void Export(List<FinalData> results1, string filePath)
        {


            using (var package = new ExcelPackage())
            {
                // 첫 번째 워크시트
                var worksheet1 = package.Workbook.Worksheets.Add("Project Parameters");
                FillWorksheet(worksheet1, results1);

                FileInfo file = new FileInfo(filePath);
                package.SaveAs(file);

            }

        }
        private static void FillWorksheet(ExcelWorksheet worksheet, List<FinalData> results)
        {
            worksheet.Cells[1, 1].Value = "File Name";
            worksheet.Cells[1, 2].Value = "Detail Level";
            worksheet.Cells[1, 3].Value = "Graphics Style";
            worksheet.Cells[1, 4].Value = "Starting View 이름";
            worksheet.Cells[1, 5].Value = "Revit Link 갯수";
            worksheet.Cells[1, 6].Value = "CAD Instance 갯수";
            worksheet.Cells[1, 7].Value = "Design Option 갯수";
            worksheet.Cells[1, 8].Value = "전체 Filter 갯수";
            worksheet.Cells[1, 9].Value = "Starting View Filter 적용 현황";
            worksheet.Cells[1, 10].Value = "미사용 문자타입 갯수";
            worksheet.Cells[1, 11].Value = "미사용 Filter 갯수";
            worksheet.Cells[1, 12].Value = "미사용 뷰탬플릿 갯수";
            worksheet.Cells[1, 13].Value = "Mass";
            worksheet.Cells[1, 14].Value = "Lines";
            worksheet.Cells[1, 15].Value = "Site";
            worksheet.Cells[1, 16].Value = "Parts";
            worksheet.Cells[1, 17].Value = "Doors";
            worksheet.Cells[1, 18].Value = "Cable Tray Fittings";
            worksheet.Cells[1, 19].Value = "Conduit Fittings";
            worksheet.Cells[1, 20].Value = "Duct Fittings";
            worksheet.Cells[1, 21].Value = "Pipe Fittings";

            for (int i = 0; i < results.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = results[i].Name;
                worksheet.Cells[i + 2, 2].Value = results[i].DetallLevel;
                worksheet.Cells[i + 2, 3].Value = results[i].GraphicsStyle;
                worksheet.Cells[i + 2, 4].Value = results[i].StartingViewName;
                worksheet.Cells[i + 2, 5].Value = results[i].RevitLinkCount;
                worksheet.Cells[i + 2, 6].Value = results[i].CadInstanceCount;
                worksheet.Cells[i + 2, 7].Value = results[i].DesignOptionCount;
                worksheet.Cells[i + 2, 8].Value = results[i].FilterCount;
                worksheet.Cells[i + 2, 9].Value = results[i].ViewFilterInfo;
                worksheet.Cells[i + 2, 10].Value = results[i].UnusedTextTypeInfo;
                worksheet.Cells[i + 2, 11].Value = results[i].UnusedFilterCount;
                worksheet.Cells[i + 2, 12].Value = results[i].UnusedViewTemplateCount;
                worksheet.Cells[i + 2, 13].Value = results[i].Mass;
                worksheet.Cells[i + 2, 14].Value = results[i].Line;
                worksheet.Cells[i + 2, 15].Value = results[i].Site;
                worksheet.Cells[i + 2, 16].Value = results[i].Part;
                worksheet.Cells[i + 2, 17].Value = results[i].Doors;
                worksheet.Cells[i + 2, 18].Value = results[i].CableTrayFitting;
                worksheet.Cells[i + 2, 19].Value = results[i].ConduitFitting;
                worksheet.Cells[i + 2, 20].Value = results[i].DuctFitting;
                worksheet.Cells[i + 2, 21].Value = results[i].PipeFitting;
            }

            // 헤더 스타일 적용


            using (var range = worksheet.Cells[1, 1, 1, 21])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                var border = range.Style.Border;
                border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            // 자동 열 너비 조정
            worksheet.Cells.AutoFitColumns();
            worksheet.Cells[1, 1, 1, 21].AutoFilter = true;
        }
    }
}
