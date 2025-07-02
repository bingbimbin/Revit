using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class ImportExcelHandler : IExternalEventHandler
    {
        public string ExcelPath { get; set; }
        public List<string> SelectedFields { get; set; } = new List<string>();

        public event Action<int> ProgressChanged;

        public void SetData(string excelPath, List<string> selectedFields)
        {
            ExcelPath = excelPath;
            SelectedFields = selectedFields;
        }


        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            if (string.IsNullOrWhiteSpace(ExcelPath) || !File.Exists(ExcelPath))
            {
                MessageBox.Show("엑셀 파일 경로가 잘못되었거나 지정되지 않았습니다.", "오류");
                return;
            }

            try
            {
                using (var package = new ExcelPackage(new FileInfo(ExcelPath)))
                {
                    var ws = package.Workbook.Worksheets.First();
                    int rowCount = ws.Dimension.Rows;
                    int colCount = ws.Dimension.Columns;

                    var headers = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                        headers.Add(ws.Cells[1, col].Text);

                    if (headers[0] != "ElementId")
                    {
                        MessageBox.Show("첫 번째 열은 반드시 'ElementId'여야 합니다.", "오류");
                        return;
                    }

                    int updatedCount = 0, failedCount = 0;

                    using (Transaction trans = new Transaction(doc, "Excel 기반 파라미터 수정"))
                    {
                        trans.Start();

                        int total = rowCount - 1;
                        int current = 0;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            current++;
                            int progress = (int)((double)current / total * 100);
                            ProgressChanged?.Invoke(progress);

                            string idText = ws.Cells[row, 1].Text;
                            if (!int.TryParse(idText, out int idInt)) continue;

                            Element elem = doc.GetElement(new ElementId(idInt));
                            if (elem == null) continue;

                            for (int col = 2; col <= colCount; col++)
                            {
                                string paramName = headers[col - 1];
                                if (!SelectedFields.Contains(paramName)) continue;

                                string newValue = ws.Cells[row, col].Text;
                                Parameter param = elem.LookupParameter(paramName);
                                if (param == null || param.IsReadOnly) continue;

                                try
                                {
                                    switch (param.StorageType)
                                    {
                                        case StorageType.String:
                                            if (param.AsString() != newValue)
                                            {
                                                param.Set(newValue);
                                                updatedCount++;
                                            }
                                            break;

                                        case StorageType.Double:
                                            if (double.TryParse(newValue, out double dVal))
                                            {
                                                if (Math.Abs(param.AsDouble() - dVal) > 1e-9)
                                                {
                                                    param.Set(dVal);
                                                    updatedCount++;
                                                }
                                            }
                                            break;

                                        case StorageType.Integer:
                                            if (int.TryParse(newValue, out int iVal))
                                            {
                                                if (param.AsInteger() != iVal)
                                                {
                                                    param.Set(iVal);
                                                    updatedCount++;
                                                }
                                            }
                                            break;

                                        case StorageType.ElementId:
                                            break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "오류");
                                }
                            }
                        }

                        trans.Commit();
                    }

                    MessageBox.Show("수정 완료","완료");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "오류");
            }
        }

        public string GetName() => "Import Excel Parameter Modifier";
    }
}
