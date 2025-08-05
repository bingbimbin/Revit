using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.ComponentModel;
using System.Windows;
using System.Diagnostics;

namespace NewJeans
{
    public class ExtractHandler : IExternalEventHandler
    {
        public List<string> SelectedFiles { get; set; }
        public string SelectedReferencePath { get; set; }
        public string ExportPath { get; set; }

        // ✅ 진행 상황 이벤트
        public event Action<int> ProgressChanged;

        public void SetData(List<string> selectedFiles, string selectedReferencePath, string exportPath)
        {
            SelectedFiles = selectedFiles;
            SelectedReferencePath = selectedReferencePath;
            ExportPath = exportPath;
        }


        public void Execute(UIApplication app)
        {
            try
            {
                if (SelectedFiles == null || string.IsNullOrWhiteSpace(ExportPath) || string.IsNullOrWhiteSpace(SelectedReferencePath))
                {
                    MessageBox.Show("파일을 선택하거나 기준파일, 경로를 정해주세요", "오류");
                    return;
                }


                var mappings = new List<(string oldName, string oldType, string newName, string newTypeName)>();

                using (var package = new ExcelPackage(new FileInfo(SelectedReferencePath)))
                {
                    var worksheet = package.Workbook.Worksheets[1];  // 첫 번째 워크시트

                    int rowCount = worksheet.Dimension.End.Row;

                    for (int row = 2; row <= rowCount; row++)  // 2행부터 시작 (1행은 헤더)
                    {
                        string oldName = worksheet.Cells[row, 1].Text;
                        string oldType = worksheet.Cells[row, 2].Text;
                        string newName = worksheet.Cells[row, 3].Text;
                        string newType = worksheet.Cells[row, 4].Text;

                        mappings.Add((oldName, oldType, newName, newType));
                    }
                }

                int total = SelectedFiles.Count;
                int count = 0;

                var notFoundFamilies = new List<string>();

                foreach (var (oldName, oldType, newName, newType) in mappings)
                {
                    string rfaPath = SelectedFiles.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f) == oldName);

                    if (string.IsNullOrEmpty(rfaPath) || !File.Exists(rfaPath))
                    {
                        notFoundFamilies.Add(oldName);
                        continue;
                    }


                    Document famDoc = app.Application.OpenDocumentFile(rfaPath);

                    using (Transaction trans = new Transaction(famDoc, "패밀리 타입명 변경"))
                    {
                        trans.Start();

                        FamilyManager familyManager = famDoc.FamilyManager;

                        if (string.IsNullOrWhiteSpace(oldType))
                        {
                            familyManager.RenameCurrentType(newType);
                        }

                        else
                        {
                            FamilyType targetType = familyManager.Types.Cast<FamilyType>().FirstOrDefault(t => t.Name == oldType);
                            if (targetType != null)
                            {
                                familyManager.CurrentType = targetType;
                                familyManager.RenameCurrentType(newType);
                            }

                            else if (oldType == "*")
                            {
                                // 패스하기
                            }
                            else
                            {
                                MessageBox.Show($"'{oldName}' 패밀리에 '{oldType}' 타입이 없습니다.", "오류");
                            }
                        }

                        trans.Commit();
                    }

                    if (oldName == newName)
                    {
                        string savePath = Path.Combine(ExportPath, newName + ".rfa");
                        famDoc.SaveAs(savePath);
                        famDoc.Close(false);
                    }

                    else
                    {
                        string savePath = Path.Combine(ExportPath, newName + ".rfa");
                        famDoc.SaveAs(savePath);
                        famDoc.Close(false);
                    }

                    count++;
                    int progress = (int)((count / (double)total) * 100);


                    // ✅ 이벤트 호출
                    ProgressChanged?.Invoke(progress);
                }

                if (notFoundFamilies.Any())
                {
                    string result = string.Join(Environment.NewLine, notFoundFamilies);

                    // ExportPath에 텍스트 파일 저장
                    string savePath = Path.Combine(ExportPath, "NotFoundFamilies.txt");
                    File.WriteAllText(savePath, result);

                    // 메모장으로 열기
                    Process.Start("notepad.exe", savePath);
                }


                MessageBox.Show("패밀리, 타입명 수정 후 저장 완료", "완료");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "오류");
            }


        }

        public string GetName()
        {
            return "Family Type Renaming Handler";
        }
    }
}
