using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Windows;
using System.Linq;

namespace NewJeans
{
    public class FamilyParameterHandler2 : IExternalEventHandler
    {
        public List<string> SelectedFiles { get; set; }
        public string ExportPath { get; set; }

        // ✅ 진행 상황 이벤트
        public event Action<int> ProgressChanged;

        public void SetData(List<string> selectedFiles, string exportPath)
        {
            SelectedFiles = selectedFiles;
            ExportPath = exportPath;
        }

        public class FamilyParameterInfo2
        {
            public string Category { get; set; }
            public string FamilyName { get; set; }
            public string TypeName { get; set; }
            public string ParamName { get; set; }

            public string Value { get; set; }
            public string Formula { get; set; }
            public string Guid { get; set; }

            public FamilyParameterInfo2(string category, string familyname, string typename, string paramName, string value, string formula, string guid)
            {
                Category = category;
                FamilyName = familyname;
                TypeName = typename;
                ParamName = paramName;
                Value = value;
                Formula = formula;
                Guid = guid;

            }
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<string> familiesnames = SelectedFiles;
            string excelPath = ExportPath;

            List<FamilyParameterInfo2> familyparameters = new List<FamilyParameterInfo2>();

            int total = familiesnames.Count;
            int current = 0;

            foreach (string family in familiesnames)
            {

                // Family 문서를 열고 FamilyManager를 가져옵니다
                Document famDoc = app.Application.OpenDocumentFile(family);
                FamilyManager fm = famDoc.FamilyManager;
                FamilyType currentType = fm.CurrentType;

                // FamilyType마다 반복

                var familyTypes = fm.Types;
                IList<FamilyType> types = familyTypes.Cast<FamilyType>().ToList();

                using (Transaction trans = new Transaction(famDoc, "패밀리 파라미터 추출"))
                {
                    trans.Start();
                    try
                    {
                        foreach (FamilyType type in types)
                        {
                            fm.CurrentType = type;

                            // 파라미터를 처리합니다
                            foreach (Autodesk.Revit.DB.FamilyParameter param in fm.Parameters)
                            {
                                string category = famDoc.OwnerFamily.FamilyCategory.Name;
                                string familyname = famDoc.Title;
                                string typename = type.Name;
                                string paramName = param.Definition.Name;
                                string value = "";
                                string formula = "";
                                string guid = "";

                                if (type.HasValue(param))
                                {
                                    if (type.AsInteger(param) == 1)
                                        value = "True";
                                    else if (type.AsInteger(param) == 0)
                                        value = "False";
                                    else
                                        value = type.AsValueString(param);
                                }
                                if (param.IsShared)
                                {
                                    guid = param.GUID.ToString();
                                }

                                if (param.IsDeterminedByFormula)
                                {
                                    try
                                    {
                                        formula = param.Formula;
                                    }
                                    catch { }
                                }
                                familyparameters.Add(new FamilyParameterInfo2(category, familyname, typename, paramName, value, formula, guid));
                            }
                        }

                        trans.Commit();
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();  // 에러 발생시 롤백
                        TaskDialog.Show("Error", $"Error in extracting family types: {ex.Message}");
                    }
                }
                famDoc.Close(false);

                // 진행 상황을 업데이트합니다
                current++;
                int progress = (int)((double)current / total * 100);
                ProgressChanged?.Invoke(progress);
            }

            try
            {
                // 엑셀 파일을 내보냅니다
                ExcelExporter2.Export(familyparameters, ExportPath);
                if (MessageBox.Show("엑셀 저장 완료!\n저장된 파일을 여시겠습니까?", "성공", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        // 엑셀 파일을 엽니다
                        Process.Start(ExportPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("엑셀을 여는 중 오류 발생:\n" + ex.Message, "오류");
                    }
                }
            }
            catch (Exception ex)
            {
                // 엑셀 저장 중 오류가 발생하면 보여줍니다
                TaskDialog.Show("엑셀 저장 오류", ex.Message);
            }
        }

        public string GetName() => "Extract Parameters Handler";
    }
}
