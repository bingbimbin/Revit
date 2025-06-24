using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Windows;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace NewJeans
{
    public class FamilyParameterHandler : IExternalEventHandler
    {
        public string ExportPath { get; set; }

        // ✅ 진행 상황 이벤트
        public event Action<int> ProgressChanged;

        public void SetData(string exportPath)
        {
            ExportPath = exportPath;
        }

        public class FamilyParameterInfo
        {
            public string Category { get; set; }
            public string FamilyName { get; set; }
            public string TypeName { get; set; }
            public string ParamName { get; set; }

            public string Value { get; set; }

            public FamilyParameterInfo(string category, string familyname, string typename, string paramName, string value)
            {
                Category = category;
                FamilyName = familyname;
                TypeName = typename;
                ParamName = paramName;
                Value = value;
            }
        }
        public class LabelParameterInfo
        {
            public string Category { get; set; }
            public string SuperFamilyName { get; set; }
            public string SubFamilyTypeName { get; set; }
            public string SubId { get; set; }
            public string ParamName { get; set; }
            public string Value { get; set; }

            public LabelParameterInfo(string category, string superfamilyname, string subfamilyname, string subid, string paramName, string value)
            {
                Category = category;
                SuperFamilyName = superfamilyname;
                SubFamilyTypeName = subfamilyname;  
                SubId = subid;
                ParamName = paramName;
                Value = value;
            }
        }
        public class ErrorInfo
        {
            public string FamilyName { get; set; }
            public string TypeName { get; set; }
            public string Error { get; set; }

            public ErrorInfo(string familyname, string typename, string error)
            {
                FamilyName = familyname;
                TypeName = typename;
                Error = error;
            }
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            string excelPath = ExportPath;

            List<FamilyParameterInfo> familyparameters = new List<FamilyParameterInfo>();
            List<LabelParameterInfo> labelsparameters = new List<LabelParameterInfo>();
            var families = new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>().Distinct().ToList();
            List<ErrorInfo> errorLogs = new List<ErrorInfo>();


            int total = families.Where(family => family.IsEditable).ToList().Count;
            int current = 0;

            foreach (Family family in families)
            {
                if (!family.IsEditable)
                    continue;

                // Family 문서를 열고 FamilyManager를 가져옵니다
                Document famDoc = doc.EditFamily(family);
                FamilyManager fm = famDoc.FamilyManager;

                // FamilyType마다 반복
                var familyTypes = fm.Types;
                IList<FamilyType> types = familyTypes.Cast<FamilyType>().ToList();

                // FamilyType 파라미터 추출
                using (Transaction trans = new Transaction(famDoc, "패밀리 타입 파라미터 추출"))
                {
                    trans.Start();
                        
                        List<FamilyInstance> nestedElements = new FilteredElementCollector(famDoc)
                                                .OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                .Where(fi => fi.SuperComponent == null).ToList();

                        foreach (FamilyType type in types)
                        {
                            try
                            {
                                fm.CurrentType = type;
                                famDoc.Regenerate();

                            // 파라미터를 처리합니다
                            foreach (FamilyParameter param in fm.Parameters)
                                {
                                    if (param.Definition.ParameterType == ParameterType.FamilyType)
                                    {
                                        string category = family.FamilyCategory.Name;
                                        string familyname = family.Name;
                                        string typename = type.Name;
                                        string paramName = param.Definition.Name;
                                        string value = "";

                                        

                                    foreach (FamilyInstance familyinstance in nestedElements)
                                    {
                                        foreach (Parameter parameter in familyinstance.Parameters)
                                        {
                                            if (parameter.Definition?.Name != "Label") continue;

                                            string subfamilytypename = familyinstance.Symbol.Family.Name;
                                            string subid = familyinstance.Id.ToString();
                                            string paramName2 = parameter.Definition.Name;
                                            string value2 = parameter.AsValueString() ?? "(null)";
                                            
                                            labelsparameters.Add(new LabelParameterInfo(category, familyname, subfamilytypename, subid, paramName2, value2));
                                        }
                                    }

                                    ElementId typeId = fm.CurrentType.AsElementId(param);
                                        if (typeId != ElementId.InvalidElementId)
                                        {
                                            var elem = famDoc.GetElement(typeId);
                                            value = elem?.Name ?? "(null)";
                                        }
                                        familyparameters.Add(new FamilyParameterInfo(category, familyname, typename, paramName, value));
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                string familyname = family.Name;
                                string typename = type.Name;
                                string error = ex.Message;
                                errorLogs.Add(new ErrorInfo(familyname, typename, error));
                            }
                        }

                    trans.Commit();
                    
                }
                // 가족 문서가 완료되면 닫습니다 (한 번만 호출)
                famDoc.Close(false);

                // 진행 상황을 업데이트합니다
                current++;
                int progress = (int)((double)current / total * 100);
                ProgressChanged?.Invoke(progress);
            }
            labelsparameters = labelsparameters
                .GroupBy(x => x.SubId)
                .Select(g => g.First()) 
                .ToList();
            try
            {
                // 엑셀 파일을 내보냅니다
                ExcelExporter.Export(familyparameters,labelsparameters, errorLogs, ExportPath);
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
