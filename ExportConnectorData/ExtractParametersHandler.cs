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
    public class ExtractParametersHandler : IExternalEventHandler
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

        public class ConnectorInfo
        {
            public string Category { get; set; }
            public string FamilyName { get; set; }
            public string Classification { get; set; }
            public string Description { get; set; }
            public string Diameter { get; set; }

            public ConnectorInfo(string category,string familyName, string classification, string description, string diameter)
            {
                Category = category;
                FamilyName = familyName;
                Classification = classification;
                Description = description;
                Diameter = diameter;
            }
        }

        public void Execute(UIApplication app)
        {
            if (SelectedFiles == null || string.IsNullOrWhiteSpace(ExportPath))
            {
                MessageBox.Show("파일을 선택하거나 경로를 지정해주세요.", "오류");
                return;
            }

            List<string> familiesnames = SelectedFiles;
            string excelPath = ExportPath;

            List<ConnectorInfo> classifications = new List<ConnectorInfo>();

            int total = familiesnames.Count;
            int current = 0;

            foreach (string family in familiesnames)
            {
                Document familyDoc = app.Application.OpenDocumentFile(family);
                if (familyDoc == null)
                {
                    MessageBox.Show($"패밀리 파일 {Path.GetFileName(family)}을(를) 열 수 없습니다.", "오류");
                    continue;
                }

                string familyCategory = familyDoc.OwnerFamily.FamilyCategory?.Name ?? "Unknown";
                string familyName = Path.GetFileNameWithoutExtension(family);

                List<ConnectorElement> col = new FilteredElementCollector(familyDoc)
                    .OfCategory(BuiltInCategory.OST_ConnectorElem)
                    .WhereElementIsNotElementType()
                    .Cast<ConnectorElement>()
                    .ToList();

                bool hasConnector = false;

                foreach (ConnectorElement e in col)
                {
                    Domain domain = e.Domain;

                    switch(domain)
                    {
                        case Domain.DomainPiping:
                            Parameter pipeparam1 = e.get_Parameter(BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM);
                            Parameter pipeparam2 = e.get_Parameter(BuiltInParameter.RBS_CONNECTOR_DESCRIPTION);
                            Parameter pipeparam3 = e.get_Parameter(BuiltInParameter.CONNECTOR_DIAMETER);

                            string pipeclassification = pipeparam1?.AsValueString() ?? "";
                            string pipedescription = pipeparam2?.AsString() ?? "";
                            string pipediameter = pipeparam3?.AsValueString() + "A" ?? "";

                            classifications.Add(new ConnectorInfo(familyCategory,familyName, pipeclassification, pipedescription, pipediameter));
                            hasConnector = true;
                            break;
                        
                        case Domain.DomainHvac:
                            Parameter ductparam1 = e.get_Parameter(BuiltInParameter.RBS_DUCT_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM);
                            Parameter ductparam2 = e.get_Parameter(BuiltInParameter.RBS_CONNECTOR_DESCRIPTION);
                            Parameter ductparam3 = e.get_Parameter(BuiltInParameter.CONNECTOR_DIAMETER);

                            string ductclassification = ductparam1?.AsValueString() ?? "";
                            string ductdescription = ductparam2?.AsString() ?? "";
                            string ductsize = "";
                            if(e.Shape == ConnectorProfileType.Rectangular || e.Shape == ConnectorProfileType.Oval)
                            {
                                double ductwidth = UnitUtils.ConvertFromInternalUnits(e.Width, DisplayUnitType.DUT_MILLIMETERS);
                                double ductheight = UnitUtils.ConvertFromInternalUnits(e.Height, DisplayUnitType.DUT_MILLIMETERS);

                                ductsize = $"{ductwidth}W x {ductheight}H";
                            }

                            if (e.Shape == ConnectorProfileType.Round)
                            {
                                ductsize = ductparam3?.AsValueString()+"A";
                            }

                            classifications.Add(new ConnectorInfo(familyCategory, familyName, ductclassification, ductdescription, ductsize));
                            hasConnector = true;
                            break;


                    }
                    
                }

                if (!hasConnector)
                {
                    classifications.Add(new ConnectorInfo(familyCategory, familyName, "", "", ""));
                }

                familyDoc.Close(false);
                // 이벤트 호출
                current++;
                int progress = (int)((double)current / total * 100);
                ProgressChanged?.Invoke(progress);
            }

            try
            {
                ExcelExporter.Export(classifications, ExportPath);
                if (MessageBox.Show("엑셀 저장 완료!\n저장된 파일을 여시겠습니까?", "성공", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
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
                TaskDialog.Show("엑셀 저장 오류", ex.Message);
            }
        }

        public string GetName() => "Extract Parameters Handler";
    }
}
