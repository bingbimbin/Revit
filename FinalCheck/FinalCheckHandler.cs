using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace NewJeans
{
    public class FinalCheckHandler : IExternalEventHandler
    {
        public List<string> SelectedFiles { get; set; }
        public string ExportPath { get; set; }

        public event Action<int> ProgressChanged;

        public void SetData(List<string> selectedFiles, string exportPath)
        {
            SelectedFiles = selectedFiles;
            ExportPath = exportPath;
        }

        public void Execute(UIApplication app)
        {
            if (SelectedFiles == null || string.IsNullOrWhiteSpace(ExportPath))
            {
                MessageBox.Show("파일을 선택하거나 경로를 지정해주세요.", "오류");
                return;
            }

            var appObj = app.Application;
            var finalResults = new List<FinalData>();

            OpenOptions openOptions = new OpenOptions();
            openOptions.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

            int total = SelectedFiles.Count;
            int count = 0;

            foreach (string file in SelectedFiles)
            {
                try
                {
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(file);
                    Document doc = appObj.OpenDocumentFile(modelPath, openOptions);

                    View viewToCheck = null;

                    StartingViewSettings svs = new FilteredElementCollector(doc)
                        .OfClass(typeof(StartingViewSettings))
                        .Cast<StartingViewSettings>()
                        .FirstOrDefault();

                    viewToCheck = doc.GetElement(svs?.ViewId) as View;

                    if (viewToCheck == null)
                    {
                        viewToCheck = new FilteredElementCollector(doc)
                            .OfClass(typeof(View3D))
                            .Cast<View3D>()
                            .FirstOrDefault(v => !v.IsTemplate);
                    }

                    ElementId startViewId = svs?.ViewId ?? ElementId.InvalidElementId;
                    View startingView = doc.GetElement(startViewId) as View;

                  
                    string startingViewName = startingView != null ? startingView.Name : "지정 안됨";

                    string detailLevel = viewToCheck?.DetailLevel.ToString() ?? "Unknown";
                    string displayStyle = viewToCheck?.DisplayStyle.ToString() ?? "Unknown";

                    int revitLinkCount = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkType))
                        .Count();

                    int cadInstanceCount = new FilteredElementCollector(doc)
                        .OfClass(typeof(ImportInstance))
                        .Cast<ImportInstance>()
                        .Count();

                    int filterCount = new FilteredElementCollector(doc)
                        .OfClass(typeof(ParameterFilterElement))
                        .Cast<ParameterFilterElement>()
                        .Count();

                    int designoptionCount = new FilteredElementCollector(doc)
                       .OfCategory(BuiltInCategory.OST_DesignOptionSets)
                       .WhereElementIsNotElementType()
                       .ToList()
                       .Count();

                    // 필터 정보 수집--------------------------------------------------------------------------------------------
                    string viewFilterInfo = string.Empty;

                    ICollection<ElementId> filterIds = startingView.GetFilters();
                    int viewFilterCount = filterIds.Count;

                    if (viewFilterCount == 1)
                    {
                        ElementId filterId = filterIds.First();
                        ParameterFilterElement filter = doc.GetElement(filterId) as ParameterFilterElement;

                        if (filter != null)
                        {
                            string filterName = filter.Name;
                            bool isVisible = startingView.GetFilterVisibility(filterId); // true → ON, false → OFF

                            viewFilterInfo = $"필터 이름 : {filterName} / Visibility : {(isVisible ? "ON" : "OFF")}";
                        }
                    }
                    else
                    {
                        viewFilterInfo = $"{viewFilterCount}";
                    }
                    //-----------------------------------------------------------------------------------------------------------

                    // 미사용 문자타입--------------------------------------------------------------------------------------------

                    var allTextTypes = new FilteredElementCollector(doc)
                                .OfClass(typeof(TextNoteType))
                                .Cast<TextNoteType>()
                                .ToList();


                    var allTextNotes = new FilteredElementCollector(doc)
                                        .OfClass(typeof(TextNote))
                                        .Cast<TextNote>();

                    var usedTextTypeIds = new HashSet<ElementId>();
                    foreach (var tn in allTextNotes)
                    {
                        usedTextTypeIds.Add(tn.GetTypeId());
                    }


                    var unusedTextTypes = allTextTypes
                                            .Where(t => !usedTextTypeIds.Contains(t.Id))
                                            .ToList();


                    string unusedTextTypeInfo;

                    if (unusedTextTypes.Count == 0)
                    {
                        unusedTextTypeInfo = "0";
                    }
                    else if (unusedTextTypes.Count == 1)
                    {
                        unusedTextTypeInfo = unusedTextTypes[0].Name;
                    }
                    else
                    {
                        unusedTextTypeInfo = $"{unusedTextTypes.Count}";
                    }

                    //-----------------------------------------------------------------------------------------------------------

                    // 미사용 뷰필터---------------------------------------------------------------------------------------------
                    var allFilters = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>();

                    var allViews = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate
                            && (v is View3D || v is ViewPlan || v is ViewSection));

                    var usedFilterIds = new HashSet<ElementId>();

                    foreach (var v in allViews)
                    {
                        foreach (var fid in v.GetFilters())
                            usedFilterIds.Add(fid);
                    }
                    int unusedFilterCount = allFilters.Count(f => !usedFilterIds.Contains(f.Id));
                    //-----------------------------------------------------------------------------------------------------------

                    // 미사용 뷰탬플릿--------------------------------------------------------------------------------------------
                    var allTemplates = new FilteredElementCollector(doc)
                                        .OfClass(typeof(View))
                                        .Cast<View>()
                                        .Where(v => v.IsTemplate)   // 뷰 템플릿만
                                        .ToList();

                    var usedTemplateIds = new HashSet<ElementId>();
                    foreach (var v in allViews)
                    {
                        if (v.ViewTemplateId != ElementId.InvalidElementId)
                            usedTemplateIds.Add(v.ViewTemplateId);
                    }

                    int unusedViewTemplateCount = allTemplates.Count(t => !usedTemplateIds.Contains(t.Id));
                    //-----------------------------------------------------------------------------------------------------------

                    
                    var categoryResults = CheckCategories(doc, viewToCheck);

                    finalResults.Add(new FinalData
                    {
                        Name = System.IO.Path.GetFileNameWithoutExtension(file),
                        DetallLevel = detailLevel,
                        GraphicsStyle = displayStyle,
                        StartingViewName = startingViewName,
                        RevitLinkCount = revitLinkCount,
                        CadInstanceCount = cadInstanceCount,
                        DesignOptionCount = designoptionCount,
                        FilterCount = filterCount,
                        ViewFilterInfo = viewFilterInfo,
                        UnusedTextTypeInfo = unusedTextTypeInfo,
                        UnusedFilterCount = unusedFilterCount,
                        UnusedViewTemplateCount = unusedViewTemplateCount,

                        Mass = categoryResults["Mass"],
                        Line = categoryResults["Line"],
                        Site = categoryResults["Site"],
                        Part = categoryResults["Part"],

                        Doors = categoryResults["Doors"],
                        CableTrayFitting = categoryResults["CableTrayFitting"],
                        ConduitFitting = categoryResults["ConduitFitting"],
                        DuctFitting = categoryResults["DuctFitting"],
                        PipeFitting = categoryResults["PipeFitting"]
                    });

                    doc.Close(false);


                }
                catch (Exception ex)
                {
                    MessageBox.Show($"파일명 : {file}\n{ex.Message}", "파일 열기 오류");
                }

                count++;
                int progress = (int)((count / (double)total) * 100);
                ProgressChanged?.Invoke(progress);
            }

            try
            {
                ExcelExporter.Export(finalResults, ExportPath);

                if (MessageBox.Show("엑셀 저장 완료!\n저장된 파일을 여시겠습니까?", "성공", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(ExportPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("엑셀을 여는 중 오류 발생:\n" + ex.Message, "오류");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "엑셀 저장 오류");
            }
        }

        private Dictionary<string, string> CheckCategories(Document doc, View view)
        {
            var results = new Dictionary<string, string>();

            // 1. Mass, Line, Site, Part 체크
            var basicCategories = new Dictionary<string, BuiltInCategory>
            {
                { "Mass", BuiltInCategory.OST_Mass },
                { "Line", BuiltInCategory.OST_Lines },
                { "Site", BuiltInCategory.OST_Site },
                { "Part", BuiltInCategory.OST_Parts }
            };

            foreach (var kvp in basicCategories)
            {
                var cat = doc.Settings.Categories.get_Item(kvp.Value);
                bool hidden = view.GetCategoryHidden(cat.Id);
                results[kvp.Key] = hidden ? "숨김O" : "숨김X";
            }

            // 2. Fitting류 체크
            var fittingCategories = new Dictionary<string, BuiltInCategory>
            {
                { "Doors", BuiltInCategory.OST_Doors},
                { "CableTrayFitting", BuiltInCategory.OST_CableTrayFitting },
                { "ConduitFitting", BuiltInCategory.OST_ConduitFitting },
                { "DuctFitting", BuiltInCategory.OST_DuctFitting },
                { "PipeFitting", BuiltInCategory.OST_PipeFitting }
            };

            foreach (var kvp in fittingCategories)
            {
                var fitCat = doc.Settings.Categories.get_Item(kvp.Value);
                bool pass = true;

                foreach (Category subCat in fitCat.SubCategories)
                {
                    bool hidden = view.GetCategoryHidden(subCat.Id);

                    if (subCat.Name.Contains("End") || subCat.Name.Contains("End-Cut"))
                    {
                        if (!hidden) { pass = false; break; }
                    }
                    else
                    {
                        if (hidden) { pass = false; break; }
                    }
                }

                results[kvp.Key] = pass ? "True" : "False";
            }

            return results;
        }


        public string GetName() => "Final Check Handler";
    }

    public class FinalData
    {
        public string Name { get; set; }
        public string DetallLevel { get; set; }
        public string GraphicsStyle { get; set; }
        public string StartingViewName { get; set; }
        public int RevitLinkCount { get; set; }
        public int CadInstanceCount { get; set; }
        public int DesignOptionCount { get; set; }
        public int FilterCount { get; set; }
        public string ViewFilterInfo { get; set; }
        public string UnusedTextTypeInfo { get; set; }
        public int UnusedFilterCount { get; set; }
        public int UnusedViewTemplateCount { get; set; }

        // 카테고리별 검사 결과
        public string Mass { get; set; }
        public string Line { get; set; }
        public string Site { get; set; }
        public string Part { get; set; }
        public string Doors { get; set; }
        public string CableTrayFitting { get; set; }
        public string ConduitFitting { get; set; }
        public string DuctFitting { get; set; }
        public string PipeFitting { get; set; }
    }
}
