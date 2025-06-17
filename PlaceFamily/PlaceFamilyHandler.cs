using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Linq;

namespace NewJeans
{
    public class PlaceFamilyHandler : IExternalEventHandler
    {
        private List<BuiltInCategory> _selectedCategories;
        public double SpacingX { get; set; }
        public double SpacingY { get; set; }

        public event Action<int> ProgressChanged;

        public void SetCategories(List<BuiltInCategory> cats)
        {
            _selectedCategories = cats;
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;


            int totalSymbolsCount = 0;

            foreach (var cat in _selectedCategories)
            {
                FilteredElementCollector symbolCollector = new FilteredElementCollector(doc)
                    .OfCategory(cat)
                    .OfClass(typeof(FamilySymbol));

                FilteredElementCollector levelCollector = new FilteredElementCollector(doc)
                                                                .OfClass(typeof(Level));
                Level targetLevel = levelCollector
                    .Cast<Level>()
                    .FirstOrDefault();

                totalSymbolsCount += symbolCollector.GetElementCount();
            }

            if (totalSymbolsCount == 0)
            {
                MessageBox.Show("선택한 카테고리에서 배치할 수 있는 Family가 없습니다!", "오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int placedCount = 0;
            int totalToPlace = _selectedCategories
                                .SelectMany(c => new FilteredElementCollector(doc).OfCategory(c).OfClass(typeof(FamilySymbol))
                                .Cast<FamilySymbol>())
                                .Count();

            using (Transaction trans = new Transaction(doc, "선택된 카테고리로 패밀리 배치"))
            {
                trans.Start();

                double spacingX = SpacingX;
                double spacingY = SpacingY;
                XYZ startPoint = new XYZ(0, 0, 0);

                int familyIndex = 0;

                foreach (var cat in _selectedCategories)
                {
                    FilteredElementCollector symbolCollector = new FilteredElementCollector(doc)
                        .OfCategory(cat)
                        .OfClass(typeof(FamilySymbol));

                    var symbolsByFamily = new Dictionary<ElementId, List<FamilySymbol>>();
                    foreach (FamilySymbol symbol in symbolCollector)
                    {
                        ElementId familyId = symbol.Family.Id;
                        if (!symbolsByFamily.ContainsKey(familyId))
                            symbolsByFamily[familyId] = new List<FamilySymbol>();
                        symbolsByFamily[familyId].Add(symbol);
                    }

                    
                    foreach (var kvp in symbolsByFamily)
                    {
                        List<FamilySymbol> symbols = kvp.Value;
                        double y = familyIndex * spacingY;
                        FilteredElementCollector levelCollector = new FilteredElementCollector(doc)
                                                                .OfClass(typeof(Level));
                        Level targetLevel = levelCollector
                            .Cast<Level>()
                            .FirstOrDefault();                      

                        for (int typeIndex = 0; typeIndex < symbols.Count; typeIndex++)
                        {
                            FamilySymbol symbol = symbols[typeIndex];

                            if (!symbol.IsActive)
                                symbol.Activate();

                            double x = typeIndex * spacingX;
                            XYZ point = startPoint + new XYZ(x, -y, 0);

                            doc.Create.NewFamilyInstance(point, symbol, targetLevel, StructuralType.NonStructural);

                            placedCount++;
                            int progress = (int)((double)placedCount / totalToPlace * 100);
                            ProgressChanged?.Invoke(progress);
                        }
                        familyIndex++;
                    }
                }

                trans.Commit();
            }

            MessageBox.Show("패밀리 배치가 완료되었습니다!", "완료", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public string GetName() => "Place Family Handler";
    }
}
