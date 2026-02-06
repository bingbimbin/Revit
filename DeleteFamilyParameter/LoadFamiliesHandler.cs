using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace NewJeans
{
    public class LoadFamiliesHandler : IExternalEventHandler
    {
        public Action<List<string>> Callback { get; set; }
        public void Execute(UIApplication app)
        {
            try
            {
                Document doc = app.ActiveUIDocument.Document;

                var familyNames = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Select(f => f.Name)
                    .Distinct()
                    .OrderBy(n => n)
                    .ToList();


                Callback?.Invoke(familyNames);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"패밀리 로딩 중 오류 발생: {ex.Message}", "오류");
            }
        }

        public string GetName() => "Family Parameter Handler";
    }
}
