using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;


namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class PlaceFamily : IExternalCommand
    {
        private static Window _existingWindow;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (_existingWindow != null && _existingWindow.IsVisible)
            {
                _existingWindow.Activate();
                return Result.Succeeded;
            }

            // ✅ 핸들러 각각 생성
            var loadHandler = new PlaceFamilyHandler();

            // ✅ ExternalEvent도 각각 생성
            var loadEvent = ExternalEvent.Create(loadHandler);

            // ✅ UserControl에 두 개 다 전달
            var control = new PlaceFamilyResult(loadEvent, loadHandler, doc);

            _existingWindow = new Window
            {
                Title = "Category별 패밀리 배치",
                Content = control,
                Width = 450,
                Height = 500,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 47, 56))
            };

            _existingWindow.Closed += (s, e) => _existingWindow = null;
            _existingWindow.Show();
            return Result.Succeeded;
        }

    }

}
