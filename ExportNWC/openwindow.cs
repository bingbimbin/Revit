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

namespace NewJeans
{

    [Transaction(TransactionMode.Manual)]
    public class ExportNWC : IExternalCommand
    {
        private static Window _existingWindow;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            if (_existingWindow != null && _existingWindow.IsVisible)
            {
                _existingWindow.Activate();
                return Result.Succeeded;
            }

            var exportHandler = new ExportNWCHandler();
            var exportEvent = ExternalEvent.Create(exportHandler);

            var linkHandler = new ManageLinkHandler();
            var linkEvent = ExternalEvent.Create(linkHandler);

            // ✅ 핸들러 인스턴스를 UI에 전달
            var control = new MainUserControl(commandData.Application, exportEvent ,exportHandler, linkEvent,linkHandler);

            _existingWindow = new Window
            {
                Title = "NWC 일괄 추출",
                Content = control,
                Width = 600,
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
