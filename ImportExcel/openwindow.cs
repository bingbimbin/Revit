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
    public class ImportExcel : IExternalCommand
    {
        private static Window _existingWindow;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var dummy = new OfficeOpenXml.ExcelPackage();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("EPPlus 테스트", ex.ToString());
            }

            var handler = new ImportExcelHandler();
            var externalEvent = ExternalEvent.Create(handler);
            var control = new ImportExcelResult(externalEvent, handler);

            if (_existingWindow != null && _existingWindow.IsVisible)
            {
                _existingWindow.Activate();
                return Result.Succeeded;
            }

            _existingWindow = new Window
            {
                Title = "Import Excel",
                Content = control,
                Width = 350,
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
