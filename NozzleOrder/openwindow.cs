using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media;
using System.Windows;
using System;


namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class NozzleOrderCheck : IExternalCommand
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

            var handler = new NozzleOrderHandler();
            var externalEvent = ExternalEvent.Create(handler);
            var control = new NozzleOrderResult(externalEvent, handler);


            if (_existingWindow != null && _existingWindow.IsVisible)
            {
                _existingWindow.Activate();
                return Result.Succeeded;
            }


            _existingWindow = new Window
            {
                Title = "Nozzle Connector 순서 검토",
                Content = control,
                Width = 450,
                Height = 400,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 47, 56))
            };

            _existingWindow.Closed += (s, e) => _existingWindow = null;
            _existingWindow.Show();
            return Result.Succeeded;
        }
    }
}


