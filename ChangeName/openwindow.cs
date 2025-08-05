using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media;
using System.Windows;
using System;

namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class ChangeName : IExternalCommand
    {
        private static Window _existingWindow;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ExtractHandler handler = new ExtractHandler();
            ExternalEvent externalEvent = ExternalEvent.Create(handler);

            if (_existingWindow != null && _existingWindow.IsVisible)
            {
                _existingWindow.Activate();
                return Result.Succeeded;
            }

            var control = new ChangeNameResult(externalEvent, handler);

            _existingWindow = new Window
            {
                Title = "패밀리, 타입이름 변경",
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


