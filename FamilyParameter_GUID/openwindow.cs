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
    public class FamilyParameter : IExternalCommand
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

            if (_existingWindow != null && _existingWindow.IsVisible)
            {
                _existingWindow.Activate();
                return Result.Succeeded;
            }

            var handler1 = new FamilyParameterHandler();
            var handler2 = new FamilyParameterHandler2();

            
            var externalEvent1 = ExternalEvent.Create(handler1);
            var externalEvent2 = ExternalEvent.Create(handler2);

            var control = new MainUserControl();

            _existingWindow = new Window
            {
                Title = "Family Parameter 추출",
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

