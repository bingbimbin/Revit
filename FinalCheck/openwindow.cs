using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media;
using System.Windows;
using System;

namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class FinalCheck : IExternalCommand
    {
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

            var handler = new FinalCheckHandler();
            var externalEvent = ExternalEvent.Create(handler);

            var control = new FinalCheckResult(externalEvent, handler);

            var window = new Window
            {
                Title = "납품파일 최종검토",
                Content = control,
                Width = 450,
                Height = 550,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 47, 56))
            };

            window.Show(); // 창을 닫지 않음
            return Result.Succeeded;
        }
    }
}

