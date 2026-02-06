using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Media;

namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class DeleteParameter : IExternalCommand
    {
        private static Window _existingWindow;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            if (_existingWindow != null && _existingWindow.IsVisible)
            {
                _existingWindow.Activate();
                return Result.Succeeded;
            }

            // ✅ 핸들러 각각 생성
            var loadHandler = new LoadFamiliesHandler();
            var deleteHandler = new DeleteParameterHandler();

            // ✅ ExternalEvent도 각각 생성
            var loadEvent = ExternalEvent.Create(loadHandler);
            var deleteEvent = ExternalEvent.Create(deleteHandler);

            // ✅ UserControl에 두 개 다 전달
            var control = new DeleteParameterResult(loadEvent, loadHandler, deleteEvent, deleteHandler);

            _existingWindow = new Window
            {
                Title = "Family Parameter 삭제",
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
