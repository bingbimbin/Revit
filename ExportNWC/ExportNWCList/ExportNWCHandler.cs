using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;

namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class ExportNWCHandler : IExternalEventHandler
    {
        public event Action<int> ProgressChanged;

        private List<View3D> _viewsToExport = new List<View3D>();
        private string _exportPath;
        private bool _overwrite = true;

        public void SetData(List<View3D> selectedViews, string exportPath, bool overwrite)
        {
            _viewsToExport = selectedViews;
            _exportPath = exportPath;
            _overwrite = overwrite;
        }

        public void Execute(UIApplication uiApp)
        {
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            string exportFolder = _exportPath;
            if (string.IsNullOrEmpty(exportFolder) || !Directory.Exists(exportFolder))
            {
                MessageBox.Show("유효한 폴더 경로가 아닙니다.", "Error");
                return;
            }

            int total = _viewsToExport.Count;
            int count = 0;

            foreach (var view in _viewsToExport)
            {
                try
                {
                    string safeName = GetSafeFileName(view.Name);
                    string filename = safeName + ".nwc";

                    var options = new NavisworksExportOptions
                    {
                        ExportScope = NavisworksExportScope.View,
                        ViewId = view.Id,
                        ExportElementIds = true,
                        ExportLinks = true,
                        ExportRoomGeometry = false,
                        Coordinates = NavisworksCoordinates.Shared
                    };

                    string fullPath = Path.Combine(exportFolder, filename);

                    if (File.Exists(fullPath))
                    {
                        if (_overwrite)
                        {
                            File.Delete(fullPath); 
                        }
                        else
                        {
                            var result = MessageBox.Show($"{filename} 파일이 이미 존재합니다.\n덮어쓰시겠습니까?",
                                                         "파일 덮어쓰기 확인",
                                                         MessageBoxButton.YesNo,
                                                         MessageBoxImage.Question);
                            if (result == MessageBoxResult.Yes)
                            {
                                File.Delete(fullPath);
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }

                    doc.Export(exportFolder, filename, options);


                }
                catch (Exception ex)
                {
                    MessageBox.Show($"뷰 '{view.Name}' 내보내기 실패:\n{ex.Message}", "Export 실패");
                }

                count++;
                int progress = (int)((double)count / total * 100);
                ProgressChanged?.Invoke(progress);
            }


            MessageBox.Show($"{count}개의 뷰를 내보냈습니다.", "완료");
            ProgressChanged?.Invoke(100);
        }

        public string GetName() => "Export NWC Handler";

        private string GetSafeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}
