using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Autodesk.Revit.DB;
using WinForms = System.Windows.Forms;

namespace NewJeans
{
    public class View3DItem
    {
        public string Name { get; set; }
        public bool IsSelected { get; set; }
        public View3D View { get; set; }
    }

    public partial class ExportNWCResult : UserControl
    {
        private ExternalEvent _externalEvent;
        private ExportNWCHandler _handler;

        public List<string> SelectedViews { get; set; }
        public string ExportPath { get; set; }

        public ExportNWCResult()
        {
            InitializeComponent();
        }   


        public void Initialize(ExternalEvent externalEvent, ExportNWCHandler handler, Document doc)
        {
            _externalEvent = externalEvent;
            _handler = handler;

            _handler.ProgressChanged += OnProgressChanged;

            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate && v.CanBePrinted)
                .OrderBy(v => v.Name)
                .ToList();


            foreach (var view in views)
            {
                ViewList.Items.Add(new View3DItem
                {
                    View = view,
                    Name = view.Name,
                    IsSelected = false
                });
            }

        }

        private void OnProgressChanged(int progress)
        {
            Dispatcher.Invoke(() =>
            {
                progressBar.Visibility = System.Windows.Visibility.Visible;
                progressBar.Value = progress;

                DoEvents();

                if (progress >= 100)
                    progressBar.Visibility = System.Windows.Visibility.Collapsed;
            });
        }

        private void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Render, new DispatcherOperationCallback(delegate (object parameter)
            {
                frame.Continue = false;
                return null;
            }), null);
            Dispatcher.PushFrame(frame);
        }

        private void chkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            foreach (View3DItem item in ViewList.Items)
                item.IsSelected = true;

            ViewList.Items.Refresh();
        }

        private void chkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (View3DItem item in ViewList.Items)
                item.IsSelected = false;    

            ViewList.Items.Refresh();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "NWC 파일을 저장할 폴더를 선택하세요.";

                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    ExportPath = dialog.SelectedPath;
                    txtFilePath.Text = ExportPath;
                }
            }
        }


        private void ViewList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ((ListBox)sender).SelectedItem = null;
        }

        public List<string> GetSelectedViews()
        {
            return ViewList.Items
                .OfType<View3DItem>()
                .Where(v => v.IsSelected)
                .Select(v => v.Name)
                .ToList();
        }

        public List<View3D> GetSelectedViewObjects()
        {
            return ViewList.Items
                .OfType<View3DItem>()
                .Where(v => v.IsSelected)
                .Select(v => v.View)
                .ToList();
        }

        public string GetExportPath()
        {
            return txtFilePath.Text;
        }

        public bool GetOverwriteOption()
        {
            return chkOverwrite.IsChecked == true;
        }

    }
}
