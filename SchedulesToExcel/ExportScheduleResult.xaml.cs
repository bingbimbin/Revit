using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Threading.Tasks;
using System.Windows.Threading;
using Autodesk.Revit.DB;

namespace NewJeans
{
    public partial class ExportScheduleResult : UserControl
    {
        public List<string> SelectedSchedules { get; set; }
        public string exportPath { get; set; }

        private ExternalEvent _externalEvent;
        private ExportScheduleHandler _handler;


        public ExportScheduleResult(ExternalEvent externalEvent, ExportScheduleHandler handler, Document doc)
        {
            InitializeComponent();

            this.Loaded += (s, e) =>
            {
                var win = Window.GetWindow(this);
                if (win != null)
                {
                    win.Topmost = true;
                }
            };

            _externalEvent = externalEvent;
            _handler = handler;

            _handler.ProgressChanged += OnProgressChanged;

            var schedules = new FilteredElementCollector(doc)
                            .OfClass(typeof(ViewSchedule))
                            .Cast<ViewSchedule>()
                            .Where(vs => !vs.IsTemplate)
                            .OrderBy(vs => vs.Name)
                            .ToList();

            foreach (var schedule in schedules)
            {
                lstFiles.Items.Add(new ListBoxItem
                {
                    Content = schedule.Name,
                    Tag = schedule
                });
            }
        }


        private void OnProgressChanged(int progress)
        {
            Dispatcher.Invoke(() =>
            {
                progressBar.Visibility = System.Windows.Visibility.Visible;
                progressBar.Value = progress;

                DoEvents(); // ✅ UI 업데이트 강제

                if (progress >= 100)
                {
                    progressBar.Visibility = System.Windows.Visibility.Collapsed;
                }
            });
        }

        // ✅ UI 강제 업데이트 함수 추가
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

        private void btnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            lstFiles.SelectAll();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == true)
            {
                string ExportPath = saveFileDialog.FileName;
                txtFilePath.Text = ExportPath;
                exportPath = ExportPath;
            }
        }
        private void txtFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {

            progressBar.Visibility = System.Windows.Visibility.Visible;
            progressBar.IsIndeterminate = false;
            progressBar.Value = 0;

            List<ViewSchedule> SelectedSchedules = new List<ViewSchedule>();
            foreach (ListBoxItem item in lstFiles.SelectedItems)
            {
                if (item.Tag is ViewSchedule schedule)
                {
                    SelectedSchedules.Add(schedule);
                }
            }
            _handler.SetData(SelectedSchedules, exportPath);
            _externalEvent.Raise();

        }
    }
}
