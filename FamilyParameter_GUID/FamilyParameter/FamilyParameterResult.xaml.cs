using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace NewJeans
{
    public partial class FamilyParameterResult : UserControl
    {
        public string SelectedFilePath { get; private set; }

        private ExternalEvent _externalEvent;
        private FamilyParameterHandler _handler;

        public FamilyParameterResult()
        {
            InitializeComponent();
        }

        public void Initialize(ExternalEvent externalEvent, FamilyParameterHandler handler)
        {
            _externalEvent = externalEvent;
            _handler = handler;
            _handler.ProgressChanged += OnProgressChanged;
        }

        private void OnProgressChanged(int progress)
        {
            Dispatcher.Invoke(() =>
            {
                progressBar.Visibility = Visibility.Visible;
                progressBar.Value = progress;

                DoEvents(); // ✅ UI 업데이트 강제

                if (progress >= 100)
                {
                    progressBar.Visibility = Visibility.Collapsed;
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



        private void btnSavePath_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == true)
            {
                SelectedFilePath = saveFileDialog.FileName;
                txtFilePath.Text = SelectedFilePath;
            }
        }


        private async void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SelectedFilePath))
            {
                MessageBox.Show("파일 경로를 먼저 설정해주세요.");
                return;
            }


            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = false;
            progressBar.Value = 0;

            // 실제 핸들러 실행
            _handler.SetData(SelectedFilePath);

            _externalEvent.Raise();

        }
    }
}
