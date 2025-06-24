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
    public partial class FamilyParameterResult2 : UserControl
    {
        public List<string> SelectedFiles { get; private set; } = new List<string>();
        public string ExportPath { get; private set; }

        private ExternalEvent _externalEvent;
        private FamilyParameterHandler2 _handler;

        public FamilyParameterResult2()
        {
            InitializeComponent();
        }

        public void Initialize(ExternalEvent externalEvent, FamilyParameterHandler2 handler)
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

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Revit 패밀리 파일 (*.rfa)|*.rfa";

            if (openFileDialog.ShowDialog() == true)
            {
                // 새로 선택된 파일 중 기존에 없는 파일만 추가
                var newFiles = openFileDialog.FileNames.Except(SelectedFiles).ToList();
                SelectedFiles.AddRange(newFiles);

                // 리스트 박스 초기화 후 전체 다시 출력
                lstFiles.Items.Clear();
                foreach (var file in SelectedFiles)
                {
                    lstFiles.Items.Add(System.IO.Path.GetFileName(file));
                }
            }
        }

        private void btnBrowse_Click2(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == true)
            {
                ExportPath = saveFileDialog.FileName;
                txtFilePath.Text = ExportPath;
            }
        }

        private void btnRemoveFile_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lstFiles.SelectedItems.Cast<string>().ToList();

            foreach (var item in selectedItems)
            {
                SelectedFiles.RemoveAll(f => System.IO.Path.GetFileName(f) == item);
                lstFiles.Items.Remove(item);
            }
        }

        private async void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ExportPath))
            {
                MessageBox.Show("파일 경로를 먼저 설정해주세요.");
                return;
            }

            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show("Revit 파일을 최소 1개 선택해주세요.");
                return;
            }

            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = false;
            progressBar.Value = 0;

            // 실제 핸들러 실행
            _handler.SetData(SelectedFiles, ExportPath);
            _externalEvent.Raise();
        }
    }
}
