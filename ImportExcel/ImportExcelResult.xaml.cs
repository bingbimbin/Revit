using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using System.Windows.Threading;

namespace NewJeans
{
    public class ParameterItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        private bool isSelected = false;

        public bool IsSelected
        {
            get => isSelected;
            set
            {
                if (isSelected != value)
                {
                    isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public partial class ImportExcelResult : UserControl
    {
        // 파라미터 리스트 바인딩용
        public ObservableCollection<ParameterItem> ParameterFields { get; set; } = new ObservableCollection<ParameterItem>();

        private ExternalEvent _externalEvent;
        private ImportExcelHandler _handler;

        public ImportExcelResult(ExternalEvent externalEvent, ImportExcelHandler handler)
        {
            InitializeComponent();
            _externalEvent = externalEvent;
            _handler = handler;
            chkFieldList.ItemsSource = ParameterFields;

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
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel Files|*.xlsx;*.xls";
            dlg.Title = "엑셀 파일 선택";

            if (dlg.ShowDialog() == true)
            {
                txtFilePath.Text = dlg.FileName;

                // 엑셀에서 헤더 읽어오기
                string[] headers = ReadExcelHeaders(dlg.FileName);

                // 파라미터 필드 리스트에 반영
                SetParameterFields(headers);
            }
        }


        public void SetParameterFields(string[] headers)
        {
            ParameterFields.Clear();

            // 예) "ElementId"는 수정 불가하니 제외
            foreach (var header in headers.Where(h => h != "ElementId"))
            {
                ParameterFields.Add(new ParameterItem { Name = header });
            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            var selectedFields = GetSelectedFields();

            if (selectedFields.Length == 0)
            {
                MessageBox.Show("수정할 필드를 최소 하나 이상 선택하세요.", "경고", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _handler.SetData(txtFilePath.Text, selectedFields.ToList());
            _externalEvent.Raise();
        }

        public string[] GetSelectedFields()
        {
            return ParameterFields.Where(p => p.IsSelected).Select(p => p.Name).ToArray();
        }


        // 전체 선택 체크박스 체크했을 때
        private void chkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var item in ParameterFields)
            {
                item.IsSelected = true;
            }
        }

        // 전체 선택 체크박스 체크 해제했을 때
        private void chkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var item in ParameterFields)
            {
                item.IsSelected = false;
            }
        }


        private string[] ReadExcelHeaders(string filePath)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var ws = package.Workbook.Worksheets.First();
                    int colCount = ws.Dimension.Columns;
                    var headers = new List<string>();

                    for (int col = 1; col <= colCount; col++)
                    {
                        string header = ws.Cells[1, col].Value?.ToString().Trim();

                        if (!string.IsNullOrEmpty(header))
                        {
                            headers.Add(header);
                        }
                    }

                    return headers.ToArray();
                }
            }
            catch
            {
                MessageBox.Show("엑셀 헤더를 읽는 중 오류가 발생했습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                return new string[0];
            }
        }

    }
}
