using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace NewJeans
{
    public partial class DeleteParameterResult : UserControl
    {
        public List<string> FamilyNames { get; set; }

        private ExternalEvent _loadEvent;
        private ExternalEvent _deleteEvent;

        private LoadFamiliesHandler _loadHandler;
        private DeleteParameterHandler _deleteHandler;

        public DeleteParameterResult(ExternalEvent loadEvent, LoadFamiliesHandler loadHandler,
                              ExternalEvent deleteEvent, DeleteParameterHandler deleteHandler)
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

            _loadEvent = loadEvent;
            _deleteEvent = deleteEvent;

            _loadHandler = loadHandler;
            _deleteHandler = deleteHandler;

            _loadHandler.Callback = OnFamilyListCollected;

            LoadFamilyList();
            // ✅ 진행도 이벤트 연결
            _deleteHandler.ProgressChanged += OnProgressChanged;
        }

        private async void LoadFamilyList()
        {
            await Task.Run(() =>
            {
                _loadEvent.Raise();  // 외부 이벤트를 자동으로 발생시킴
            });
        }

        // 콜백 메서드: FamilyNames를 ListBox에 업데이트
        private void OnFamilyListCollected(List<string> familyNames)
        {
            // UI 스레드에서 ListBox를 업데이트
            Dispatcher.Invoke(() =>
            {
                lstFiles.Items.Clear();  // 기존 항목 지우기
                foreach (var name in familyNames)
                {
                    lstFiles.Items.Add(name);  // 패밀리 이름을 ListBox에 추가
                }
            });
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
        private void btnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            lstFiles.SelectAll();
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            // 사용자가 선택한 패밀리 이름들 가져오기
            List<string> selectedFamilyNames = lstFiles.SelectedItems.Cast<string>().ToList();


            if (selectedFamilyNames.Count == 0)
            {
                MessageBox.Show("삭제할 패밀리를 선택해주세요!", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string parameterName = txtParameterName.Text;
            if (string.IsNullOrWhiteSpace(parameterName))
            {
                MessageBox.Show("삭제할 파라미터 이름을 입력해주세요!", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = false;
            progressBar.Value = 0;

            _deleteHandler.SelectedFamilyNames = selectedFamilyNames;
            _deleteHandler.ParameterName = parameterName;

            _deleteEvent.Raise();

        }
    }


}
