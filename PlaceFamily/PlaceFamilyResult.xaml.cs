using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace NewJeans
{
    public partial class PlaceFamilyResult : UserControl
    {
        private ExternalEvent _externalEvent;
        private PlaceFamilyHandler _handler;
        private Document _doc;


        public PlaceFamilyResult(ExternalEvent externalEvent, PlaceFamilyHandler handler, Document doc)
        {
            InitializeComponent();
            _externalEvent = externalEvent;
            _handler = handler;
            _doc = doc;


            _handler.ProgressChanged += OnProgressChanged;

            var modelCategories = new List<Category>();

            foreach (Category cat in _doc.Settings.Categories)
            {
                if (cat.CategoryType != CategoryType.Model)
                    continue;

                if (Enum.IsDefined(typeof(BuiltInCategory), cat.Id.IntegerValue))
                {
                    modelCategories.Add(cat);
                }
            }

            var sortedCategories = modelCategories.OrderBy(cat => cat.Name).ToList();

            foreach (Category cat in sortedCategories)
            {
                ListBoxItem item = new ListBoxItem
                {
                    Content = cat.Name
                };
                lstCategories.Items.Add(item);
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


        private void btnPlace_Click(object sender, RoutedEventArgs e)
        {
            progressBar.Visibility = System.Windows.Visibility.Visible;
            progressBar.IsIndeterminate = false;
            progressBar.Value = 0;
            List<BuiltInCategory> selectedCats = new List<BuiltInCategory>();

            foreach (ListBoxItem item in lstCategories.SelectedItems)
            {
                string categoryName = item.Content.ToString();


                foreach (Category cat in _doc.Settings.Categories)
                {
                    if (cat.Name == categoryName && Enum.IsDefined(typeof(BuiltInCategory), cat.Id.IntegerValue))
                    {
                        BuiltInCategory bic = (BuiltInCategory)cat.Id.IntegerValue;
                        selectedCats.Add(bic);
                        break;
                    }
                }
            }

            if (selectedCats.Count == 0)
            {
                MessageBox.Show("카테고리를 선택하세요!");
                return;
            }

            if (!double.TryParse(xspan.Text, out double spacingX_mm))
            {
                MessageBox.Show("가로 간격이 잘못되었습니다!");
                return;
            }

            if (!double.TryParse(yspan.Text, out double spacingY_mm))
            {
                MessageBox.Show("세로 간격이 잘못되었습니다!");
                return;
            }

            _handler.SpacingX = UnitUtils.ConvertToInternalUnits(spacingX_mm, DisplayUnitType.DUT_MILLIMETERS);
            _handler.SpacingY = UnitUtils.ConvertToInternalUnits(spacingY_mm, DisplayUnitType.DUT_MILLIMETERS);

            _handler.SetCategories(selectedCats);
            _externalEvent.Raise();
        }
    }
}
