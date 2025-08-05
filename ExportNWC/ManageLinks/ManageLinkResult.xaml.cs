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
using WinForms = System.Windows.Forms;

namespace NewJeans
{
    public class RevitLinkInstances
    {
        public string Name { get; set; }
        public bool IsSelected { get; set; }
        public string Status { get; set; } = "Unload";
        public RevitLinkInstance LinkInstance { get; set; }
    }

    public partial class ManageLinkResult : UserControl
    {
        private ExternalEvent _externalEvent;
        private ManageLinkHandler _handler;

        public List<string> SelectedViews { get; set; }
        public string ExportPath { get; set; }

        public ManageLinkResult()
        {
            InitializeComponent();
        }

        public void Initialize(ExternalEvent externalEvent, ManageLinkHandler handler, Document doc)
        {

            _externalEvent = externalEvent;
            _handler = handler;

            var Links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .OrderBy(l => l.Name)
                .ToList();


            foreach (var link in Links)
            {
                RevitLinkType linkType = doc.GetElement(link.GetTypeId()) as RevitLinkType;

                RevitLinkList.Items.Add(new RevitLinkInstances
                {
                    LinkInstance = link,
                    Name = linkType?.Name,
                    IsSelected = false
                });
            }

        }


        private void chkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            foreach (RevitLinkInstances item in RevitLinkList.Items)
                item.IsSelected = true;

            RevitLinkList.Items.Refresh();
        }

        private void chkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (RevitLinkInstances item in RevitLinkList.Items)
                item.IsSelected = false;

            RevitLinkList.Items.Refresh();
        }


        private void RevitLinkList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ((ListBox)sender).SelectedItem = null;
        }


        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            foreach (RevitLinkInstances item in RevitLinkList.Items)
            {
                if (item.IsSelected)
                {
                    item.Status = "Load";
                }
            }
            RevitLinkList.Items.Refresh();

            var selectedToLoad = RevitLinkList.Items
                                .Cast<RevitLinkInstances>()
                                .Where(x => x.Status == "Load")
                                .ToList();

            _handler.SetData(selectedToLoad);
        }


        public List<RevitLinkInstances> GetSelectedLinks()
        {
            return RevitLinkList.Items
                .OfType<RevitLinkInstances>()
                .Where(item => item.IsSelected && item.Status == "Load")
                .ToList();
        }

        private void btnUnload_Click(object sender, RoutedEventArgs e)
        {
            foreach (RevitLinkInstances item in RevitLinkList.Items)
            {
                if (item.IsSelected)
                {
                    item.Status = "Unload";
                }
            }
            RevitLinkList.Items.Refresh();
        }

    }
}
