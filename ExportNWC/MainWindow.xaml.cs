using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace NewJeans
{
    public partial class MainUserControl : UserControl
    {
        private UIApplication _uiApp;
        private ExternalEvent _linkEvent;
        private ManageLinkHandler _linkHandler;

        private ExternalEvent _exportEvent;
        private ExportNWCHandler _exportHandler;

        private ManageLinkResult _linkControl;
        private ExportNWCResult _exportControl;

        public MainUserControl(UIApplication uiApp, 
                                ExternalEvent exportEvent,
                                ExportNWCHandler exportHandler,
                                ExternalEvent linkEvent,
                                ManageLinkHandler linkHandler)
        {
            InitializeComponent();
            _uiApp = uiApp;

            Document doc = _uiApp.ActiveUIDocument.Document;

            _exportHandler = exportHandler;
            _exportEvent = exportEvent;
            _linkHandler = linkHandler;
            _linkEvent = linkEvent;

            _linkHandler.LinkLoadCompleted += OnLinkLoadCompleted;

            // 컨트롤 생성 및 초기화
            _exportControl = new ExportNWCResult();
            _exportControl.Initialize(_exportEvent, _exportHandler, doc);

            _linkControl = new ManageLinkResult();
            _linkControl.Initialize(_linkEvent, _linkHandler, doc);

            tabProjectFamily.Content = _exportControl;
            tabRfaFamily.Content = _linkControl;
        }


        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            var selectedLinks = _linkControl.GetSelectedLinks();
            _linkHandler.SetData(selectedLinks);
            _linkEvent.Raise();
        }

        private void OnLinkLoadCompleted()
        {
            var selectedViews = _exportControl.GetSelectedViewObjects();
            var exportPath = _exportControl.GetExportPath();
            var overwrite = _exportControl.GetOverwriteOption();

            _exportHandler.SetData(selectedViews, exportPath, overwrite);
            _exportEvent.Raise();
        }

    }
}
