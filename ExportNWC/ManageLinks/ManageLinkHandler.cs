using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class ManageLinkHandler : IExternalEventHandler
    {
        public event Action LinkLoadCompleted;

        private List<RevitLinkInstances> _selectedLinks;

        public void SetData(List<RevitLinkInstances> selectedLinks)
        {
            _selectedLinks = selectedLinks;
        }

        public void Execute(UIApplication uiApp)
        {
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            if (_selectedLinks == null || _selectedLinks.Count == 0)
            {
                MessageBox.Show("선택된 링크가 없습니다.");
                LinkLoadCompleted?.Invoke();
                return;
            }

            foreach (var linkInstance in _selectedLinks)
            {
                if (linkInstance.Status == "Load")
                {
                    RevitLinkInstance instance = linkInstance.LinkInstance;
                    RevitLinkType linkType = doc.GetElement(instance.GetTypeId()) as RevitLinkType;

                    if (linkType == null)
                    {
                        MessageBox.Show("링크 타입을 찾을 수 없습니다.");
                        continue;
                    }

                    try
                    {
                        if (!RevitLinkType.IsLoaded(doc, linkType.Id))
                        {
                            var result = linkType.Load();
                        }
                        else
                        {
                            MessageBox.Show($"링크 {linkType.Name} 이미 로드됨");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"링크 {linkType.Name} 로드 중 예외 발생:\n{ex.Message}");
                    }
                }
            }
            LinkLoadCompleted?.Invoke();
        }


        public string GetName() => "Manage Link Handler";
    }
}
