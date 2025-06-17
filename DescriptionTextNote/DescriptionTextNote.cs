
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using static System.Net.Mime.MediaTypeNames;
namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class DescriptionText : IExternalCommand
    {
        public class TextInfo
        {
            public string Description { get; set; }
            public string Diameter { get; set; }
            public XYZ Origin { get; set; }

            public TextInfo(string description, string diameter, XYZ origin)
            {
                Description = description;
                Diameter = diameter;
                Origin = origin;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            View activeView = doc.ActiveView;

            var connectors = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_ConnectorElem)
                    .Cast<ConnectorElement>().ToList();

            TextNoteType textNoteType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault();

            List<TextInfo> textInfos = new List<TextInfo>();

            foreach (ConnectorElement e in connectors)
            {
                Parameter paramDescription = e.get_Parameter(BuiltInParameter.RBS_CONNECTOR_DESCRIPTION);
                Parameter paramDiameter = e.get_Parameter(BuiltInParameter.CONNECTOR_DIAMETER);

                string description = paramDescription.AsString();
                string diameter = paramDiameter.AsValueString()+"A";
                XYZ origin = e.Origin;

                if (!string.IsNullOrEmpty(description))
                {
                    textInfos.Add(new TextInfo(description, diameter, origin));
                }

            }

            TextNoteOptions options = new TextNoteOptions
            {
                HorizontalAlignment = HorizontalTextAlignment.Center,
                VerticalAlignment = VerticalTextAlignment.Top,
                TypeId = textNoteType.Id
            };

            using (Transaction trans = new Transaction(doc, "Text 생성"))
            {
                trans.Start();

                for (int i = 0; i < textInfos.Count; i++)
                {
                    TextInfo textInfo = textInfos[i];
                    XYZ origin = textInfo.Origin;
                    string description = textInfo.Description;
                    string diameter = textInfo.Diameter;
                    TextNote textNote = TextNote.Create(doc, activeView.Id, origin, $"{description}\n{diameter}", options);
                }
                trans.Commit();
            }

            return Result.Succeeded;

        }
    }
}
