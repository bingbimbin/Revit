using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;

namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class MakeConnectorsOnPlane : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                TaskDialog.Show("시작", "기준 plane 선택하세요.");

                Reference pickedRef = uidoc.Selection.PickObject(ObjectType.Face, "기준 평면이 될 면을 선택하세요.");
                Element pickedElem = doc.GetElement(pickedRef.ElementId);
                Face pickedFace = pickedElem.GetGeometryObjectFromReference(pickedRef) as Face;

                if (pickedFace == null)
                {
                    TaskDialog.Show("오류", "선택한 면에서 Face를 가져올 수 없습니다.");
                    return Result.Failed;
                }

                UV uv = new UV(0.5, 0.5);
                XYZ pickedOrigin = pickedFace.Evaluate(uv);
                XYZ pickedNormal = pickedFace is PlanarFace pf
                    ? pf.FaceNormal.Normalize()
                    : pickedFace.ComputeNormal(uv);


                Options opt = new Options { ComputeReferences = true };

                using (Transaction tx = new Transaction(doc, "같은 평면 위 Connector 생성"))
                {
                    tx.Start();

                    FilteredElementCollector collector = new FilteredElementCollector(doc)
                        .OfClass(typeof(Extrusion))
                        .WhereElementIsNotElementType();

                    foreach (Element elem in collector)
                    {
                        GeometryElement geomElem = elem.get_Geometry(opt);
                        if (geomElem == null) continue;

                        foreach (GeometryObject obj in geomElem)
                        {
                            Solid solid = obj as Solid;
                            if (solid == null || solid.Faces.Size == 0) continue;

                            foreach (Face face in solid.Faces)
                            {

                                if (!(face is PlanarFace targetPf))
                                    continue;

                                if (!IsOnSameInfinitePlane(targetPf, pickedNormal, pickedOrigin))
                                    continue;
                                                                
                                
                                double? diameter = TryGetCircularFaceDiameter(targetPf);
                                if (diameter.HasValue)
                                {
                                    ConnectorElement connector = ConnectorElement.CreatePipeConnector(
                                        doc,
                                        PipeSystemType.Global,
                                        targetPf.Reference);

                                    double condiameter = diameter.Value;
                                    Parameter param = connector.LookupParameter("Diameter");
                                    param.Set(condiameter);
                                }

                            }
                        }
                    }

                    tx.Commit();
                }

                return Result.Succeeded;
            }

            catch (Exception ex)
            {
                TaskDialog.Show("오류", ex.Message);
                return Result.Failed;
            }
        }

        private bool IsOnSameInfinitePlane(PlanarFace pf, XYZ refNormal, XYZ refPoint)
        {
            XYZ normal = pf.FaceNormal.Normalize();
            if (!normal.IsAlmostEqualTo(refNormal, 1e-6))
                return false;

            double offset = (pf.Origin - refPoint).DotProduct(refNormal);
            return Math.Abs(offset) < 1e-6;
        }

        private double? TryGetCircularFaceDiameter(PlanarFace pf)
        {
            foreach (EdgeArray edgeLoop in pf.EdgeLoops)
            {
                foreach (Edge edge in edgeLoop)
                {
                    Curve curve = edge.AsCurve();               
                    if (curve is Arc arc)
                    {
                        return 2 * arc.Radius;
                    }
                }
            }
            return null;
        }
    }
}
