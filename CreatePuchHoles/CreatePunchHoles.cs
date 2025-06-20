using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class CreatePunchHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            TaskDialog.Show("선택", "기준 Extrusion을 선택하세요");
            
            Reference pickedRef = uidoc.Selection.PickObject(ObjectType.Element, "Extrusion을 선택하세요.");
            Element elem = doc.GetElement(pickedRef.ElementId);

            double endOffsetMM = 0.0;

            // 3. Extrusion 타입인지 확인하고 캐스팅
            if (elem is Extrusion extrusion)
            {
                double endOffsetInch = extrusion.EndOffset;

                // 단위 변환 (Feet → mm)
                endOffsetMM = UnitUtils.ConvertFromInternalUnits(endOffsetInch, DisplayUnitType.DUT_MILLIMETERS);
            }
            else
            {
                TaskDialog.Show("오류", "선택한 요소는 Extrusion이 아닙니다.");
            }

            if (!doc.IsFamilyDocument)
            {
                TaskDialog.Show("환경 오류", "패밀리 파일(.rfa) 환경에서만 실행 가능합니다.");
                return Result.Failed;
            }

            double extrusionDepth = UnitUtils.ConvertToInternalUnits(150.0, DisplayUnitType.DUT_MILLIMETERS);
            string targetLayer = "Punch_hole";

            try
            {
                var importInst = new FilteredElementCollector(doc)
                    .OfClass(typeof(ImportInstance))
                    .Cast<ImportInstance>()
                    .FirstOrDefault();

                if (importInst == null)
                {
                    message = "Import된 CAD 요소를 찾을 수 없습니다.";
                    TaskDialog.Show("오류", message);
                    return Result.Failed;
                }

                var opts = new Options { ComputeReferences = true };
                GeometryElement geomElem = importInst.get_Geometry(opts);

                List<Curve> ellipseCurves = new List<Curve>();

                foreach (GeometryObject geoObj in geomElem)
                {
                    if (geoObj is GeometryInstance geomInst)
                    {
                        foreach (GeometryObject instObj in geomInst.GetInstanceGeometry())
                        {
                            if (instObj.GraphicsStyleId == ElementId.InvalidElementId)
                                continue;

                            var gStyle = doc.GetElement(instObj.GraphicsStyleId) as GraphicsStyle;
                            if (gStyle == null || !string.Equals(
                                    gStyle.GraphicsStyleCategory.Name,
                                    targetLayer,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;

                            if (instObj is Arc cadArc)
                            {
                                try
                                {
                                    XYZ center = cadArc.Center;
                                    double radius = cadArc.Radius;
                                    XYZ xAxis = cadArc.XDirection;
                                    XYZ yAxis = cadArc.YDirection;

                                    Curve ellipse = Ellipse.CreateCurve(
                                        center, radius, radius,
                                        xAxis, yAxis,
                                        0, 2 * Math.PI
                                    );
                                    ellipse.MakeBound(0, 2 * Math.PI);

                                    if (ellipse.IsBound)
                                    {
                                        ellipseCurves.Add(ellipse);
                                    }
                                }
                                catch (Exception arcEx)
                                {
                                    TaskDialog.Show("Arc 오류", arcEx.Message);
                                    continue;
                                }
                            }
                        }
                    }
                }

                if (!ellipseCurves.Any())
                {
                    TaskDialog.Show("결과 없음", "사용 가능한 원이 없습니다.");
                    return Result.Failed;
                }


                using (Transaction tx = new Transaction(doc, "Create Combined Void"))
                {
                    tx.Start();

                    foreach (Curve curve in ellipseCurves)
                    {
                        CurveArray curveArray = new CurveArray();
                        curveArray.Append(curve);

                        CurveArrArray singleProfile = new CurveArrArray();
                        singleProfile.Append(curveArray);

                        // 중심점 기준 Plane 정의
                        XYZ center = (curve.GetEndPoint(0) + curve.GetEndPoint(1)) / 2;
                        Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, center);
                        SketchPlane sketchPlane = SketchPlane.Create(doc, plane);

                        // 개별 Void Extrusion 생성
                        var voidExtrusion = doc.FamilyCreate.NewExtrusion(
                            isSolid: true,
                            profile: singleProfile,
                            sketchPlane: sketchPlane,
                            end: extrusionDepth);

                        // 이동 벡터 계산
                        double moveZMM = endOffsetMM - UnitUtils.ConvertFromInternalUnits(extrusionDepth, DisplayUnitType.DUT_MILLIMETERS);
                        double moveZ = UnitUtils.ConvertToInternalUnits(moveZMM, DisplayUnitType.DUT_MILLIMETERS);
                        XYZ moveVector = new XYZ(0, 0, moveZ);

                        ElementTransformUtils.MoveElement(doc, voidExtrusion.Id, moveVector);
                    }

                    TaskDialog.Show("완료", $"Extrusion 생성 완료");

                    tx.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("예외", ex.Message);
                return Result.Failed;
            }
        }
    }
}
