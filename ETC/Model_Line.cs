using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.VisualBasic;  // InputBox 사용을 위해 필요

namespace NewJeans
{
    [Transaction(TransactionMode.Manual)]
    public class ModelLine : IExternalCommand
    {
        // mm 단위 → feet 단위로 변환
        public XYZ MmToXYZ(double xMm, double yMm, double zMm)
        {
            return new XYZ(
                UnitUtils.ConvertToInternalUnits(xMm, DisplayUnitType.DUT_MILLIMETERS),
                UnitUtils.ConvertToInternalUnits(yMm, DisplayUnitType.DUT_MILLIMETERS),
                UnitUtils.ConvertToInternalUnits(zMm, DisplayUnitType.DUT_MILLIMETERS)
            );
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 사용자로부터 선 개수 입력 받기
            string input = Interaction.InputBox("생성할 선 개수를 입력하세요", "선 개수 입력", "100");
            if (string.IsNullOrWhiteSpace(input))
            {
                TaskDialog.Show("입력 취소됨", "입력이 취소되었거나 비어 있습니다.");
                return Result.Cancelled;
            }

            int count = 0;
            if (!int.TryParse(input, out count) || count <= 0)
            {
                TaskDialog.Show("입력 오류", "양의 정수를 입력해주세요.");
                return Result.Failed;
            }

            // 시작, 끝점 설정
            XYZ start = MmToXYZ(0, 0, 0);
            XYZ end = MmToXYZ(1200, 0, 0);  // 1200mm 길이

            int batchSize = 500;
            SketchPlane sketchPlane = null;

            try
            {
                using (TransactionGroup tg = new TransactionGroup(doc, "Draw Many Lines"))
                {
                    tg.Start();

                    // SketchPlane을 생성하는 별도의 트랜잭션
                    using (Transaction txSketch = new Transaction(doc, "Create SketchPlane"))
                    {
                        txSketch.Start();

                        Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, start);
                        sketchPlane = SketchPlane.Create(doc, plane);

                        txSketch.Commit();
                    }

                    // 라인 생성 - 트랜잭션 나눠서 배치 처리
                    for (int i = 0; i < count; i += batchSize)
                    {
                        using (Transaction tx = new Transaction(doc, $"Batch {i}"))
                        {
                            tx.Start();

                            for (int j = 0; j < batchSize && (i + j) < count; j++)
                            {
                                int index = i + j;
                                XYZ offset = MmToXYZ(0, index * 100, 0);  // Y 방향 100mm 간격
                                Line line = Line.CreateBound(start + offset, end + offset);
                               
                                doc.Create.NewModelCurve(line, sketchPlane);
                            }

                            tx.Commit();
                        }
                    }

                    tg.Assimilate(); // 전체 트랜잭션 커밋
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("오류", ex.Message);
                return Result.Failed;
            }
        }
    }
}
