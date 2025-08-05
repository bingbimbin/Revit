using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Windows;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB.Architecture;


namespace NewJeans
{
    public class NozzleOrderHandler : IExternalEventHandler
    {
        public string ExportPath { get; set; }

        // ✅ 진행 상황 이벤트
        public event Action<int> ProgressChanged;

        public void SetData(string exportPath)
        {
            ExportPath = exportPath;
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                List<FamilyInstance> col = new FilteredElementCollector(doc)
                                                .OfClass(typeof(FamilyInstance))
                                                .Cast<FamilyInstance>()
                                                .Where(fi => fi.SuperComponent == null)
                                                .Where(fi => fi.MEPModel?.ConnectorManager != null)
                                                .ToList();

                List<(string familyName, string typeName, string result)> finalResults = new List<(string, string, string)>();

                int total = col.Count;
                int count = 0;

                foreach (FamilyInstance familyInstance in col)
                {
                    var connectorManager = familyInstance.MEPModel?.ConnectorManager;
                    if (connectorManager == null) continue;

                    var connectors = connectorManager
                                     .Connectors
                                     .Cast<Connector>()
                                     .ToList();

                    var sorted = connectors
                        .OrderBy(c => c.Origin.X)
                        .Select(c =>
                        {
                            string desc = c.Description ?? " ";
                            var m = Regex.Match(desc, @"(\d+)$");
                            int number = m.Success ? int.Parse(m.Groups[1].Value) : -1;
                            return (c, desc, number);
                        })
                        .ToList();

                    // 2) 축 방향 판별: BasisZ 벡터의 Z 성분 vs Y 성분
                    var sampleAxis = sorted[0].c.CoordinateSystem.BasisZ;
                    bool groupByY = Math.Abs(sampleAxis.Z) > Math.Abs(sampleAxis.Y);
                    // groupByY==true → Z방향 커넥터 → Y로 그룹화
                    // groupByY==false → Y방향 커넥터 → Z로 그룹화

                    Func<(Connector c, string desc, int number), double> getAxisCoord;

                    if (groupByY)
                    {
                        getAxisCoord = item => Math.Round(item.c.Origin.Y, 4);
                    }
                    else
                    {
                        getAxisCoord = item => Math.Round(item.c.Origin.Z, 4);
                    }
                    // 3) yGroups → axisGroups 로 이름만 바꾸고 getAxisCoord 사용
                    var axisGroups = sorted
                        .GroupBy(getAxisCoord)
                        .ToDictionary(g => g.Key, g => g.Count());

                    int maxCount = axisGroups.Values.Max();
                    var axisCandidates = axisGroups
                        .Where(kv => kv.Value == maxCount)
                        .Select(kv => kv.Key)
                        .ToList();

                    double dominantAxis;
                    if (axisCandidates.Count == 1)
                    {
                        dominantAxis = axisCandidates[0];
                    }
                    else
                    {
                        var numberedAxes = new HashSet<double>(
                            sorted
                            .Where(x => x.number != -1)
                            .Select(getAxisCoord)
                        );
                        var axisWithNumbers = axisCandidates
                            .Where(a => numberedAxes.Contains(a))
                            .ToList();

                        if (axisWithNumbers.Count == 1)
                            dominantAxis = axisWithNumbers[0];
                        else
                            dominantAxis = axisCandidates.First();
                    }

                    // 4) expectedSorted 계산에도 getAxisCoord 사용
                    var expectedSorted = sorted
                        .Where(x => x.number != -1
                                    && getAxisCoord(x) == dominantAxis)
                        .OrderBy(x => x.number)
                        .Select(x => x.number)
                        .ToList();

                    var numberedConns = sorted.Where(x => x.number != -1).ToList();
                    bool singleOne = numberedConns.Count == 1
                                     && numberedConns[0].number == 1;

                    int mainIndex = 0;
                    foreach (var item in sorted)
                    {
                        bool isMainLevel = getAxisCoord(item) == dominantAxis;
                        string suffix;

                        if (item.number == -1)
                        {
                            suffix = "";
                        }
                        else if (singleOne)
                        {
                            suffix = "_OO";
                        }
                        else if (!isMainLevel)
                        {
                            suffix = "_!!";
                        }
                        else
                        {
                            int expectedNumber = expectedSorted[mainIndex++];
                            suffix = (item.number == expectedNumber)
                                     ? "_OO"
                                     : "_XX";
                        }

                        string finalDescription = item.desc + suffix;
                        finalResults.Add((
                            familyInstance.Symbol.Family.Name,
                            familyInstance.Symbol.Name,
                            finalDescription));
                    }

                    count++;
                    int progress = (int)((count / (double)total) * 100);
                    ProgressChanged?.Invoke(progress);
                }



                try
                {
                    ExcelExporter.Export(finalResults, ExportPath);

                    if (MessageBox.Show("엑셀 저장 완료!\n저장된 파일을 여시겠습니까?", "성공", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        try
                        {
                            System.Diagnostics.Process.Start(ExportPath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("엑셀을 여는 중 오류 발생:\n" + ex.Message, "오류");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "엑셀 저장 오류");
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "오류");
            }
        }
        public string GetName() => "Extract Parameters Handler";
    }
}
