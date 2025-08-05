using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Windows;
using System.Linq;

namespace NewJeans
{
    public class DuplicateHandler : IExternalEventHandler
    {
        public string ExportPath { get; set; }
        public event Action<int> ProgressChanged;

        public void SetData(string exportPath)
        {
            ExportPath = exportPath;
        }

        public class ConnectorBaseInfo
        {
            public FamilyInstance FamilyInstance { get; set; }
            public List<XYZ> ConnectorLocations { get; set; }
            public List<string> ConnectorDirections { get; set; }
            public ConnectorBaseInfo(FamilyInstance familyInstance, List<XYZ> connectorLocations, List<string> connectorDirections)
            {
                FamilyInstance = familyInstance;
                ConnectorLocations = connectorLocations;
                ConnectorDirections = connectorDirections;
            }
        }

        public class OutPut
        {
            public string FamilyName { get; set; }
            public string TypeName { get; set; }
            public string Result { get; set; }
            public OutPut(string familyName, string typeName, string result)
            {
                FamilyName = familyName;
                TypeName = typeName;
                Result = result;
            }
        }

        public void Execute(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                List<FamilyInstance> col = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList();
                List<ConnectorBaseInfo> connectorBaseInfo = new List<ConnectorBaseInfo>();
                List<OutPut> final = new List<OutPut>();

                foreach (FamilyInstance familyInstance in col)
                {
                    var connectorManager = familyInstance.MEPModel?.ConnectorManager;

                    if (connectorManager != null)
                    {
                        var connectorLocations = new List<XYZ>();
                        var connectorDirections = new List<string>();

                        foreach (Connector connector in connectorManager.Connectors)
                        {
                            connectorLocations.Add(connector.Origin);
                            XYZ dir = connector.CoordinateSystem.BasisZ.Normalize();

                            if (dir.X > 0.9) connectorDirections.Add("X+");
                            else if (dir.X < -0.9) connectorDirections.Add("X-");
                            else if (dir.Y > 0.9) connectorDirections.Add("Y+");
                            else if (dir.Y < -0.9) connectorDirections.Add("Y-");
                            else if (dir.Z > 0.9) connectorDirections.Add("Z+");
                            else if (dir.Z < -0.9) connectorDirections.Add("Z-");
                            else connectorDirections.Add("UNKNOWN");
                        }

                        connectorBaseInfo.Add(new ConnectorBaseInfo(familyInstance, connectorLocations, connectorDirections));
                    }
                }

                int total = connectorBaseInfo.Count;
                int count = 0;

                foreach (var fam in connectorBaseInfo)
                {
                    string familyName = fam.FamilyInstance.Symbol.Family.Name;
                    string typeName = fam.FamilyInstance.Symbol.Name;

                    var xyzSet = new HashSet<Tuple<double, double, double>>();
                    bool hasXYZDuplicate = false;

                    var directionGroups = new Dictionary<string, List<XYZ>>();

                    for (int i = 0; i < fam.ConnectorLocations.Count; i++)
                    {
                        XYZ loc = fam.ConnectorLocations[i];
                        string dir = fam.ConnectorDirections[i];

                        double x = Math.Round(loc.X, 4);
                        double y = Math.Round(loc.Y, 4);
                        double z = Math.Round(loc.Z, 4);

                        var key3D = Tuple.Create(x, y, z);
                        if (!xyzSet.Add(key3D))
                        {
                            hasXYZDuplicate = true;
                            continue;
                        }

                        if (!directionGroups.ContainsKey(dir))
                            directionGroups[dir] = new List<XYZ>();

                        directionGroups[dir].Add(loc);
                    }

                    bool hasPlaneDuplicate = false;

                    foreach (var kvp in directionGroups)
                    {
                        string dir = kvp.Key;
                        var locs = kvp.Value;

                        if (locs.Count <= 1) continue;

                        var set2D = new HashSet<Tuple<double, double>>();

                        if (dir == "Z+" || dir == "Z-")
                        {
                            foreach (var loc in locs)
                            {
                                var key = Tuple.Create(Math.Round(loc.X, 4), Math.Round(loc.Y, 4));
                                if (!set2D.Add(key))
                                {
                                    hasPlaneDuplicate = true;
                                    break;
                                }
                            }
                        }
                        else if (dir == "Y+" || dir == "Y-")
                        {
                            foreach (var loc in locs)
                            {
                                var key = Tuple.Create(Math.Round(loc.X, 4), Math.Round(loc.Z, 4));
                                if (!set2D.Add(key))
                                {
                                    hasPlaneDuplicate = true;
                                    break;
                                }
                            }
                        }
                        else if (dir == "X+" || dir == "X-")
                        {
                            foreach (var loc in locs)
                            {
                                var key = Tuple.Create(Math.Round(loc.Y, 4), Math.Round(loc.Z, 4));
                                if (!set2D.Add(key))
                                {
                                    hasPlaneDuplicate = true;
                                    break;
                                }
                            }
                        }
                    }

                    string result;
                    if (hasXYZDuplicate && hasPlaneDuplicate) result = "중복 & 확인필요";
                    else if (hasXYZDuplicate) result = "중복";
                    else if (hasPlaneDuplicate) result = "확인필요";
                    else result = "정상";

                    final.Add(new OutPut(familyName, typeName, result));

                    count++;
                    int progress = (int)((count / (double)total) * 100);
                    ProgressChanged?.Invoke(progress);
                }

                try
                {
                    ExcelExporter.Export(final, ExportPath);

                    if (MessageBox.Show("엑셀 저장 완료!\n저장된 파일을 여시겠습니까?", "성공", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        Process.Start(ExportPath);
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
