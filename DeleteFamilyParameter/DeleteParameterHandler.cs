using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace NewJeans
{
    public class DeleteParameterHandler : IExternalEventHandler
    {
        public List<string> SelectedFamilyNames { get; set; }
        public string ParameterName { get; set; }

        public event Action<int> ProgressChanged;

        public void Execute(UIApplication app)
        {
            var parameterNamesToDelete = ParameterName
            .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(n => n.Trim())
            .Where(n => !string.IsNullOrEmpty(n))
            .ToList();

            if (parameterNamesToDelete.Count == 0)
            {
                MessageBox.Show("삭제할 파라미터 이름을 입력해주세요!", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Document doc = app.ActiveUIDocument.Document;

            List<Family> selectedfamilies = new FilteredElementCollector(doc)
                 .OfClass(typeof(Family)).Cast<Family>()
                 .Where(fi => SelectedFamilyNames.Contains(fi.Name)).ToList(); // 선택된 패밀리만 필터링



            int total = selectedfamilies.Count;
            int count = 0;

            var nonparamfamilies = new List<string>();
            var okfamilies = new List<string>();
            foreach (var family in selectedfamilies)
            {
                if (!family.IsEditable) continue;

                // ① 최초 한 번 열어서 캐시만 잡고 바로 닫기
                var famDoc = doc.EditFamily(family);
                famDoc.Close(false);

                // ② 캐시가 해제된 상태에서 다시 열기 — 이때부터 파라미터가 정상 로드됨
                famDoc = doc.EditFamily(family);
                var fm = famDoc.FamilyManager;

                using (var trans = new Transaction(famDoc, "패밀리 파라미터 삭제"))
                {
                    trans.Start();

                 
                    var allParams = fm.Parameters
                        .Cast<FamilyParameter>()
                        .ToList();

                    var toDelete = allParams
                        .Where(p => parameterNamesToDelete.Contains(p.Definition.Name))
                        .ToList();

                    if (toDelete.Count == 0) {nonparamfamilies.Add(family.Name);}
                    else {okfamilies.Add(family.Name);}

                    foreach (var p in toDelete)
                    {
                        fm.RemoveParameter(p);
                     }

                    trans.Commit();
                }

                famDoc.LoadFamily(doc, new FamilyLoadOptions());
                count++;
                int progress = (int)((count / (double)total) * 100);
                ProgressChanged?.Invoke(progress);
            }

            if (nonparamfamilies.Count == 0)
            {
                MessageBox.Show($"삭제 완료 패밀리 = {okfamilies.Count}개" 
                , "결과",
                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show($"삭제 완료 패밀리 = {okfamilies.Count}개\n\n" +
               $"파리미터 없는 패밀리\n{string.Join("\n", nonparamfamilies)}", "결과",
               MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }
        public string GetName() => "DeleteParameterHandler";
    }

    public class FamilyLoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            // 이미 프로젝트에 같은 패밀리가 있으면 덮어쓰기
            overwriteParameterValues = true;
            return true;
        }
        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse,
                                        out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Project;
            overwriteParameterValues = true;
            return true;
        }
    }
}
