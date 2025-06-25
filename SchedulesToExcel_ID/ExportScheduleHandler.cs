using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace NewJeans
{
    public class ExportScheduleHandler : IExternalEventHandler
    {
        public List<ViewSchedule> SelectedSchedules { get; set; }
        public string Exportpath { get; set; }

        public void SetData(List<ViewSchedule> schedules, string exportpath)
        {
            SelectedSchedules = schedules;
            Exportpath = exportpath;
        }


        public class ScheduleDocument
        {
            public string Title { get; set; }
            public List<ScheduleTable> Tables { get; set; }

            public ScheduleDocument(string title, List<ScheduleTable> tables)
            {
                Title = title;
                Tables = tables;
            }
        }
        public class ScheduleTable
        {
            public string Name { get; set; }
            public List<string> Headers { get; set; } = new List<string>();
            public List<List<string>> BodyRows { get; set; } = new List<List<string>>();
        }


        public event Action<int> ProgressChanged;
        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<ScheduleTable> tables = new List<ScheduleTable>();
            List<ScheduleDocument> scheduleDocuments = new List<ScheduleDocument>();

            int total = SelectedSchedules
                        .Select(s => s.GetTableData().GetSectionData(SectionType.Body).NumberOfRows - 2) // 2는 헤더 제외
                        .Sum();
            int current = 0;
            int lastShownProgress = -1;

            foreach (ViewSchedule schedule in SelectedSchedules)
            {
                ScheduleDefinition definition = schedule.Definition;

                var fields = definition.GetFieldOrder()
                    .Select(id=>definition.GetField(id))
                    .Where(f=>!f.IsHidden)
                    .ToList();

                var collector = new FilteredElementCollector(doc,schedule.Id)
                    .WhereElementIsNotElementType().ToList();

                IList<ScheduleSortGroupField> sortFields = definition.GetSortGroupFields();
                IEnumerable<Element> sortedElements = collector;


                var table = new ScheduleTable { Name = schedule.Name };

                table.Headers.Add("Element ID");
                foreach (var f in fields) table.Headers.Add(f.GetName());

                // 본문 세팅
                foreach (var elem in sortedElements)
                {
                    var row = new List<string> { elem.Id.IntegerValue.ToString() };
                    foreach (var f in fields)
                    {
                        var p = elem.LookupParameter(f.GetName());
                        row.Add(p != null ? (p.AsValueString() ?? p.AsString() ?? "") : "");
                    }
                    table.BodyRows.Add(row);

                    // 진행률
                    current++;
                    var progress = (int)(current / (double)total * 100);
                    if (progress != lastShownProgress)
                        ProgressChanged?.Invoke(lastShownProgress = progress);
                }

                tables.Add(table);
                scheduleDocuments.Add(
                    new ScheduleDocument(schedule.Name, new List<ScheduleTable> { table })
                );
            }

            try
            {
                ExcelExporter.Export(scheduleDocuments, Exportpath);
                if (MessageBox.Show("엑셀 저장 완료!\n저장된 파일을 여시겠습니까?", "성공", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        Process.Start(Exportpath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("엑셀을 여는 중 오류 발생:\n" + ex.Message, "오류");
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("엑셀 저장 오류", ex.ToString());
            }
        }

        public string GetName() => "ExportScheduleHandler";

    }
}
