using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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
                var table = new ScheduleTable { Name = schedule.Name };

                TableData tabledata = schedule.GetTableData();
                TableSectionData bodyData = tabledata.GetSectionData(SectionType.Body);
                
                int colCount = bodyData.NumberOfColumns;
                int rowCount = bodyData.NumberOfRows;

                // 헤더부분
                List<string> headers = new List<string>();

                for (int i = 0; i < colCount; i++)
                {
                    string header = schedule.GetCellText(SectionType.Body,0,i);
                    headers.Add(header);
                }
                table.Headers = headers;

                // 본문 부분

                for (int row = 2; row < bodyData.NumberOfRows; row++)
                {
                    List<string> rowData = new List<string>();

                    for (int col = 0; col < colCount; col++)
                    {
                        string value = schedule.GetCellText(SectionType.Body, row, col);
                        rowData.Add(value);
                    }
                    table.BodyRows.Add(rowData);

                    current++;
                    int progress = (int)((double)current / total * 100);
                    ProgressChanged?.Invoke(progress);
                }
                tables.Add(table);
                scheduleDocuments.Add(new ScheduleDocument(schedule.Name, new List<ScheduleTable> { table }));

                

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
                TaskDialog.Show("엑셀 저장 오류", ex.Message);
            }
        }

        public string GetName() => "ExportScheduleHandler";

    }
}
