using ExcelDataReader;
using Schedule_ExcelReader_Core.Enums;
using Schedule_ExcelReader_Core.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Schedule_ExcelReader_Core.Controllers
{
    public static class Excel
    {
        public static Schedule ReadScheduleFromExcel(string fileDirectory)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Schedule schedule = new Schedule();

            if (fileDirectory == null || fileDirectory == "" || (!Path.GetFileName(fileDirectory).EndsWith(".xls") && !Path.GetFileName(fileDirectory).EndsWith(".xlsx")))
                return null;

            using (FileStream stream = File.Open($"{fileDirectory}", FileMode.Open, FileAccess.Read))
            {
                using (var edr = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet dataSet = edr.AsDataSet();
                    DataView dtView = dataSet.Tables[0].AsDataView();

                    schedule.Date = Date.ExtractFromString(dtView.Table.Rows[0].ItemArray[0].ToString());

                    Dictionary<ScheduleExcelColumn, int> keyValuePairs = GetKeysSchedule(dtView);

                    List<Group> group = new List<Group>();

                    for (int i = 3; i < dtView.Table.Rows.Count; i++)
                    {
                        if (dtView.Table.Rows[i].ItemArray[0].ToString().ToLower() == "номер пары") continue;
                        var groupName = dtView.Table.Rows[i].ItemArray[keyValuePairs.Where(p => p.Key == ScheduleExcelColumn.StudyGroup).Select(p => p.Value).FirstOrDefault()].ToString();
                        if (groupName.Length > 1)
                        {
                            group.Add(new Group() { NameGroup = groupName, Couples = new List<Couple>() });
                            continue;
                        }
                        int numCouple = 0;
                        if (!int.TryParse(dtView.Table.Rows[i].ItemArray[keyValuePairs.Where(p => p.Key == ScheduleExcelColumn.NumCouple).Select(p => p.Value).FirstOrDefault()].ToString(), out numCouple))
                            continue;
                        var discipline = dtView.Table.Rows[i].ItemArray[keyValuePairs.Where(p => p.Key == ScheduleExcelColumn.Discipline).Select(p => p.Value).FirstOrDefault()].ToString();
                        var auditorium = dtView.Table.Rows[i].ItemArray[keyValuePairs.Where(p => p.Key == ScheduleExcelColumn.Auditorium).Select(p => p.Value).FirstOrDefault()].ToString();
                        var teacher = dtView.Table.Rows[i].ItemArray[keyValuePairs.Where(p => p.Key == ScheduleExcelColumn.Teacher).Select(p => p.Value).FirstOrDefault()].ToString();

                        group[group.Count - 1].Couples.Add(new Couple()
                        {
                            Discipline = discipline,
                            Auditorium = auditorium,
                            Teacher = teacher,
                            NumCouple = numCouple
                        });
                    }
                    schedule.Groups = group;
                }

            }
            return schedule;

        }

        private static Dictionary<ScheduleExcelColumn, int> GetKeysSchedule(DataView dtView)
        {
            Dictionary<ScheduleExcelColumn, int> result = new Dictionary<ScheduleExcelColumn, int>();

            for (int i = 0; i < dtView.Table.Rows.Count; i++)
                {
                if (dtView.Table.Rows[i++].ItemArray[0].ToString().ToLower() == "учебная группа")
                {
                    for (int j = 0; j < dtView.Table.Rows[i].ItemArray.Count(); j++)
                    {
                        if (dtView.Table.Rows[i].ItemArray[j].ToString() != null && dtView.Table.Rows[i].ItemArray[j].ToString().Trim() != "")
                        {
                            switch (dtView.Table.Rows[i].ItemArray[j].ToString().ToLower())
                            {
                                case "номер пары":
                                    result.Add(ScheduleExcelColumn.NumCouple, j);
                                    result.Add(ScheduleExcelColumn.StudyGroup, j);
                                    break;
                                case "дисциплина":
                                    result.Add(ScheduleExcelColumn.Discipline, j);
                                    break;
                                case "преподаватель":
                                    result.Add(ScheduleExcelColumn.Teacher, j);
                                    break;
                                case "аудитория":
                                    result.Add(ScheduleExcelColumn.Auditorium, j);
                                    break;
                            }
                        }
                        
                    }
                    break;
                }
            }

            return result;
        }
    }
}
