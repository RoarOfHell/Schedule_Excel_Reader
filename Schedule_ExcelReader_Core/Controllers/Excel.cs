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

namespace Schedule_ExcelReader_Core.Controllers;
public class Excel
{
    private readonly string fileDirectory;
    private DataTable dataTable;
    private bool checkRead;
    private Dictionary<ScheduleExcelColumn, int> scheduleExcelColumns;

    public Excel(string fileDirectory)
    {
        if (fileDirectory is null || fileDirectory == "") throw new Exception("Empty filename");

        this.fileDirectory = fileDirectory;
        checkRead = false;
    }

    private string GetData(int i, ScheduleExcelColumn exelColumn)
    {
        return dataTable.Rows[i].ItemArray[scheduleExcelColumns[exelColumn]].ToString();
    }
    private Dictionary<ScheduleExcelColumn, int> GetScheduleColumn()
    {
        Dictionary<ScheduleExcelColumn, int> result = new();

        for (int i = 0; i < dataTable.Rows.Count; i++)
        {
            if (dataTable.Rows[i++].ItemArray[0].ToString().ToLower() != "учебная группа") continue;

            for (int j = 0; j < dataTable.Rows[i].ItemArray.Length; j++)
            {
                string element = dataTable.Rows[i].ItemArray[j].ToString();
                if (element is null || element.Trim() == "") continue;

                ScheduleExcelColumn addItem = element.ToLower() switch
                {
                    "дисциплина" => ScheduleExcelColumn.Discipline,
                    "преподаватель" => ScheduleExcelColumn.Teacher,
                    "аудитория" => ScheduleExcelColumn.Auditorium,
                    _ => ScheduleExcelColumn.None // Без этого предуприжения даются если с ним буде тне работать убери
                };
                if (element.ToLower() == "номер пары")
                {
                    result.Add(ScheduleExcelColumn.NumCouple, j);
                    result.Add(ScheduleExcelColumn.StudyGroup, j);
                    continue;
                }

                result.Add(addItem, j);
            }
            break;
        }

        return result;
    }
    public void ReadExel()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        FileInfo file = new(fileDirectory);

        if (!file.Exists && !file.Name.EndsWith(".xls") && !file.Name.EndsWith(".xlsx")) throw new Exception("File does not exist");

        using FileStream stream = File.Open(fileDirectory, FileMode.Open, FileAccess.Read);
        using var edr = ExcelReaderFactory.CreateReader(stream);

        dataTable = edr
            .AsDataSet()
            .Tables[0]
            .AsDataView()
            .Table;

        scheduleExcelColumns = GetScheduleColumn();

        checkRead = true;
    }
    public Schedule ToSchedule()
    {
        if (!checkRead) throw new Exception("You dont read file");

        Schedule schedule = new()
        {
            Date = Date.ExtractFromString(dataTable.Rows[0].ItemArray[0].ToString())
        };
        DataRowCollection dataRowCollection = dataTable.Rows;

        List<Group> group = new();

        for (int i = 3; i < dataRowCollection.Count; i++)
        {
            if (dataRowCollection[i].ItemArray[0].ToString().ToLower() != "номер пары") continue;
            var groupName = GetData(i, ScheduleExcelColumn.StudyGroup);
            if (groupName.Length > 1)
            {
                group.Add(new Group() { NameGroup = groupName, Couples = new List<Couple>() });
                continue;
            }

            if (!int.TryParse(GetData(i, ScheduleExcelColumn.NumCouple), out int numCouple)) continue;
            group[^1].Couples.Add(new()
            {
                Discipline = GetData(i, ScheduleExcelColumn.Discipline),
                Auditorium = GetData(i, ScheduleExcelColumn.Auditorium),
                Teacher = GetData(i, ScheduleExcelColumn.Teacher),
                NumCouple = numCouple
            });
        }
        schedule.Groups = group;

        return schedule;
    }
}
