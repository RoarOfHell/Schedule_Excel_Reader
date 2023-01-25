using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Schedule_ExcelReader_Core.Models
{
    public class Schedule
    {
        public List<Group> Groups { get; set; }
        public DateTime Date { get; set; }
    }
}
