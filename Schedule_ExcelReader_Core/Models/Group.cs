using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Schedule_ExcelReader_Core.Models
{
    public class Group
    {
        public string NameGroup { get; set; }
        public List<Couple> Couples { get; set; }
    }
}
