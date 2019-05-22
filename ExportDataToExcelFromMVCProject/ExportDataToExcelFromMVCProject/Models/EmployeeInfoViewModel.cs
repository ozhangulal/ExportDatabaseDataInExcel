using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDataToExcelFromMVCProject.Models
{
    public class EmployeeInfoViewModel
    {
        public int EmployeeID { get; set; }
        public string EmployeeName { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public int? Experience { get; set; } //Db içerisinde Null olabilecek bir sütun ile eşleşeceğinden ötürü int tipinin yanına '?' koyduk.
    }

}
