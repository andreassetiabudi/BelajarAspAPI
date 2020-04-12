using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace API.ViewModel
{
    public class DivisionVM
    {
        public int Id { get; set; }
        public string DivisionName { get; set; }
        public bool IsDelete { get; set; }
        public DateTimeOffset CreateDate { get; set; }
        public DateTimeOffset UpdateDate { get; set; }
        public DateTimeOffset DeleteDate { get; set; }
        public int Department_Id { get; set; }
        public string DepartmentName { get; set; }
    }
}