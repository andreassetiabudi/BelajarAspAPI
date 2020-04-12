using API.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace API.Models
{
    [Table("TB_M_Department")]
    public class Department
    {

        //constructor
        public Department()
        {
        }

        [Key]
        public int Id { get; set; }
        public string DepartmentName { get; set; }
        public bool IsDelete { get; set; }

        public DateTimeOffset CreateDate { get; set; }

        public DateTimeOffset UpdateDate { get; set; }

        public DateTimeOffset DeleteDate { get; set; }



        public Department(Department department) //create
        {
            this.DepartmentName = department.DepartmentName;
            this.CreateDate = DateTimeOffset.Now.ToLocalTime();
            this.IsDelete = false;
        }

        public void Update(Department department) //update
        {
            this.DepartmentName = department.DepartmentName;
            this.UpdateDate = DateTimeOffset.Now.ToLocalTime();
        }

        public void Delete () //delete
        {
            this.IsDelete = true;
            this.DeleteDate = DateTimeOffset.Now;
        }

    }
}