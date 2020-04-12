using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace API.Models
{
    [Table("TB_M_Division")]
    public class Division
    {
        //constructor
        public Division()
        {
        }


        [Key]
        public int Id { get; set; }
        public string DivisionName { get; set; }
        public bool IsDelete { get; set; }

        public DateTimeOffset CreateDate { get; set; }

        public DateTimeOffset UpdateDate { get; set; }

        public DateTimeOffset DeleteDate { get; set; }

        [ForeignKey("Department")]
        public int Department_Id { get; set; } //manual virtual foreign key

        public Department Department { get; set; } //automatic foreign key


        public Division(Division division) //create
        {
            this.DivisionName = division.DivisionName;
            this.IsDelete = false;
            this.CreateDate = DateTimeOffset.Now.ToLocalTime();
            this.Department = division.Department;
        }

        public void Update(Division division) //update
        {
            this.DivisionName = division.DivisionName;
            this.Department = division.Department;
            this.UpdateDate = DateTimeOffset.Now.ToLocalTime();
        }

        public void Delete()
        {
            this.IsDelete = true;
            this.DeleteDate = DateTimeOffset.Now.ToLocalTime();
        }





    }
}