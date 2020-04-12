using API.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace API.MyContext
{
    public class myContext : DbContext
    {
        public myContext() : base("BelajarAPI") { }
        public DbSet<Department> Departments { get; set; }

        public DbSet<Division> Divisions { get; set; }

    }
}