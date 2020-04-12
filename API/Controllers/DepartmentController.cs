using API.Models;
using API.Repository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;

namespace API.Controllers
{
    
    public class DepartmentController : ApiController
    {
        DepartmentRepository department = new DepartmentRepository();


        // GET API/Department
        [HttpGet]
        public IEnumerable<Department> Get()
        {
            return department.Get();
        }


        // POST API/Department
        public IHttpActionResult Post(Department departments)
        {
            if ((departments.DepartmentName != null) && (departments.DepartmentName != ""))
            {
                department.Create(departments);
                return Ok("Department Insert Succesfully!"); 
            }
            return BadRequest("Failed to Insert Department");
        }


        // GETBYID API/Department
        [ResponseType(typeof(Department))]
        public async Task<IEnumerable<Department>> Get(int Id)
        {
            return await department.Get(Id);
        }


        // UPDATE API/Department
        public IHttpActionResult Put (int Id, Department departments)
        {
            if ((departments.DepartmentName != null) && (departments.DepartmentName != ""))
            {
                department.Update(Id, departments);
                return Ok("Department Updated Succesfully!");
            }
            return BadRequest("Failed to Update Department");
        }

        //DELETE API/Department
        public IHttpActionResult Delete(int Id)
        {
            var delete = department.Delete(Id);
            if (delete > 0)
            {
                return Ok("Department Deleted Succesfully!");
            }
            return BadRequest("Failed to Deleted Department");
        }



    }
}
