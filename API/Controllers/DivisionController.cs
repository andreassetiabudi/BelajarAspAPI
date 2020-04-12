using API.Models;
using API.Repository;
using API.ViewModel;
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
    public class DivisionController : ApiController
    {
        DivisionRepository division = new DivisionRepository();


        // GET API/Division
        [HttpGet]
        public IEnumerable<DivisionVM> Get()
        {
            return division.Get();
        }


        // POST API/Division
        public IHttpActionResult Post(Division divisions)
        {
            if ((divisions.DivisionName != null) && (divisions.DivisionName != ""))
            {
                division.Create(divisions);
                return Ok("Division Insert Succesfully!");
            }
            return BadRequest("Failed to Insert Division");
        }


        // GETBYID API/Division
        [ResponseType(typeof(DivisionVM))]
        public async Task<IEnumerable<DivisionVM>> GetById(int Id)
        {
            if (await division.Get(Id) == null)
            {
                return null;
            }
            return await division.Get(Id);
        }


        // UPDATE API/Division
        public IHttpActionResult Put(int Id, Division divisions)
        {
            if ((divisions.DivisionName != null) && (divisions.DivisionName != ""))
            {
                division.Update(Id, divisions);
                return Ok("Division Updated Succesfully!");
            }
            return BadRequest("Failed to Update Division");
        }

        //DELETE API/Division
        public IHttpActionResult Delete(int Id)
        {
            var delete = division.Delete(Id);
            if (delete > 0)
            {
                return Ok("Division Deleted Succesfully!");
            }
            return BadRequest("Failed to Deleted Division");
        }

    }
}
