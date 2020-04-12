using API.Models;
using API.ViewModel;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace AspAPI.Controllers
{
    public class DepartmentController : Controller
    {
        readonly HttpClient client = new HttpClient
        {
            BaseAddress = new Uri("https://localhost:44375/API/")
        };

        // GET: Department
        public ActionResult Index()
        {
            return View(LoadDepartment());
        }

        public JsonResult LoadDepartment()
        {
            IEnumerable<Department> departments = null;
            var responseTask = client.GetAsync("Department");
            responseTask.Wait();
            var result = responseTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var readTask = result.Content.ReadAsAsync<IList<Department>>();
                readTask.Wait();
                departments = readTask.Result;
            }
            else
            {
                departments = Enumerable.Empty<Department>();
                ModelState.AddModelError(string.Empty, "server error, try after some time");
            }
            return new JsonResult { Data = departments, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }

        public JsonResult InsertOrUpdate(Department department)
        {
            var myContent = JsonConvert.SerializeObject(department);
            var buffer = System.Text.Encoding.UTF8.GetBytes(myContent);
            var byteContent = new ByteArrayContent(buffer);
            byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            if (department.Id == 0) //insert
            {
                var result = client.PostAsync("Department", byteContent).Result;
                return new JsonResult { Data = result, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
            else //update
            {
                var result = client.PutAsync("Department/"+ department.Id, byteContent).Result;
                return new JsonResult { Data = result, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
        }

        public async Task<JsonResult> GetById(int id)
        {
            HttpResponseMessage response = await client.GetAsync("Department");
            if (response.IsSuccessStatusCode)
            {
                var data = await response.Content.ReadAsAsync<IList<Department>>();
                var dept = data.FirstOrDefault(S => S.Id == id);
                var json = JsonConvert.SerializeObject(dept, Formatting.None, new JsonSerializerSettings()
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                });
                return new JsonResult { Data = json, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
            return Json("Internal Server Error");
        }

        public JsonResult Delete(int id)
        {
            var result = client.DeleteAsync("Department/" +id).Result;
            return new JsonResult { Data = result, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }
    }
}