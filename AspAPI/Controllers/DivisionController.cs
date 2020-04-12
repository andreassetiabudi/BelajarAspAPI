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
    public class DivisionController : Controller
    {
        readonly HttpClient client = new HttpClient
        {
            BaseAddress = new Uri("https://localhost:44375/API/")
        };

        // GET: Division
        public ActionResult Index()
        {
            return View(LoadDivision());
        }

        public JsonResult LoadDivision()
        {
            IEnumerable<DivisionVM> divisions = null;
            var responseTask = client.GetAsync("Division");
            responseTask.Wait();
            var result = responseTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var readTask = result.Content.ReadAsAsync<IList<DivisionVM>>();
                readTask.Wait();
                divisions = readTask.Result;
            }
            else
            {
                divisions = Enumerable.Empty<DivisionVM>();
                ModelState.AddModelError(string.Empty, "server error, try after some time");
            }
            return new JsonResult { Data = divisions, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }

        public JsonResult InsertOrUpdate(Division division)
        {
            var myContent = JsonConvert.SerializeObject(division);
            var buffer = System.Text.Encoding.UTF8.GetBytes(myContent);
            var byteContent = new ByteArrayContent(buffer);
            byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            if (division.Id == 0)
            {
                var result = client.PostAsync("Division", byteContent).Result;
                return new JsonResult { Data = result, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
            else
            {
                var result = client.PutAsync("Division/" + division.Id, byteContent).Result;
                return new JsonResult { Data = result, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
        }

        public async Task<JsonResult> GetById(int id)
        {
            HttpResponseMessage response = await client.GetAsync("Division");
            if (response.IsSuccessStatusCode)
            {
                var data = await response.Content.ReadAsAsync<IList<DivisionVM>>();
                var div = data.FirstOrDefault(S => S.Id == id);
                var json = JsonConvert.SerializeObject(div, Formatting.None, new JsonSerializerSettings()
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                });
                return new JsonResult { Data = json, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
            return Json("Internal Server Error");
        }

        public JsonResult Delete(int id)
        {
            var result = client.DeleteAsync("Division/" + id).Result;
            return new JsonResult { Data = result, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }
    }
}