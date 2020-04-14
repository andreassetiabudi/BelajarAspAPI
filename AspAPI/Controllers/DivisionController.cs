using API.Models;
using API.ViewModel;
using AspAPI.Report;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace AspAPI.Controllers
{
    public class DivisionController : Controller
    {
        readonly HttpClient client = new HttpClient()
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
            IEnumerable<DivisionVM> datadivision = null;
            var responseTask = client.GetAsync("Division");
            responseTask.Wait();
            var result = responseTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var readTask = result.Content.ReadAsAsync<IList<DivisionVM>>();
                readTask.Wait();
                datadivision = readTask.Result;
            }
            else
            {
                datadivision = Enumerable.Empty<DivisionVM>();
                ModelState.AddModelError(string.Empty, "Sorry Server Error, Try Again");
            }

            return new JsonResult { Data = datadivision, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }

        public JsonResult InsertorUpdate(DivisionVM division)
        {
            var myContent = JsonConvert.SerializeObject(division);
            var buffer = System.Text.Encoding.UTF8.GetBytes(myContent);
            var byteContent = new ByteArrayContent(buffer);
            byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            if (division.Id == 0)
            {
                var result = client.PostAsync("Division", byteContent).Result;
                return new JsonResult { Data = result, JsonRequestBehavior = JsonRequestBehavior.AllowGet }; //Return if Success
            }
            else
            {
                var result = client.PutAsync("Division/" + division.Id, byteContent).Result;
                return new JsonResult { Data = result, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
        }

        public async Task<JsonResult> GetById(int Id)
        {
            HttpResponseMessage response = await client.GetAsync("Division");
            if (response.IsSuccessStatusCode)
            {
                var data = await response.Content.ReadAsAsync<IList<DivisionVM>>();
                var dept = data.FirstOrDefault(t => t.Id == Id);
                var json = JsonConvert.SerializeObject(dept, Formatting.None, new JsonSerializerSettings()
                {
                    ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore
                });
                return new JsonResult { Data = json, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
            return Json("Internal Server Error"); // return if Error
        }

        public JsonResult Delete(int Id)
        {
            var Result = client.DeleteAsync("Division/" + Id).Result;
            return new JsonResult { Data = Result, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }

        //Export
        //Export
        public ActionResult ExportToPDF(DivisionVM division)
        {
            DivisionReport divreport = new DivisionReport();
            byte[] abytes = divreport.PrepareReport(GetDivision());
            return File(abytes, "application/pdf");
        }

        public List<DivisionVM> GetDivision()
        {
            IEnumerable<DivisionVM> datadivision = null;
            var responseTask = client.GetAsync("Division");
            responseTask.Wait();
            var result = responseTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var readTask = result.Content.ReadAsAsync<IList<DivisionVM>>();
                readTask.Wait();
                datadivision = readTask.Result;
            }
            else
            {
                datadivision = Enumerable.Empty<DivisionVM>();
                ModelState.AddModelError(string.Empty, "Sorry Server Error, Try Again");
            }

            return datadivision.ToList();
        }

        public ActionResult ExportToExcel()
        {
            var comlumHeaders = new string[]
            {
                "Id",
                "Name Division",
                "Name Department",
                "Tanggal Dibuat",
                "Tanggal Diubah"
            };

            byte[] result;

            using (var package = new ExcelPackage())
            {
                // add a new worksheet to the empty workbook

                var worksheet = package.Workbook.Worksheets.Add("Division List"); //Worksheet name
                using (var cells = worksheet.Cells[1, 1, 1, 4]) //(1,1) (1,5)
                {
                    cells.Style.Font.Bold = true;
                }

                //First add the headers
                for (var i = 0; i < comlumHeaders.Count(); i++)
                {
                    worksheet.Cells[1, i + 1].Value = comlumHeaders[i];
                }

                //Add values
                var j = 2;
                foreach (var divn in GetDivision())
                {
                    worksheet.Cells["A" + j].Value = divn.Id;
                    worksheet.Cells["B" + j].Value = divn.DivisionName;
                    worksheet.Cells["C" + j].Value = divn.DepartmentName;
                    worksheet.Cells["D" + j].Value = divn.CreateDate.ToString();
                    worksheet.Cells["E" + j].Value = divn.UpdateDate.ToString();
                    j++;
                }
                result = package.GetAsByteArray();
            }

            return File(result, "application/ms-excel", $"DivisionList.xlsx");
        }

        public ActionResult ExportToCSV()
        {
            var comlumHeaders = new string[]
            {
                "Id",
                "Name Division",
                "Name Department",
                "Tanggal Dibuat",
                "Tanggal Diubah"
            };

            var deptRecords = (from divn in GetDivision()
                               select new object[]
                               {
                                            divn.Id,
                                            divn.DivisionName,
                                            divn.DepartmentName,
                                            divn.CreateDate.ToString(),
                                            divn.UpdateDate.ToString()
                               }).ToList();

            // Build the file content
            var deptcsv = new StringBuilder();
            deptRecords.ForEach(line =>
            {
                deptcsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.ASCII.GetBytes($"{string.Join(",", comlumHeaders)}\r\n{deptcsv.ToString()}");
            return File(buffer, "text/csv", $"DivisionList.csv");
        }
    }
}