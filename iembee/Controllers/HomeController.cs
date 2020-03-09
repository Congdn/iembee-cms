using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;

namespace iembee.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult ExportFile(Data model)
        {
            var filePath = "";
            //Save upload file
            if (model.fileInput.ContentLength > 0)
            {
                var filename = Path.GetFileName(model.fileInput.FileName);
                filePath = Path.Combine(Server.MapPath("~/App_Data/uploads"), filename);
                model.fileInput.SaveAs(filePath);
            }

            //Export data


            //var body = request.Content.ReadAsByteArrayAsync();
            return Content("test");
        }

    }
    public class Data
    {
        [Required(ErrorMessage ="File does not exist")]
        public HttpPostedFileBase fileInput { get; set; }
        [Required]
        public int hangtoithieu { get; set; }
        [Required]
        public int hangtoida { get; set; }
        public string tenkh { get; set; }
        public string diachi { get; set; }
        public string dienthoai { get; set; }
        public int? tongnhap3 { get; set; }
        public int? tongnhap2 { get; set; }
        public int? tongnhap1 { get; set; }
        public int? tongxuat3 { get; set; }
        public int? tongxuat2 { get; set; }
        public int? tongxuat1 { get; set; }
    }
}