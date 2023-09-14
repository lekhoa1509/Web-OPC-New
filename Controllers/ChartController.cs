using ASPNET_MVC_ChartsDemo.Models;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Web.Mvc;

namespace ASPNET_MVC_ChartsDemo.Controllers
{
    public class ChartController : Controller
    {
        // GET: Home    
        public ActionResult Chart()
        {
            List<Chart> dataPoint = new List<Chart>();

            dataPoint.Add(new Chart("Samsung", 25));
            dataPoint.Add(new Chart("Micromax", 13));
            dataPoint.Add(new Chart("Lenovo", 8));
            dataPoint.Add(new Chart("Intex", 7));
            dataPoint.Add(new Chart("Reliance", (int)6.8));
            dataPoint.Add(new Chart("Others", (int)40.2));


            ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoint);

            return View();
        }
        public ActionResult DoanhThuChiNhanh_Admin()
        {
            List<Chart> dataPoint = new List<Chart>();
                
            dataPoint.Add(new Chart("Samsung", 25));
            dataPoint.Add(new Chart("Micromax", 13));
            dataPoint.Add(new Chart("Lenovo", 8));
            dataPoint.Add(new Chart("Intex", 7));
            dataPoint.Add(new Chart("Reliance", (int)6.8));
            dataPoint.Add(new Chart("Others",(int) 40.2));

            ViewBag.DataPoints = JsonConvert.SerializeObject(dataPoint);

            return View("DoanhThuChiNhanh_Admin");
        }
    }
}