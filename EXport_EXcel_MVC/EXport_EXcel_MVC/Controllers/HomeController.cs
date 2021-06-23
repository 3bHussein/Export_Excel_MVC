using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using EXport_EXcel_MVC.Models;

namespace EXport_EXcel_MVC.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            testEntities db = new testEntities();
            //NorthwindEntities entities = new NorthwindEntities();
            //return View(from customer in entities.Customers.Take(10)
            //            select customer);

            return View(from t in db.tbl_registration.Take(10) select t);
        }

        [HttpPost]
        public FileResult Export()
        {
            testEntities db = new testEntities();
            DataTable dt = new DataTable("Sheet1");
            dt.Columns.AddRange(new DataColumn[5] { new DataColumn("Email"),
                                            new DataColumn("Password"),
                                            new DataColumn("Name"),
                                            new DataColumn("Address"),
                                            new DataColumn("City") });

            var customers = from t in db.tbl_registration.Take(10)
                            select t;

            foreach (var f in customers)
            {
                //dt.Rows.Add(customer.,customer.CompanyName, customer.ContactName, customer.City, customer.Country);
                dt.Rows.Add(f.Email, f.Password, f.Name, f.Address, f.City);

            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Data.xlsx");
                }
            }
        }
    }
}









