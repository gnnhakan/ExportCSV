using ExportCSV.Models;
using ExportCSV.Models.DataViewModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;

namespace ExportCSV.Controllers
{
    public class HomeController : Controller
    {
       
        public ActionResult Index()
        {
            return View();
        }

        public void ExportCSVFile()
        {
            StringWriter sw = new StringWriter();

            sw.WriteLine("\"ContactName\",\"Address\",\"ContactTitle\",\"Country\"");
            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment;filename=Exported_Users.csv");
            Response.ContentType = "text/csv";

            using (var mycontext = new xyzEntities())
            {
                foreach (var cst in mycontext.Customers)
                {
                    sw.WriteLine(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\"",
                                           cst.ContactName,
                                           cst.Address,
                                           cst.ContactTitle,
                                           cst.Country));
                }
            }

            Response.Write(sw.ToString());
            Response.End();
        }

        public void ExportToExcel()
        {
            using (var mycontext = new xyzEntities())
            {


                var grid = new System.Web.UI.WebControls.GridView();

                grid.DataSource = from d in mycontext.Customers
                                  select new
                                  {
                                      Address = d.Address,
                                      CompanyName = d.CompanyName,
                                      Country = d.Country,
                                      ContactName = d.ContactName

                                  };

                grid.DataBind();

                Response.ClearContent();
                Response.AddHeader("content-disposition", "attachment; filename=Exported_Diners.xls");
                Response.ContentType = "application/excel";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);

                grid.RenderControl(htw);

                Response.Write(sw.ToString());

                Response.End();
            }
        }
    }
}