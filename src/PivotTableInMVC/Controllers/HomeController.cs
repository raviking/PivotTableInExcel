using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PivotTableInMVC.Models;
using PivotTableInMVC.Utility;
using OfficeOpenXml;
using System.IO;

namespace PivotTableInMVC.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            List<TProduct> _listdata = new List<TProduct>();
            pivotdata dal = new pivotdata();
            _listdata = dal.PivotData();
            LoadExcelData(_listdata);
            return View();
        }
        public void LoadExcelData(List<TProduct> _lstproduct)
        {
            string filepath = Server.MapPath("~/Content/ProductReport.xlsx");
            System.Data.DataTable _dt = new System.Data.DataTable();
            try
            {
                if (System.IO.File.Exists(filepath))
                {
                    System.IO.File.Delete(filepath);
                }
                ExcelPackage app = new ExcelPackage();
                var sheet = app.Workbook.Worksheets.Add("ProductReport");
                sheet.Cells[1, 1].LoadFromCollection(_lstproduct, true);
                Stream stream = System.IO.File.Create(filepath);
                app.SaveAs(stream);
                stream.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }       
        public FileResult Export()
        {
            string finalpath = "";
            Export obj = new Utility.Export();
            finalpath= obj.OfficeDll();
            return File(finalpath, "application/vnd.ms-excel");
        }
    }
}