using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using MvcHomeWork.Models;

namespace MvcHomeWork.Controllers
{
    public class V_客戶資料清單Controller : Controller
    {       
        V_客戶資料清單Repository repo = RepositoryHelper.GetV_客戶資料清單Repository();

        // GET: V_客戶資料清單
        public ActionResult Index()
        {
            return View(repo.All());
        }
     
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                repo.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult ExportExcel()
        {
            var list = repo.All();

            //建立Excel
            XLWorkbook workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("客戶資料清單");

            int colIdx = 1;
            foreach (string colName in "客戶名稱;聯絡人數量;銀行帳戶數量".Split(';'))
            {
                sheet.Cell(1, colIdx++).Value = colName;
            }

            int rowIdx = 2;
            foreach (var item in list)
            {
                sheet.Cell(rowIdx, 1).Value = item.客戶名稱;
                sheet.Cell(rowIdx, 2).Value = item.聯絡人數量;
                sheet.Cell(rowIdx, 3).Value = item.銀行帳戶數量;

                rowIdx++;
            }

            var fs = new MemoryStream();
            workbook.SaveAs(fs);

            return File(fs.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"客戶資料清單{DateTime.Now.ToString("yyyyMMddHHmmssfff") }.xlsx");
        }
    }
}
