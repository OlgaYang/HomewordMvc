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
    public class 客戶聯絡人Controller : Controller
    {
        客戶聯絡人Repository repo;
        客戶資料Repository repo客戶;

        public 客戶聯絡人Controller()
        {
            repo = RepositoryHelper.Get客戶聯絡人Repository();
            repo客戶 = RepositoryHelper.Get客戶資料Repository(repo.UnitOfWork);
        }

        // GET: 客戶聯絡人
        public ActionResult Index(string sortOrder)
        {
            
            ViewBag.姓名SortParm = String.IsNullOrEmpty(sortOrder) ? "姓名_desc" : "";
            ViewBag.職稱SortParm = sortOrder == "職稱" ? "職稱_desc" : "職稱";
            ViewBag.EmailSortParm = sortOrder == "Email" ? "Email_desc" : "Email";
            ViewBag.手機SortParm = sortOrder == "手機" ? "手機_desc" : "手機";
            ViewBag.電話SortParm = sortOrder == "電話" ? "電話_desc" : "電話";
            ViewBag.客戶名稱SortParm = sortOrder == "客戶名稱" ? "客戶名稱_desc" : "客戶名稱";

            var data = repo.All();
            switch (sortOrder)
            {
                case "姓名_desc":
                    data = data.OrderByDescending(p => p.姓名);
                    break;
                case "職稱":
                    data = data.OrderBy(p => p.職稱);
                    break;
                case "職稱_desc":
                    data = data.OrderByDescending(p => p.職稱);
                    break;
                case "Email":
                    data = data.OrderBy(p => p.Email);
                    break;
                case "Email_desc":
                    data = data.OrderByDescending(p => p.Email);
                    break;
                case "電話":
                    data = data.OrderBy(p => p.電話);
                    break;
                case "電話_desc":
                    data = data.OrderByDescending(p => p.電話);
                    break;
                case "手機":
                    data = data.OrderBy(p => p.手機);
                    break;
                case "手機_desc":
                    data = data.OrderByDescending(p => p.手機);
                    break;
                case "客戶名稱":
                    data = data.OrderBy(p => p.客戶資料.客戶名稱);
                    break;
                case "客戶名稱_desc":
                    data = data.OrderByDescending(p => p.客戶資料.客戶名稱);
                    break;
                default:
                    data = data.OrderBy(p => p.姓名);
                    break;
            }


            ViewBag.職稱 = new SelectList(repo.All().Select(p => p.職稱).Distinct());            
            return View(data);
        }

        [HttpPost]
        public ActionResult Index(客戶聯絡人 客戶聯絡人)
        {
            var data = repo.All();
            if (!string.IsNullOrEmpty(客戶聯絡人.職稱))
            {
                data = data.Where(p => p.職稱 == 客戶聯絡人.職稱);
            }

            ViewBag.職稱 = new SelectList(repo.All().Select(p => p.職稱).Distinct());

            return View(data);
        }

        // GET: 客戶聯絡人/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo.Find(id);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // GET: 客戶聯絡人/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(repo客戶.All(), "Id", "客戶名稱");
            return View();
        }

        // POST: 客戶聯絡人/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話,是否已刪除")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                repo.Add(客戶聯絡人);
                repo.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(repo客戶.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: 客戶聯絡人/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo.Find(id);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(repo客戶.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // POST: 客戶聯絡人/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話,是否已刪除")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                repo.UnitOfWork.Context.Entry(客戶聯絡人).State = EntityState.Modified;
                repo.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(repo客戶.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: 客戶聯絡人/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo.Find(id);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // POST: 客戶聯絡人/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶聯絡人 客戶聯絡人 = repo.Find(id);
            repo.Delete(客戶聯絡人);
            repo.UnitOfWork.Commit();
            return RedirectToAction("Index");
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
            var sheet = workbook.Worksheets.Add("客戶聯絡人");

            int colIdx = 1;
            foreach (string colName in "職稱;姓名;Email;手機;電話".Split(';'))
            {
                sheet.Cell(1, colIdx++).Value = colName;
            }

            int rowIdx = 2;
            foreach (var item in list)
            {
                sheet.Cell(rowIdx, 1).Value = item.職稱;
                sheet.Cell(rowIdx, 2).Value = item.姓名;
                sheet.Cell(rowIdx, 3).Value = item.Email;
                sheet.Cell(rowIdx, 4).Value = item.手機;
                sheet.Cell(rowIdx, 5).Value = item.電話;
                rowIdx++;
            }

            var fs = new MemoryStream();
            workbook.SaveAs(fs);

            return File(fs.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"客戶聯絡人{DateTime.Now.ToString("yyyyMMddHHmmssfff") }.xlsx");
        }
    }
}
