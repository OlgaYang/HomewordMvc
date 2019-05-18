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
    public class 客戶銀行資訊Controller : Controller
    {
        客戶銀行資訊Repository repo;
        客戶資料Repository repo客戶;

        public 客戶銀行資訊Controller()
        {
            repo = RepositoryHelper.Get客戶銀行資訊Repository();
            repo客戶 = RepositoryHelper.Get客戶資料Repository(repo.UnitOfWork);
        }

        // GET: 客戶銀行資訊
        public ActionResult Index()
        {
            var 客戶銀行資訊 = repo.All();
            return View(客戶銀行資訊.ToList());
        }

        // GET: 客戶銀行資訊/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo.Find(id);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(repo客戶.All(), "Id", "客戶名稱");
            return View();
        }

        // POST: 客戶銀行資訊/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼,是否已刪除")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                repo.Add(客戶銀行資訊);
                repo.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(repo客戶.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo.Find(id);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(repo客戶.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // POST: 客戶銀行資訊/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼,是否已刪除")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                repo.UnitOfWork.Context.Entry(客戶銀行資訊).State = EntityState.Modified;
                repo.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(repo客戶.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo.Find(id);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // POST: 客戶銀行資訊/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶銀行資訊 客戶銀行資訊 = repo.Find(id);
            repo.Delete(客戶銀行資訊);
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
            var sheet = workbook.Worksheets.Add("客戶銀行資訊");

            int colIdx = 1;
            foreach (string colName in "銀行名稱;銀行代碼;分行代碼;帳戶名稱;帳戶號碼".Split(';'))
            {
                sheet.Cell(1, colIdx++).Value = colName;
            }

            int rowIdx = 2;
            foreach (var item in list)
            {
                sheet.Cell(rowIdx, 1).Value = item.銀行名稱;
                sheet.Cell(rowIdx, 2).Value = item.銀行代碼;
                sheet.Cell(rowIdx, 3).Value = item.分行代碼;
                sheet.Cell(rowIdx, 4).Value = item.帳戶名稱;
                sheet.Cell(rowIdx, 5).Value = item.帳戶號碼;              

                rowIdx++;
            }

            var fs = new MemoryStream();
            workbook.SaveAs(fs);

            return File(fs.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"客戶銀行資訊{DateTime.Now.ToString("yyyyMMddHHmmssfff") }.xlsx");
        }
    }
}
