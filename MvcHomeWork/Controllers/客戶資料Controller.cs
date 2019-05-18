using ClosedXML.Excel;
using MvcHomeWork.Models;
using System;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Mvc;

namespace MvcHomeWork.Controllers
{
    public class 客戶資料Controller : Controller
    {
        private 客戶資料Repository repo = RepositoryHelper.Get客戶資料Repository();

        // GET: 客戶資料
        public ActionResult Index(string sortOrder)
        {
            ViewBag.客戶名稱SortParm = String.IsNullOrEmpty(sortOrder) ? "name_desc" : "";
            ViewBag.統一編號SortParm = sortOrder == "統一編號" ? "統一編號_desc" : "統一編號";
            ViewBag.客戶分類SortParm = sortOrder == "客戶分類" ? "客戶分類_desc" : "客戶分類";
            ViewBag.電話SortParm = sortOrder == "電話" ? "電話_desc" : "電話";
            ViewBag.傳真SortParm = sortOrder == "傳真" ? "傳真_desc" : "傳真";
            ViewBag.地址SortParm = sortOrder == "地址" ? "地址_desc" : "地址";
            ViewBag.EmailSortParm = sortOrder == "Email" ? "Email_desc" : "Email";

            var data = repo.All();
            switch (sortOrder)
            {
                case "name_desc":
                    data = data.OrderByDescending(p => p.客戶名稱);
                    break;
                case "統一編號":
                    data = data.OrderBy(p => p.統一編號);
                    break;
                case "統一編號_desc":
                    data = data.OrderByDescending(p => p.統一編號);
                    break;
                case "客戶分類":
                    data = data.OrderBy(p => p.客戶分類);
                    break;
                case "客戶分類_desc":
                    data = data.OrderByDescending(p => p.客戶分類);
                    break;
                case "電話":
                    data = data.OrderBy(p => p.電話);
                    break;
                case "電話_desc":
                    data = data.OrderByDescending(p => p.電話);
                    break;
                case "傳真":
                    data = data.OrderBy(p => p.傳真);
                    break;
                case "傳真_desc":
                    data = data.OrderByDescending(p => p.傳真);
                    break;
                case "地址":
                    data = data.OrderBy(p => p.地址);
                    break;
                case "地址_desc":
                    data = data.OrderByDescending(p => p.地址);
                    break;
                case "Email":
                    data = data.OrderBy(p => p.Email);
                    break;
                case "Email_desc":
                    data = data.OrderByDescending(p => p.Email);
                    break;
                default:
                    data = data.OrderBy(p => p.客戶名稱);
                    break;
            }

            ViewBag.客戶分類 = new SelectList(repo.All().Select(p => p.客戶分類).Distinct());
            return View(data);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(客戶資料 客戶資料, string sortOrder)
        {
            var data = repo.All();

            if (!string.IsNullOrEmpty(客戶資料.客戶分類))
            {
                data = data.Where(p => p.客戶分類 == 客戶資料.客戶分類);
            }
            ViewBag.客戶分類 = new SelectList(repo.All().Select(p => p.客戶分類).Distinct());
            return View(data);
        }

        // GET: 客戶資料/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = repo.Find(id);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // GET: 客戶資料/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: 客戶資料/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶名稱,客戶分類,統一編號,電話,傳真,地址,Email,是否已刪除")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                repo.Add(客戶資料);
                repo.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            return View(客戶資料);
        }

        // GET: 客戶資料/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = repo.Find(id);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // POST: 客戶資料/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶名稱,客戶分類,統一編號,電話,傳真,地址,Email,是否已刪除")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                repo.UnitOfWork.Context.Entry(客戶資料).State = EntityState.Modified;
                repo.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            return View(客戶資料);
        }

        // GET: 客戶資料/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = repo.Find(id);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // POST: 客戶資料/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶資料 客戶資料 = repo.Find(id);
            repo.Delete(客戶資料);
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
            var sheet = workbook.Worksheets.Add("客戶資料");

            int colIdx = 1;
            foreach (string colName in "客戶名稱;客戶分類;統一編號;電話;傳真;地址;Email".Split(';'))
            {
                sheet.Cell(1, colIdx++).Value = colName;
            }

            int rowIdx = 2;
            foreach (var item in list)
            {
                sheet.Cell(rowIdx, 1).Value = item.客戶名稱;
                sheet.Cell(rowIdx, 2).Value = item.客戶分類;
                sheet.Cell(rowIdx, 3).Value = item.統一編號;
                sheet.Cell(rowIdx, 4).Value = item.電話;
                sheet.Cell(rowIdx, 5).Value = item.傳真;
                sheet.Cell(rowIdx, 6).Value = item.地址;
                sheet.Cell(rowIdx, 7).Value = item.Email;

                rowIdx++;
            }

            var fs = new MemoryStream();
            workbook.SaveAs(fs);

            return File(fs.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"客戶資料{DateTime.Now.ToString("yyyyMMddHHmmssfff") }.xlsx");
        }
    }
}