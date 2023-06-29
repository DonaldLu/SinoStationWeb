using SinoStationWeb.Models;
using System.Collections.Generic;
using System.Web;
using System.Web.Mvc;

namespace SinoStationWeb.Controllers
{
    public class RegulatoryReviewController : Controller
    {
        private RegulatoryReviewService _service;

        public RegulatoryReviewController()
        {
            _service = new RegulatoryReviewService();
        }
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Upload()
        {
            return View();
        }
        public ActionResult Room()
        {
            return View();
        }
        public ActionResult Sheet1()
        {
            return PartialView();
        }
        public ActionResult Sheet2()
        {
            return PartialView();
        }
        public ActionResult Sheet3()
        {
            return PartialView();
        }

        // =============== Web API ================

        // 上傳Excel檔
        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file == null) return Json(new { Status = 0, Message = "No File Selected" });

            List<Room> roomList = _service.Upload(file);
            string names = string.Empty;
            foreach (Room room in roomList)
            {
                if (room.name != "")
                {
                    names += room.name + "、";
                }
            }
            var ret = names;
            return Json(roomList);
        }
        // 讀取所有規則
        public ActionResult AllRule()
        {
            List<RuleName> ret = _service.AllRule();
            return Json(ret);
        }
        // 讀取SQL資料
        [HttpPost]
        public ActionResult GetSQLData(string sqlName)
        {
            List<Room> ret = _service.GetSQLData(sqlName);
            return Json(ret);
        }
    }
}