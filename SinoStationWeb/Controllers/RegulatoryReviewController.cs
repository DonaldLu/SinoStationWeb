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

        // =============== Web API ================

        // 上傳Excel檔
        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file == null) return Json(new { Status = 0, Message = "No File Selected" });

            List<Room> memberList = _service.Upload(file);
            string names = string.Empty;
            foreach (Room member in memberList)
            {
                if (member.name != "")
                {
                    names += member.name + "、";
                }
            }
            var ret = names;
            return Json(memberList);
        }
    }
}