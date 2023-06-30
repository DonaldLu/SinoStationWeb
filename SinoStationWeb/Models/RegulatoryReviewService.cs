//using ClosedXML.Excel;
//using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Web;

namespace SinoStationWeb.Models
{
    public class RegulatoryReviewService
    {
        private IRegulatoryReviewRepository _regulatoryReviewRepository;

        public RegulatoryReviewService()
        {
            _regulatoryReviewRepository = new RegulatoryReviewRepository();
        }
        // 上傳Excel檔
        internal List<Room> Upload(HttpPostedFileBase file)
        {
            var ret = _regulatoryReviewRepository.Upload(file);
            return ret;
        }
        // 讀取所有規則
        public List<RuleName> AllRule()
        {
            var ret = _regulatoryReviewRepository.AllRule();
            return ret;
        }
        // 取得SQL名稱
        public string GetName(string sqlName)
        {
            var ret = _regulatoryReviewRepository.GetName(sqlName);
            return ret;
        }
        // 讀取SQL資料
        internal List<Room> GetSQLData(string sqlName)
        {
            var ret = _regulatoryReviewRepository.GetSQLData(sqlName);
            return ret;
        }
    }
}