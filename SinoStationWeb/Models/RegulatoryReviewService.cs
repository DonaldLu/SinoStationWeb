//using ClosedXML.Excel;
//using DocumentFormat.OpenXml.Spreadsheet;
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
    }
}