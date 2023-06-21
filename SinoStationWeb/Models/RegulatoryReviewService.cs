//using ClosedXML.Excel;
//using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;
using System.Windows;

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