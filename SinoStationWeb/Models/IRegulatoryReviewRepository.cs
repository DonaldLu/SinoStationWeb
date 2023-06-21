using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SinoStationWeb.Models
{
    public interface IRegulatoryReviewRepository
    {
        // 上傳Excel檔
        List<Room> Upload(HttpPostedFileBase file);
    }
}
