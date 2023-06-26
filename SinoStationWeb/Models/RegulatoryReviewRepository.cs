using Dapper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Web;

namespace SinoStationWeb.Models
{
    public class RegulatoryReviewRepository : IRegulatoryReviewRepository, IDisposable
    {
        private IDbTransaction Transaction { get; set; }
        private IDbConnection conn;
        public RegulatoryReviewRepository()
        {
            string regulatoryReviewConnection = ConfigurationManager.ConnectionStrings["RegulatoryReviewConnection"].ConnectionString;
            conn = new SQLiteConnection(regulatoryReviewConnection);
        }
        // 上傳單一檔案
        public List<Room> Upload(HttpPostedFileBase file)
        {
            string filePath = string.Empty;
            // 先檢視是否有設定好要移除的特殊符號
            List<string> charsToRemove = CreateCharsToRemoveTXT();
            List<Room> roomList = new List<Room>();

            try
            {
                // 儲存檔案
                if (file != null)
                {
                    if (file.ContentLength > 0)
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        var path = Path.Combine(HttpContext.Current.Server.MapPath("~/FileUploads"), fileName);
                        file.SaveAs(path);
                        filePath = path;
                    }
                }

                //開檔
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // 關閉新許可模式通知
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    //載入Excel檔案
                    using (ExcelPackage ep = new ExcelPackage(fs))
                    {
                        ExcelWorksheet workSheet = ep.Workbook.Worksheets[0];//取得Sheet1
                        int rowCount = workSheet.Dimension.End.Row;//結束列編號，從1算起
                        int colCount = workSheet.Dimension.End.Column;//結束欄編號，從1算起

                        // 記錄標頭的欄位數
                        TitalNames titleNames = SaveTitleNames(colCount, workSheet);

                        // 讀取Excel檔中, 所有物件的名稱、類別、數量
                        int id = 1;
                        for (int i = 2; i <= rowCount; i++)
                        {
                            // 空間名稱(中文)
                            if (workSheet.Cells[i, titleNames.name].Value != null)
                            {
                                Room room = new Room();
                                room = SaveExcelValue(id, room, titleNames, workSheet, charsToRemove, i); // 儲存Excel資料
                                roomList.Add(room);
                                id++;
                            }
                        }
                    }
                }

                InsertToSQL(roomList); // 新增Excel資料至SQL
            }
            catch (Exception)
            {

            }

            return roomList;
        }
        // 先檢視是否有設定好要移除的特殊符號
        public static List<string> CreateCharsToRemoveTXT()
        {
            List<string> charsToRemove = new List<string>();
            string charsToRemovePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\CharsToRemove.txt"; // 取得使用者文件路徑
            // 先檢查是否有此檔案, 沒有的話則新增
            if (!File.Exists(charsToRemovePath))
            {
                string[] signs = new string[] { "@", ",", ".", ";", "'", "(", ")", "_", "-", "\\", "/", " ", "\"" }; // 特殊符號
                foreach (string sign in signs)
                {
                    charsToRemove.Add(sign);
                }
                using (StreamWriter outputFile = new StreamWriter(charsToRemovePath))
                {
                    foreach (string sign in charsToRemove)
                    {
                        outputFile.WriteLine(sign);
                    }
                }
            }
            else
            {
                charsToRemove = new List<string>();
                using (StreamReader sr = new StreamReader(charsToRemovePath))
                {
                    string textContent;
                    while ((textContent = sr.ReadLine()) != null)
                    {
                        charsToRemove.Add(textContent);
                    }
                }
            }

            return charsToRemove;
        }
        // 讀取標頭順序
        private TitalNames SaveTitleNames(int colCount, ExcelWorksheet workSheet)
        {
            TitalNames titleNames = new TitalNames();
            for (int i = 1; i <= colCount; i++)
            {
                string titleName = (string)workSheet.Cells[1, i].Value;
                string title = titleName.Replace("\n", "");
                if (title.Equals("代碼"))
                {
                    titleNames.code = i;
                }
                else if (title.Equals("區域"))
                {
                    titleNames.classification = i;
                }
                else if (title.Equals("樓層"))
                {
                    titleNames.level = i;
                }
                else if (title.Equals("空間名稱(中文)"))
                {
                    titleNames.name = i;
                }
                else if (title.Equals("空間名稱(英文)"))
                {
                    titleNames.engName = i;
                }
                else if (title.Equals("其他名稱"))
                {
                    titleNames.otherName = i;
                }
                else if (title.Equals("類別") || title.Equals("設備/系統"))
                {
                    titleNames.system = i;
                }
                else if (title.Equals("數量"))
                {
                    titleNames.count = i;
                }
                else if (title.Equals("最大面積(m2)") || title.Equals("規範最大面積(m2)"))
                {
                    titleNames.maxArea = i;
                }
                else if (title.Equals("最小面積(m2)") || title.Equals("規範最小面積(m2)"))
                {
                    titleNames.minArea = i;
                }
                else if (title.Equals("需求面積(m2)"))
                {
                    titleNames.demandArea = i;
                }
                else if (title.Equals("容許差異(±%)") || title.Equals("面積容許差異(±%)"))
                {
                    titleNames.permit = i;
                }
                else if (title.Equals("規範最小寬度(m)"))
                {
                    titleNames.specificationMinWidth = i;
                }
                else if (title.Equals("實際最小寬度(m)") || title.Equals("需求最小寬度(m)"))
                {
                    titleNames.demandMinWidth = i;
                }
                else if (title.Equals("門(mm)"))
                {
                    titleNames.door = i;
                }

                if (title.Equals("規範淨高(m)"))
                {
                    titleNames.unboundedHeight = i;
                }
                else if (title.Equals("淨高(m)"))
                {
                    titleNames.unboundedHeight = i;
                }

                if (title.Equals("需求淨高(m)"))
                {
                    titleNames.demandUnboundedHeight = i;
                }
                else if (title.Equals("淨高(m)"))
                {
                    titleNames.demandUnboundedHeight = i;
                }
            }
            return titleNames;
        }
        // 儲存Excel資料
        private Room SaveExcelValue(int id, Room excelCompare, TitalNames titleNames, ExcelWorksheet workSheet, List<string> charsToRemove, int i)
        {
            try
            {
                excelCompare.id = id;
                excelCompare.code = (string)workSheet.Cells[i, titleNames.code].Value; // 代碼
                if (excelCompare.code == null)
                {
                    excelCompare.code = "";
                }
                excelCompare.classification = (string)workSheet.Cells[i, titleNames.classification].Value; // 區域
                excelCompare.level = (string)workSheet.Cells[i, titleNames.level].Value; // 樓層
                // 名稱(設定)
                string editName = (string)workSheet.Cells[i, titleNames.name].Value;
                foreach (string c in charsToRemove)
                {
                    try
                    {
                        editName = editName.Replace(c, string.Empty); // 空間名稱(中文)
                    }
                    catch (Exception ex)
                    {
                        string error = ex.Message + "\n" + ex.ToString();
                    }
                }
                excelCompare.name = editName;
                try
                {
                    excelCompare.engName = (string)workSheet.Cells[i, titleNames.engName].Value; // 空間名稱(英文)
                }
                catch (Exception)
                {
                    excelCompare.engName = "";
                }
                try
                {
                    // 檢查此空間名稱是否有"其他名稱", 有的話則過濾分隔符號後儲存
                    try
                    {
                        string otherFullName = (string)workSheet.Cells[i, titleNames.otherName].Value; // 其他名稱
                        if (otherFullName != "")
                        {
                            otherFullName = otherFullName.Replace(",", "、").Replace("/", "、");
                            string[] otherNames = otherFullName.Split('、');
                            string allOtherNames = string.Empty;
                            foreach (string otherName in otherNames)
                            {
                                allOtherNames += otherName + "、";
                            }
                            excelCompare.otherNames = allOtherNames.Substring(0, allOtherNames.Length - 1);
                        }
                    }
                    catch (Exception)
                    {
                        excelCompare.engName = "";
                    }
                }
                catch (Exception)
                {

                }
                try
                {
                    excelCompare.system = (string)workSheet.Cells[i, titleNames.system].Value; // 設備/系統
                }
                catch (Exception)
                {

                }
                try
                {
                    excelCompare.count = (double)workSheet.Cells[i, titleNames.count].Value; // 數量
                }
                catch (Exception)
                {
                    excelCompare.count = 0;
                }
                try
                {
                    excelCompare.permit = (double)workSheet.Cells[i, titleNames.permit].Value; // 容許差異
                }
                catch (Exception)
                {
                    excelCompare.permit = 0.0;
                }
                try
                {
                    excelCompare.unboundedHeight = (double)workSheet.Cells[i, titleNames.unboundedHeight].Value; // 規範淨高
                }
                catch (Exception)
                {
                    excelCompare.unboundedHeight = 0.0;
                }
                try
                {
                    excelCompare.demandUnboundedHeight = (double)workSheet.Cells[i, titleNames.demandUnboundedHeight].Value; // 需求淨高
                }
                catch (Exception)
                {
                    excelCompare.demandUnboundedHeight = 0.0;
                }
                string doorWidthHeight = (string)workSheet.Cells[i, titleNames.door].Value;
                if (doorWidthHeight != null)
                {
                    try
                    {
                        excelCompare.doorWidth = Convert.ToDouble(doorWidthHeight.Split('x')[0]); // 門寬(mm)
                        excelCompare.doorHeight = Convert.ToDouble(doorWidthHeight.Split('x')[1]); // 門高(mm)
                    }
                    catch (Exception)
                    {
                        excelCompare.doorWidth = 0.0;
                        excelCompare.doorHeight = 0.0;
                    }
                }
                try
                {
                    excelCompare.maxArea = (double)workSheet.Cells[i, titleNames.maxArea].Value; // 規範最大面積
                }
                catch (Exception)
                {
                    excelCompare.maxArea = 0;
                }
                try
                {
                    excelCompare.minArea = (double)workSheet.Cells[i, titleNames.minArea].Value; // 規範最小面積
                }
                catch (Exception)
                {
                    excelCompare.minArea = 0;
                }
                try
                {
                    excelCompare.specificationMinWidth = (double)workSheet.Cells[i, titleNames.specificationMinWidth].Value; // 規範最小寬度
                }
                catch (Exception)
                {
                    excelCompare.specificationMinWidth = 0;
                }
                try
                {
                    excelCompare.demandArea = (double)workSheet.Cells[i, titleNames.demandArea].Value; // 需求面積
                }
                catch (Exception)
                {
                    excelCompare.demandArea = 0;
                }
                try
                {
                    excelCompare.demandMinWidth = (double)workSheet.Cells[i, titleNames.demandMinWidth].Value; // 需求最小寬度
                }
                catch (Exception)
                {
                    excelCompare.demandMinWidth = 0;
                }
            }
            catch (Exception)
            {

            }

            return excelCompare;
        }
        // 新增Excel資料至SQL
        private void InsertToSQL(List<Room> roomList)
        {
            if (conn.State == 0)
                conn.Open();

            using (var tran = conn.BeginTransaction())
            {
                try
                {
                    string sql = @"INSERT INTO Room (id, code, classification, level, name, engName, otherNames, system, count, maxArea, minArea, demandArea, permit, specificationMinWidth, demandMinWidth, unboundedHeight, demandUnboundedHeight, door, doorWidth, doorHeight)
                                   VALUES(@id, @code, @classification, @level, @name, @engName, @otherNames, @system, @count, @maxArea, @minArea, @demandArea, @permit, @specificationMinWidth, @demandMinWidth, @unboundedHeight, @demandUnboundedHeight, @door, @doorWidth, @doorHeight)";
                    conn.Execute(sql, roomList);
                    tran.Commit();
                }
                catch(Exception)
                { 

                }
            }
        }
        public void Dispose()
        {
            conn.Close();
            conn.Dispose();
            return;
        }
    }
}