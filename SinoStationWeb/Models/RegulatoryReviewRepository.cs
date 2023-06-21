//using ClosedXML.Excel;
using Dapper;
//using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

                if (Path.GetExtension(file.FileName) != ".xlsx") throw new ApplicationException("請使用Excel 2007(.xlsx)格式");

                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(filePath);
                Worksheet workSheet = workbook.Sheets[1];
                Range Range = workSheet.UsedRange;

                int rowCount = Range.Rows.Count;
                int colCount = Range.Columns.Count;

                // 記錄標頭的欄位數
                TitalNames titleNames = SaveTitleNames(colCount, workSheet);

                // 讀取Excel檔中, 所有物件的名稱、類別、數量
                for (int i = 2; i <= rowCount; i++)
                {
                    // 空間名稱(中文)
                    if (workSheet.Cells[i, titleNames.name].Value != null)
                    {
                        Room room = new Room();
                        room = SaveExcelValue(room, titleNames, workSheet, charsToRemove, i); // 儲存Excel資料
                        roomList.Add(room);
                    }
                }

                // 清理記憶體
                GC.Collect();
                GC.WaitForPendingFinalizers();
                // 釋放COM對象的經驗法則, 單獨引用與釋放COM對象, 不要使用多"."釋放
                Marshal.ReleaseComObject(Range);
                Marshal.ReleaseComObject(workSheet);
                // 關閉與釋放
                workbook.Close();
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

                EditToSQL(roomList);
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
        private TitalNames SaveTitleNames(int colCount, Worksheet workSheet)
        {
            TitalNames titleNames = new TitalNames();
            for (int i = 1; i <= colCount; i++)
            {
                string titleName = workSheet.Cells[1, i].Value;
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
                    titleNames.category = i;
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
        private Room SaveExcelValue(Room excelCompare, TitalNames titleNames, Worksheet workSheet, List<string> charsToRemove, int i)
        {
            try
            {
                excelCompare.code = workSheet.Cells[i, titleNames.code].Value; // 代碼
                if (excelCompare.code == null)
                {
                    excelCompare.code = "";
                }
                excelCompare.classification = workSheet.Cells[i, titleNames.classification].Value; // 區域
                excelCompare.level = workSheet.Cells[i, titleNames.level].Value; // 樓層
                // 名稱(設定)
                string editName = workSheet.Cells[i, titleNames.name].Value;
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
                    excelCompare.engName = workSheet.Cells[i, titleNames.engName].Value; // 空間名稱(英文)
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
                        string otherFullName = workSheet.Cells[i, titleNames.otherName].Value; // 其他名稱
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
                    excelCompare.permit = workSheet.Cells[i, titleNames.permit].Value; // 容許差異
                }
                catch (Exception)
                {
                    excelCompare.permit = 0.0;
                }
                try
                {
                    excelCompare.unboundedHeight = workSheet.Cells[i, titleNames.unboundedHeight].Value; // 規範淨高
                }
                catch (Exception)
                {
                    excelCompare.unboundedHeight = 0.0;
                }
                try
                {
                    excelCompare.demandUnboundedHeight = workSheet.Cells[i, titleNames.demandUnboundedHeight].Value; // 需求淨高
                }
                catch (Exception)
                {
                    excelCompare.demandUnboundedHeight = 0.0;
                }
                string doorWidthHeight = workSheet.Cells[i, titleNames.door].Value;
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
                    excelCompare.maxArea = workSheet.Cells[i, titleNames.maxArea].Value; // 規範最大面積
                }
                catch (Exception)
                {
                    excelCompare.maxArea = 0;
                }
                try
                {
                    excelCompare.minArea = workSheet.Cells[i, titleNames.minArea].Value; // 規範最小面積
                }
                catch (Exception)
                {
                    excelCompare.minArea = 0;
                }
                try
                {
                    excelCompare.specificationMinWidth = workSheet.Cells[i, titleNames.specificationMinWidth].Value; // 規範最小寬度
                }
                catch (Exception)
                {
                    excelCompare.specificationMinWidth = 0;
                }
                try
                {
                    excelCompare.demandArea = workSheet.Cells[i, titleNames.demandArea].Value; // 需求面積
                }
                catch (Exception)
                {
                    excelCompare.demandArea = 0;
                }
                try
                {
                    excelCompare.demandMinWidth = workSheet.Cells[i, titleNames.demandMinWidth].Value; // 需求最小寬度
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
        // 更新SQL分數
        private void EditToSQL(List<Room> roomList)
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