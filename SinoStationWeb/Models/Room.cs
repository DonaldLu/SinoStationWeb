namespace SinoStationWeb.Models
{
    public class RuleName
    {
        public string name { get; set; }
        public string sqlName { get; set; }
    }
    public class Room
    {
        public int id { get; set; } // ID
        public string code { get; set; } // 代碼
        public string classification { get; set; } // 區域
        public string level { get; set; } // 樓層
        public string name { get; set; } // 空間名稱(中文)
        public string engName { get; set; }  // 空間名稱(英文)
        public string otherNames { get; set; } // 其他名稱
        public string system { get; set; } // 設備/系統
        public double count { get; set; } // 數量
        public double maxArea { get; set; } // 最大面積(m2)
        public double minArea { get; set; } // 最小面積(m2)
        public double demandArea { get; set; } // 需求面積
        public double permit { get; set; } // 容許差異(±%)
        public double specificationMinWidth { get; set; } // 最小規範寬度
        public double demandMinWidth { get; set; } // 最小需求寬度
        public double unboundedHeight { get; set; } // 規範淨高
        public double demandUnboundedHeight { get; set; } // 需求淨高
        public string door { get; set; } // 門
        public double doorWidth { get; set; } // 門寬(mm)
        public double doorHeight { get; set; } // 門高(mm)
    }
    public class TitalNames
    {
        public int id = 0; // Id
        public int code = 0; // 代碼
        public int classification = 0; // 區域
        public int level = 0; // 樓層
        public int name = 0; // 空間名稱(中文)
        public int engName = 0; // 空間名稱(英文)
        public int otherName = 0; // 其他名稱
        public int system = 0; // 設備/系統
        public int count = 0; // 數量
        public int demandArea = 0; // 需求面積
        public int maxArea = 0; // 最大面積(m2)
        public int minArea = 0; // 最小面積(m2)
        public int permit = 0; // 容許差異(±%)
        public int specificationMinWidth = 0; // 最小規範寬度
        public int demandMinWidth = 0; // 最小需求寬度
        public int unboundedHeight = 0; // 規範淨高
        public int demandUnboundedHeight = 0; // 需求淨高
        public int door = 0; // 門(mm)
    }
}