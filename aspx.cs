using System;
using System.Collections.Generic;
using System.Data;
using System.Web.Script.Serialization;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class AnnualReport : System.Web.UI.Page
{
    // 靜態 WebMethod：接受前端 AJAX 請求，回傳 JSON 資料字串
    [WebMethod]
    public static string GetReportData(string dept, string business, string region, int[] years, string customer, string custItem, string substrate)
    {
        // 構建查詢結果的 DataTable
        DataTable dt = BuildReportData(dept, business, region, years, customer, custItem, substrate);
        // 將 DataTable 序列化為 JSON 字串返回11
        JavaScriptSerializer js = new JavaScriptSerializer();
        js.MaxJsonLength = int.MaxValue; // 如資料量大，可適當增加序列化最大長度
        List<Dictionary<string, object>> rowsList = new List<Dictionary<string, object>>();
        foreach (DataRow dr in dt.Rows)
        {
            Dictionary<string, object> rowData = new Dictionary<string, object>();
            foreach (DataColumn col in dt.Columns)
            {
                rowData[col.ColumnName] = dr[col];
            }
            rowsList.Add(rowData);
        }
        string jsonResult = js.Serialize(rowsList);
        return jsonResult;
    }

    // Excel 匯出按鈕事件處理：將完整報表資料輸出成 Excel 檔案
    protected void btnExport_Click(object sender, EventArgs e)
    {
        // 從表單取得當前查詢條件值（與前端相同，使用 Request.Form）
        string dept = Request.Form["deptSelect"];
        string business = Request.Form["businessSelect"];
        string region = Request.Form["regionSelect"];
        string[] yearValues = Request.Form.GetValues("year"); // 取得多選年份陣列
        int[] years = Array.ConvertAll(yearValues ?? new string[0], int.Parse);
        string customer = Request.Form["customerSelect"];
        string custItem = Request.Form["itemSelect"];
        string substrate = Request.Form["substrateSelect"];

        // 生成完整 DataTable 資料
        DataTable dt = BuildReportData(dept, business, region, years, customer, custItem, substrate);

        // 綁定 GridView 並匯出為 Excel
        GridView gv = new GridView();
        gv.DataSource = dt;
        gv.DataBind();

        // 清除響應並設定輸出格式為 Excel 檔案
        Response.ClearContent();
        Response.Buffer = true;
        Response.AddHeader("content-disposition", "attachment; filename=AnnualReport.xls");
        Response.ContentType = "application/vnd.ms-excel";
        Response.Charset = "UTF-8";
        // 注意：使用 HTML 輸出 Excel 可能會有格式警告12

        System.IO.StringWriter sw = new System.IO.StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        // 將 GridView (HTML 表格) 寫入 HtmlTextWriter
        gv.RenderControl(htw);
        // 輸出到回應流
        Response.Write(sw.ToString());
        Response.End();
    }

    // 為允許 GridView.RenderControl 匯出，需覆寫此方法
    public override void VerifyRenderingInServerForm(Control control)
    {
        // 不做任何處理，僅用來繞過 ASP.NET 檢查
        // (GridView 匯出 Excel 需要此覆寫)
    }

    /// <summary>
    /// 根據查詢條件建立報表 DataTable，動態加入所選年份的欄位並合併資料。
    /// （此函式模擬資料來源，實際應改為從資料庫查詢填充）
    /// </summary>
    private static DataTable BuildReportData(string dept, string business, string region, int[] years, string customer, string custItem, string substrate)
    {
        DataTable dt = new DataTable();
        // 定義靜態欄位結構
        dt.Columns.Add("部門", typeof(string));
        dt.Columns.Add("業務", typeof(string));
        dt.Columns.Add("區域", typeof(string));
        dt.Columns.Add("客戶", typeof(string));
        dt.Columns.Add("Cust Item", typeof(string));
        dt.Columns.Add("Substrate", typeof(string));
        dt.Columns.Add("Thk1", typeof(double));
        dt.Columns.Add("Thk2", typeof(double));
        dt.Columns.Add("Thk3", typeof(double));
        dt.Columns.Add("Res1", typeof(int));
        dt.Columns.Add("Res2", typeof(int));
        dt.Columns.Add("Res3", typeof(int));
        // 排序年份，確保欄位順序一致
        Array.Sort(years);
        // 動態定義每個年份的欄位 (每年 12個月 + 4個季度合計 + 年總計 = 17 欄)
        string[] months = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
        foreach (int year in years)
        {
            // 每月欄位
            foreach (string m in months)
            {
                dt.Columns.Add(m + "_" + year, typeof(double));
            }
            // 每季合計欄位 Q1~Q4
            for (int q = 1; q <= 4; q++)
            {
                dt.Columns.Add("Q" + q + "_" + year, typeof(double));
            }
            // 年度總計欄位
            dt.Columns.Add("Total_" + year, typeof(double));
        }

        // 模擬資料填充：
        // 假設有兩筆產品/客戶資料，其鍵欄可能有重複（用於測試多年度合併）
        // 範例資料1: 在2024和2025年都有紀錄
        DataRow r1_2024 = dt.NewRow();
        r1_2024["部門"] = "S100";
        r1_2024["業務"] = "Alice";
        r1_2024["區域"] = "TW";
        r1_2024["客戶"] = "VI";
        r1_2024["Cust Item"] = "Internal";
        r1_2024["Substrate"] = "SubA";
        r1_2024["Thk1"] = 1.2; r1_2024["Thk2"] = 1.5; r1_2024["Thk3"] = 1.7;
        r1_2024["Res1"] = 100; r1_2024["Res2"] = 110; r1_2024["Res3"] = 120;
        // 填入2024年的月份與合計值（僅示意，實際應計算合計）
        r1_2024["Jan_2024"] = 50; r1_2024["Feb_2024"] = 60; r1_2024["Mar_2024"] = 55;
        r1_2024["Q1_2024"] = 165; // 假設 Q1 合計
        r1_2024["Apr_2024"] = 70; r1_2024["May_2024"] = 65; r1_2024["Jun_2024"] = 80;
        r1_2024["Q2_2024"] = 215;
        r1_2024["Jul_2024"] = 90; r1_2024["Aug_2024"] = 85; r1_2024["Sep_2024"] = 95;
        r1_2024["Q3_2024"] = 270;
        r1_2024["Oct_2024"] = 100; r1_2024["Nov_2024"] = 110; r1_2024["Dec_2024"] = 105;
        r1_2024["Q4_2024"] = 315;
        r1_2024["Total_2024"] = 965;
        // 2025年同鍵資料
        DataRow r1_2025 = dt.NewRow();
        r1_2025.ItemArray = r1_2024.ItemArray.Clone() as object[]; // 複製靜態欄資料與Thk/Res
        r1_2025["Jan_2025"] = 55; r1_2025["Feb_2025"] = 65; r1_2025["Mar_2025"] = 60;
        r1_2025["Q1_2025"] = 180;
        r1_2025["Apr_2025"] = 75; r1_2025["May_2025"] = 70; r1_2025["Jun_2025"] = 85;
        r1_2025["Q2_2025"] = 230;
        r1_2025["Jul_2025"] = 95; r1_2025["Aug_2025"] = 100; r1_2025["Sep_2025"] = 90;
        r1_2025["Q3_2025"] = 285;
        r1_2025["Oct_2025"] = 110; r1_2025["Nov_2025"] = 120; r1_2025["Dec_2025"] = 115;
        r1_2025["Q4_2025"] = 345;
        r1_2025["Total_2025"] = 1040;
        // 範例資料2: 僅在2025年有紀錄
        DataRow r2_2025 = dt.NewRow();
        r2_2025["部門"] = "S200";
        r2_2025["業務"] = "Bob";
        r2_2025["區域"] = "CN";
        r2_2025["客戶"] = "EPS";
        r2_2025["Cust Item"] = "External";
        r2_2025["Substrate"] = "SubB";
        r2_2025["Thk1"] = 0.8; r2_2025["Thk2"] = 1.0; r2_2025["Thk3"] = 1.1;
        r2_2025["Res1"] = 80; r2_2025["Res2"] = 85; r2_2025["Res3"] = 90;
        // 填入2025年的值
        r2_2025["Jan_2025"] = 30; r2_2025["Feb_2025"] = 35; r2_2025["Mar_2025"] = 40;
        r2_2025["Q1_2025"] = 105;
        r2_2025["Apr_2025"] = 45; r2_2025["May_2025"] = 50; r2_2025["Jun_2025"] = 55;
        r2_2025["Q2_2025"] = 150;
        r2_2025["Jul_2025"] = 60; r2_2025["Aug_2025"] = 65; r2_2025["Sep_2025"] = 70;
        r2_2025["Q3_2025"] = 195;
        r2_2025["Oct_2025"] = 75; r2_2025["Nov_2025"] = 80; r2_2025["Dec_2025"] = 85;
        r2_2025["Q4_2025"] = 240;
        r2_2025["Total_2025"] = 690;

        // 將資料列加入 DataTable
        dt.Rows.Add(r1_2024);
        dt.Rows.Add(r1_2025);
        dt.Rows.Add(r2_2025);

        // 合併多年度資料：由於我們直接在同一 DataTable 插入模擬的多年度欄位資料，
        // 這裡無需額外合併處理。如果實際上各年度資料分開查詢，可在此進行合併。
        // （例如，用主鍵比對填充不同年度欄位或新增新的資料列）

        return dt;
    }
}
