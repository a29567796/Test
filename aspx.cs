using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Services;
using OfficeOpenXml;      // 需先以 NuGet 安裝 EPPlus
using System.IO;

namespace WebApp
{
    public partial class AnnualReport : System.Web.UI.Page
    {
        /* ===== WebMethod：AJAX 取數 ===== */
        [WebMethod]
        public static List<Dictionary<string, object>> GetReportData(
            string department, string business, string region, int[] yearList,
            string customer, string custItem, string substrate,
            string productType, string shipTo, string realDate,
            string status, string orderType)
        {
            DateTime? realDateVal = null;
            if (!string.IsNullOrEmpty(realDate))
                realDateVal = DateTime.ParseExact(realDate, "yyyy/MM/dd", null);

            DataTable dt = FetchReportData(
                department, business, region, yearList ?? new int[0],
                customer, custItem, substrate,
                productType, shipTo, realDateVal, status, orderType);

            // DataTable → List<Dictionary> 方便 JSON 序列化
            var list = new List<Dictionary<string, object>>();
            foreach (DataRow row in dt.Rows)
            {
                var dict = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                    dict[col.ColumnName] = row[col];
                list.Add(dict);
            }
            return list;
        }

        /* ===== 匯出 Excel 按鈕 ===== */
        protected void btnExport_Click(object sender, EventArgs e)
        {
            // 透過 Request.Form 取值（與 AJAX 參數一致）
            int[] years = Request.Form.GetValues("year")?
                             .Select(int.Parse).ToArray() ?? new int[0];

            DataTable dt = FetchReportData(
                Request.Form["ddlDept"],
                Request.Form["ddlBiz"],
                Request.Form["ddlRegion"],
                years,
                Request.Form["ddlCustomer"],
                Request.Form["txtCustItem"],
                Request.Form["ddlSubstrate"],
                Request.Form["ddlProductType"],
                Request.Form["txtShipTo"],
                DateTime.TryParse(Request.Form["txtRealDate"], out DateTime rd) ? (DateTime?)rd : null,
                Request.Form["ddlStatus"],
                Request.Form["ddlOrderType"]
            );

            // ===== 使用 EPPlus 產出 .xlsx =====
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("Report");
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=AnnualReport.xlsx");
                using (var ms = new MemoryStream())
                {
                    pck.SaveAs(ms);
                    ms.WriteTo(Response.OutputStream);
                }
                Response.End();
            }
        }

        /* ===== 取數核心：實務上請改為資料庫查詢 ===== */
        private static DataTable FetchReportData(
            string dept, string biz, string region, int[] years,
            string customer, string custItem, string substrate,
            string productType, string shipTo, DateTime? realDate,
            string status, string orderType)
        {
            // === 此處示範用假資料，請自行改為 SQL / SP 取數 ===
            DataTable dt = new DataTable();

            // 靜態欄
            dt.Columns.AddRange(new[] {
                new DataColumn("Thk1", typeof(double)),
                new DataColumn("Thk2", typeof(double)),
                new DataColumn("Thk3", typeof(double)),
                new DataColumn("Res1", typeof(int)),
                new DataColumn("Res2", typeof(int)),
                new DataColumn("Res3", typeof(int))
            });

            // 動態年度欄
            string[] months = { "Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec" };
            foreach (int y in years.OrderBy(x => x))
            {
                foreach (string m in months)       dt.Columns.Add($"{y}_{m}", typeof(double));
                dt.Columns.Add($"{y}_Q1_Total", typeof(double));
                dt.Columns.Add($"{y}_Q2_Total", typeof(double));
                dt.Columns.Add($"{y}_Q3_Total", typeof(double));
                dt.Columns.Add($"{y}_Q4_Total", typeof(double));
                dt.Columns.Add($"{y}_Total",      typeof(double));
            }

            // ------ 假資料列 1 ------
            DataRow r1 = dt.NewRow();
            r1["Thk1"] = 1.2; r1["Thk2"] = 1.4; r1["Thk3"] = 1.6;
            r1["Res1"] = 100; r1["Res2"] = 110; r1["Res3"] = 120;
            foreach (int y in years)
            {
                r1[$"{y}_Jan"] = 10;  r1[$"{y}_Feb"] = 15;  r1[$"{y}_Mar"] = 20;  r1[$"{y}_Q1_Total"] = 45;
                r1[$"{y}_Apr"] = 12;  r1[$"{y}_May"] = 18;  r1[$"{y}_Jun"] = 22;  r1[$"{y}_Q2_Total"] = 52;
                r1[$"{y}_Jul"] = 14;  r1[$"{y}_Aug"] = 19;  r1[$"{y}_Sep"] = 24;  r1[$"{y}_Q3_Total"] = 57;
                r1[$"{y}_Oct"] = 16;  r1[$"{y}_Nov"] = 20;  r1[$"{y}_Dec"] = 25;  r1[$"{y}_Q4_Total"] = 61;
                r1[$"{y}_Total"] = 215;
            }
            dt.Rows.Add(r1);

            // ------ 假資料列 2 ------
            DataRow r2 = dt.NewRow();
            r2["Thk1"] = 0.8; r2["Thk2"] = 1.0; r2["Thk3"] = 1.1;
            r2["Res1"] = 80;  r2["Res2"] = 85;  r2["Res3"] = 90;
            foreach (int y in years)
            {
                r2[$"{y}_Jan"] = 5;   r2[$"{y}_Feb"] = 6;   r2[$"{y}_Mar"] = 7;   r2[$"{y}_Q1_Total"] = 18;
                r2[$"{y}_Apr"] = 6;   r2[$"{y}_May"] = 7;   r2[$"{y}_Jun"] = 8;   r2[$"{y}_Q2_Total"] = 21;
                r2[$"{y}_Jul"] = 7;   r2[$"{y}_Aug"] = 8;   r2[$"{y}_Sep"] = 9;   r2[$"{y}_Q3_Total"] = 24;
                r2[$"{y}_Oct"] = 8;   r2[$"{y}_Nov"] = 9;   r2[$"{y}_Dec"] =10;   r2[$"{y}_Q4_Total"] = 27;
                r2[$"{y}_Total"] = 90;
            }
            dt.Rows.Add(r2);

            return dt;
        }
    }
}