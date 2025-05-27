using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Services;
using OfficeOpenXml;
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

            // DataTable -> List<Dictionary> 供 JSON
            return dt.AsEnumerable()
                     .Select(dr => dt.Columns.Cast<DataColumn>()
                        .ToDictionary(c => c.ColumnName, c => dr[c]))
                     .ToList();
        }

        /* ===== 匯出 Excel ===== */
        protected void btnExport_Click(object sender, EventArgs e)
        {
            int[] years = Request.Form.GetValues("year")?.Select(int.Parse).ToArray() ?? new int[0];

            DataTable dt = FetchReportData(
                Request.Form["ddlDept"], Request.Form["ddlBiz"], Request.Form["ddlRegion"],
                years,
                Request.Form["ddlCustomer"], Request.Form["txtCustItem"], Request.Form["ddlSubstrate"],
                Request.Form["ddlProductType"], Request.Form["txtShipTo"],
                DateTime.TryParse(Request.Form["txtRealDate"], out DateTime rd) ? (DateTime?)rd : null,
                Request.Form["ddlStatus"], Request.Form["ddlOrderType"]);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Report");
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.Cells.AutoFitColumns();

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=AnnualReport.xlsx");
                using (var ms = new MemoryStream())
                {
                    p.SaveAs(ms);
                    ms.WriteTo(Response.OutputStream);
                }
                Response.End();
            }
        }

        /* ===== 取數核心 (示範 20 筆假資料) ===== */
        private static DataTable FetchReportData(
            string dept, string biz, string region, int[] years,
            string customer, string custItem, string substrate,
            string productType, string shipTo, DateTime? realDate,
            string status, string orderType)
        {
            if (years.Length == 0) years = new[] { DateTime.Now.Year };

            DataTable dt = new DataTable();

            /* --- 12 固定欄 --- */
            dt.Columns.Add("Customer", typeof(string));
            dt.Columns.Add("Inch", typeof(double));
            dt.Columns.Add("Cust Item", typeof(string));
            dt.Columns.Add("OA Item", typeof(string));
            dt.Columns.Add("Substrate", typeof(string));
            dt.Columns.Add("Price", typeof(double));
            dt.Columns.Add("Thk1", typeof(double));
            dt.Columns.Add("Res1", typeof(int));
            dt.Columns.Add("Thk2", typeof(double));
            dt.Columns.Add("Res2", typeof(int));
            dt.Columns.Add("Thk3", typeof(double));
            dt.Columns.Add("Res3", typeof(int));

            /* --- 動態年度欄 (17 欄 / 年) --- */
            string[] m = { "Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec" };
            foreach (int y in years.OrderBy(x => x))
            {
                foreach (string mon in m) dt.Columns.Add($"{y}_{mon}", typeof(double));
                dt.Columns.Add($"{y}_Q1_Total", typeof(double));
                dt.Columns.Add($"{y}_Q2_Total", typeof(double));
                dt.Columns.Add($"{y}_Q3_Total", typeof(double));
                dt.Columns.Add($"{y}_Q4_Total", typeof(double));
                dt.Columns.Add($"{y}_Total",      typeof(double));
            }

            /* --- 產生 20 筆假資料 --- */
            Random rnd = new Random();
            for (int i = 1; i <= 20; i++)
            {
                DataRow r = dt.NewRow();
                r["Customer"] = "Cust" + (i % 3 + 1);
                r["Inch"] = (i % 2 == 0) ? 6 : 8;
                r["Cust Item"] = "Item" + (char)('A' + i % 3);
                r["OA Item"] = "OA" + i.ToString("D3");
                r["Substrate"] = (i % 2 == 0) ? "SubA" : "SubB";
                r["Price"] = 100 + i;
                r["Thk1"] = Math.Round(rnd.NextDouble() + 1, 2);
                r["Res1"] = rnd.Next(80, 120);
                r["Thk2"] = Math.Round(rnd.NextDouble() + 1.2, 2);
                r["Res2"] = rnd.Next(80, 120);
                r["Thk3"] = Math.Round(rnd.NextDouble() + 1.4, 2);
                r["Res3"] = rnd.Next(80, 120);

                foreach (int y in years)
                {
                    double sum = 0;
                    foreach (string mon in m)
                    {
                        double v = rnd.Next(5, 25);
                        r[$"{y}_{mon}"] = v;
                        sum += v;
                    }
                    r[$"{y}_Q1_Total"] = Convert.ToDouble(r[$"{y}_Jan"]) + Convert.ToDouble(r[$"{y}_Feb"]) + Convert.ToDouble(r[$"{y}_Mar"]);
                    r[$"{y}_Q2_Total"] = Convert.ToDouble(r[$"{y}_Apr"]) + Convert.ToDouble(r[$"{y}_May"]) + Convert.ToDouble(r[$"{y}_Jun"]);
                    r[$"{y}_Q3_Total"] = Convert.ToDouble(r[$"{y}_Jul"]) + Convert.ToDouble(r[$"{y}_Aug"]) + Convert.ToDouble(r[$"{y}_Sep"]);
                    r[$"{y}_Q4_Total"] = Convert.ToDouble(r[$"{y}_Oct"]) + Convert.ToDouble(r[$"{y}_Nov"]) + Convert.ToDouble(r[$"{y}_Dec"]);
                    r[$"{y}_Total"] = sum;
                }
                dt.Rows.Add(r);
            }
            return dt;
        }
    }
}
