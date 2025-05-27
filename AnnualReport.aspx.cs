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
        /* ========= AJAX WebMethod ========= */
        [WebMethod]
        public static List<Dictionary<string,object>> GetReportData(int[] yearList)
        {
            // 其餘查詢參數略，僅示範
            var dt = BuildTestData(yearList??new[]{2024,2025});

            return dt.AsEnumerable()
                     .Select(r => dt.Columns.Cast<DataColumn>()
                         .ToDictionary(c=>c.ColumnName, c=>r[c]))
                     .ToList();
        }

        /* ========= 匯出 Excel ========= */
        protected void btnExport_Click(object sender, EventArgs e)
        {
            int[] years = Request.Form.GetValues("year")?.Select(int.Parse).ToArray()
                          ?? new[]{2024,2025};
            var dt = BuildTestData(years);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using(var p=new ExcelPackage())
            {
                var ws=p.Workbook.Worksheets.Add("Report");
                ws.Cells["A1"].LoadFromDataTable(dt,true);
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                Response.Clear();
                Response.ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition","attachment; filename=AnnualReport.xlsx");
                using(var ms=new MemoryStream())
                {
                    p.SaveAs(ms);
                    ms.WriteTo(Response.OutputStream);
                }
                Response.End();
            }
        }

        /* ========= 測試資料 (20 筆) ========= */
        private static DataTable BuildTestData(int[] years)
        {
            var dt=new DataTable();
            // 12 固定欄
            dt.Columns.Add("Customer",typeof(string));
            dt.Columns.Add("Inch",typeof(string));
            dt.Columns.Add("CustItem",typeof(string));
            dt.Columns.Add("OAItem",typeof(string));
            dt.Columns.Add("Substrate",typeof(string));
            dt.Columns.Add("Price",typeof(double));
            dt.Columns.Add("Thk1",typeof(double));
            dt.Columns.Add("Res1",typeof(int));
            dt.Columns.Add("Thk2",typeof(double));
            dt.Columns.Add("Res2",typeof(int));
            dt.Columns.Add("Thk3",typeof(double));
            dt.Columns.Add("Res3",typeof(int));

            // 動態年度欄
            string[] m={"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"};
            foreach(int y in years.OrderBy(x=>x))
            {
                foreach(string mon in m) dt.Columns.Add($"{y}_{mon}",typeof(double));
                dt.Columns.Add($"{y}_Q1_Total",typeof(double));
                dt.Columns.Add($"{y}_Q2_Total",typeof(double));
                dt.Columns.Add($"{y}_Q3_Total",typeof(double));
                dt.Columns.Add($"{y}_Q4_Total",typeof(double));
                dt.Columns.Add($"{y}_Total",typeof(double));
            }

            // 產生 20 筆示範資料
            var rnd=new Random(0);
            for(int i=1;i<=20;i++)
            {
                var r=dt.NewRow();
                r["Customer"]=$"Cust{(i%3)+1}";
                r["Inch"]    =$"{6+(i%3)}\"";
                r["CustItem"]=$"CI{(i%5)+1}";
                r["OAItem"]  =$"OA{(i%4)+1}";
                r["Substrate"]= (i%2==0)?"SubA":"SubB";
                r["Price"]   = 100+5*i;
                r["Thk1"]=1.0+0.1*(i%5); r["Thk2"]=1.2+0.1*(i%5); r["Thk3"]=1.4+0.1*(i%5);
                r["Res1"]=80+i; r["Res2"]=90+i; r["Res3"]=100+i;

                foreach(int y in years)
                {
                    double baseVal= rnd.Next(10,30);
                    r[$"{y}_Jan"]=baseVal;
                    r[$"{y}_Feb"]=baseVal+1;
                    r[$"{y}_Mar"]=baseVal+2;
                    r[$"{y}_Q1_Total"]=baseVal*3+3;

                    r[$"{y}_Apr"]=baseVal+3;
                    r[$"{y}_May"]=baseVal+4;
                    r[$"{y}_Jun"]=baseVal+5;
                    r[$"{y}_Q2_Total"]=baseVal*3+12;

                    r[$"{y}_Jul"]=baseVal+6;
                    r[$"{y}_Aug"]=baseVal+7;
                    r[$"{y}_Sep"]=baseVal+8;
                    r[$"{y}_Q3_Total"]=baseVal*3+21;

                    r[$"{y}_Oct"]=baseVal+9;
                    r[$"{y}_Nov"]=baseVal+10;
                    r[$"{y}_Dec"]=baseVal+11;
                    r[$"{y}_Q4_Total"]=baseVal*3+30;

                    r[$"{y}_Total"]= (double)r[$"{y}_Q1_Total"] +
                                     (double)r[$"{y}_Q2_Total"] +
                                     (double)r[$"{y}_Q3_Total"] +
                                     (double)r[$"{y}_Q4_Total"];
                }
                dt.Rows.Add(r);
            }
            return dt;
        }
    }
}
