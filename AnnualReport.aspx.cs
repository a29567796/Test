using System;
using System.Data;
using System.Web.UI.WebControls;
using OfficeOpenXml;

public partial class AnnualReport : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            // Optionally set a default year (e.g., current year) in the year input
            // txtYear.Text = DateTime.Now.Year.ToString();
        }
    }

    /// <summary>
    /// Generate dummy report data for the given year, with 20 records and all required columns.
    /// </summary>
    private DataTable GenerateData(int year)
    {
        DataTable dt = new DataTable();
        // 1. Define columns for new fixed fields and existing ones
        dt.Columns.Add("Customer", typeof(string));
        dt.Columns.Add("Inch", typeof(string));      // use string to preserve format like "2.0"
        dt.Columns.Add("CustItem", typeof(string));
        dt.Columns.Add("OAItem", typeof(string));
        dt.Columns.Add("Substrate", typeof(string));
        dt.Columns.Add("Price", typeof(double));
        dt.Columns.Add("Thk1", typeof(int));
        dt.Columns.Add("Res1", typeof(int));
        dt.Columns.Add("Thk2", typeof(int));
        dt.Columns.Add("Res2", typeof(int));
        dt.Columns.Add("Thk3", typeof(int));
        dt.Columns.Add("Res3", typeof(int));
        dt.Columns.Add("Schedule", typeof(string));
        // Month columns M01–M12
        for (int m = 1; m <= 12; m++)
        {
            dt.Columns.Add("M" + m.ToString("D2"), typeof(int));
        }
        // Quarter sums and Year total
        dt.Columns.Add("Q1Sum", typeof(int));
        dt.Columns.Add("Q2Sum", typeof(int));
        dt.Columns.Add("Q3Sum", typeof(int));
        dt.Columns.Add("Q4Sum", typeof(int));
        dt.Columns.Add("YearTotal", typeof(int));

        // Sample data generation for 20 rows
        string[] customerOptions = { "VI", "AP", "ART", "LAX", "CP", "MX" };
        string[] custItemOptions = { "CI01", "CI02", "CI03", "CI99" };
        string[] oaItemOptions   = { "OA99", "OA88", "OA77", "OA66" };
        string[] substrateOptions = { "SubA", "SubB", "SubC" };
        Random rand = new Random();
        for (int i = 1; i <= 20; i++)
        {
            string customer = customerOptions[rand.Next(customerOptions.Length)];
            // Alternate inch format between integer and one-decimal (for example)
            string inch = (i % 2 == 0) ? "2.0" : "8";
            string custItem = custItemOptions[rand.Next(custItemOptions.Length)];
            string oaItem = oaItemOptions[rand.Next(oaItemOptions.Length)];
            string substrate = substrateOptions[rand.Next(substrateOptions.Length)];
            double price = Math.Round(300 + rand.NextDouble() * 1000, 2);  // random price between 300 and 1300

            // Random Thk/Res values (0–100)
            int thk1 = rand.Next(0, 101);
            int thk2 = rand.Next(0, 101);
            int thk3 = rand.Next(0, 101);
            int res1 = rand.Next(0, 101);
            int res2 = rand.Next(0, 101);
            int res3 = rand.Next(0, 101);

            string schedule = "";  // leave Schedule blank or use placeholder if needed

            // Random monthly values (0–1000) and compute quarter sums
            int[] months = new int[12];
            for (int m = 0; m < 12; m++)
            {
                months[m] = rand.Next(0, 1001);
            }
            int q1Sum = months[0] + months[1] + months[2];
            int q2Sum = months[3] + months[4] + months[5];
            int q3Sum = months[6] + months[7] + months[8];
            int q4Sum = months[9] + months[10] + months[11];
            int yearTotal = q1Sum + q2Sum + q3Sum + q4Sum;

            // Create a new row with all values
            DataRow row = dt.NewRow();
            row["Customer"] = customer;
            row["Inch"] = inch;
            row["CustItem"] = custItem;
            row["OAItem"] = oaItem;
            row["Substrate"] = substrate;
            row["Price"] = price;
            row["Thk1"] = thk1;
            row["Res1"] = res1;
            row["Thk2"] = thk2;
            row["Res2"] = res2;
            row["Thk3"] = thk3;
            row["Res3"] = res3;
            row["Schedule"] = schedule;
            // Month values
            for (int m = 1; m <= 12; m++)
            {
                row["M" + m.ToString("D2")] = months[m - 1];
            }
            // Totals
            row["Q1Sum"] = q1Sum;
            row["Q2Sum"] = q2Sum;
            row["Q3Sum"] = q3Sum;
            row["Q4Sum"] = q4Sum;
            row["YearTotal"] = yearTotal;

            dt.Rows.Add(row);
        }
        return dt;
    }

    protected void btnQuery_Click(object sender, EventArgs e)
    {
        // 5. Front-end query triggers data generation and binding
        int year;
        if (!int.TryParse(txtYear.Text.Trim(), out year))
        {
            year = DateTime.Now.Year; // default to current year if input not valid
        }

        DataTable dt = GenerateData(year);
        // Bind data to the repeater (table body)
        rptData.DataSource = dt;
        rptData.DataBind();

        // Set header text for year and months
        litYear.Text = year.ToString();
        litYearTotal.Text = year + "合計";
        litM01.Text = year + "01"; litM02.Text = year + "02"; litM03.Text = year + "03";
        litM04.Text = year + "04"; litM05.Text = year + "05"; litM06.Text = year + "06";
        litM07.Text = year + "07"; litM08.Text = year + "08"; litM09.Text = year + "09";
        litM10.Text = year + "10"; litM11.Text = year + "11"; litM12.Text = year + "12";

        // Store data in Session for use during export (to avoid regeneration on export)
        Session["AnnualReportData"] = dt;
        Session["SelectedYear"] = year;
    }

    protected void btnExport_Click(object sender, EventArgs e)
    {
        // 6. EPPlus export – include new columns and headers
        DataTable dt = Session["AnnualReportData"] as DataTable;
        int year = (Session["SelectedYear"] != null) ? (int)Session["SelectedYear"] : DateTime.Now.Year;
        if (dt == null)
        {
            // If no session data (e.g., user directly clicked export), generate data on the fly
            dt = GenerateData(year);
        }

        // Prepare Excel package
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  // EPPlus license setup8
        using (ExcelPackage pck = new ExcelPackage())
        {
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("AnnualReport");

            // Build the multi-row header in Excel, matching the web table structure:
            // First row: blank spanning fixed columns + year spanning 17 columns
            ws.Cells[1, 1].Value = string.Empty;
            ws.Cells[1, 1, 1, 13].Merge = true;
            ws.Cells[1, 14].Value = year.ToString();
            ws.Cells[1, 14, 1, 30].Merge = true;

            // Second row: blank spanning fixed + quarter group headers + year total (merged down)
            ws.Cells[2, 1].Value = string.Empty;
            ws.Cells[2, 1, 2, 13].Merge = true;
            ws.Cells[2, 14].Value = "Q1";
            ws.Cells[2, 14, 2, 17].Merge = true;
            ws.Cells[2, 18].Value = "Q2";
            ws.Cells[2, 18, 2, 21].Merge = true;
            ws.Cells[2, 22].Value = "Q3";
            ws.Cells[2, 22, 2, 25].Merge = true;
            ws.Cells[2, 26].Value = "Q4";
            ws.Cells[2, 26, 2, 29].Merge = true;
            ws.Cells[2, 30].Value = year + "合計";
            ws.Cells[2, 30, 3, 30].Merge = true;  // merge year total header over second & third row

            // Third row: fixed headers + month and quarter-total headers
            string[] fixedHeaders = { "Customer", "Inch", "Cust Item", "OA Item", "Substrate", "Price",
                                      "Thk1", "Res1", "Thk2", "Res2", "Thk3", "Res3", "排程列表" };
            for (int col = 1; col <= fixedHeaders.Length; col++)
            {
                ws.Cells[3, col].Value = fixedHeaders[col - 1];
            }
            // Month labels and quarter "合計" labels under each quarter group
            int colIndex = 14;
            for (int q = 1; q <= 4; q++)
            {
                // Months for quarter q
                for (int m = 1; m <= 3; m++)
                {
                    int monthNum = (q - 1) * 3 + m;
                    string monthLabel = year.ToString() + monthNum.ToString("D2");
                    ws.Cells[3, colIndex++].Value = monthLabel;
                }
                // Quarter sum label
                ws.Cells[3, colIndex++].Value = "合計";
            }
            // (Year total column header is already handled by the merged cell above)

            // Fill data rows starting from row 4
            int startRow = 4;
            foreach (DataRow dr in dt.Rows)
            {
                int c = 1;
                // Fixed columns
                ws.Cells[startRow, c++].Value = dr["Customer"];
                ws.Cells[startRow, c++].Value = dr["Inch"];
                ws.Cells[startRow, c++].Value = dr["CustItem"];
                ws.Cells[startRow, c++].Value = dr["OAItem"];
                ws.Cells[startRow, c++].Value = dr["Substrate"];
                ws.Cells[startRow, c++].Value = Convert.ToDouble(dr["Price"]);
                ws.Cells[startRow, c++].Value = dr["Thk1"];
                ws.Cells[startRow, c++].Value = dr["Res1"];
                ws.Cells[startRow, c++].Value = dr["Thk2"];
                ws.Cells[startRow, c++].Value = dr["Res2"];
                ws.Cells[startRow, c++].Value = dr["Thk3"];
                ws.Cells[startRow, c++].Value = dr["Res3"];
                ws.Cells[startRow, c++].Value = dr["Schedule"];
                // Month columns
                for (int m = 1; m <= 12; m++)
                {
                    ws.Cells[startRow, c++].Value = dr["M" + m.ToString("D2")];
                }
                // Quarter sums and year total
                ws.Cells[startRow, c++].Value = dr["Q1Sum"];
                ws.Cells[startRow, c++].Value = dr["Q2Sum"];
                ws.Cells[startRow, c++].Value = dr["Q3Sum"];
                ws.Cells[startRow, c++].Value = dr["Q4Sum"];
                ws.Cells[startRow, c++].Value = dr["YearTotal"];
                startRow++;
            }

            // Optional: adjust column widths (here we set a uniform width similar to 100px)
            for (int col = 1; col <= 30; col++)
            {
                ws.Column(col).Width = 15;  // roughly corresponds to 100px width
            }

            // Send the Excel file to the client
            string filename = $"AnnualReport_{year}.xlsx";
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + filename);
            Response.BinaryWrite(pck.GetAsByteArray());
            Response.End();
        }
    }
}
