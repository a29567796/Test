using System;
using System.Data;
using System.Collections.Generic;
using Newtonsoft.Json;

public partial class AnnualReport1 : System.Web.UI.Page
{
    private const int FixedColumnCount = 13;  // number of fixed (non-dynamic) columns

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            rdoQueryType.SelectedValue = "M";  // default to Month mode
            // TODO: Initialize ddlYear options, e.g. bind available years
        }
    }

    protected void btnSearch_Click(object sender, EventArgs e)
    {
        // Get selected filters
        string selectedYear = ddlYear.SelectedValue;
        string mode = rdoQueryType.SelectedValue;  // "M", "Q", or "Y"

        // Retrieve data for the selected year (and possibly multiple years if needed)
        DataTable dt = GetAnnualReportData(selectedYear);
        if (dt == null || dt.Rows.Count == 0)
        {
            // No data: clear any existing table content on client
            string clearScript = "<script>document.getElementById('tblHeaderDyn').getElementsByTagName('thead')[0].innerHTML='';" +
                                 "document.getElementById('tblFixed').getElementsByTagName('tbody')[0].innerHTML='';" +
                                 "document.getElementById('tblDynamic').getElementsByTagName('tbody')[0].innerHTML='';</script>";
            ClientScript.RegisterStartupScript(this.GetType(), "nodata", clearScript, false);
            return;
        }

        // Compute quarter subtotal columns and annual total column
        HashSet<string> years = new HashSet<string>();
        for (int i = FixedColumnCount; i < dt.Columns.Count; i++)
        {
            string colName = dt.Columns[i].ColumnName;
            if (colName.Length >= 4 && char.IsDigit(colName[0]) && char.IsDigit(colName[1]) 
                && char.IsDigit(colName[2]) && char.IsDigit(colName[3]))
            {
                string yr = colName.Substring(0, 4);
                years.Add(yr);
            }
        }
        // Add quarter sum and total columns for each year group
        foreach (string yr in years)
        {
            string q1Col = yr + "Q1Sum";
            string q2Col = yr + "Q2Sum";
            string q3Col = yr + "Q3Sum";
            string q4Col = yr + "Q4Sum";
            string totalCol = yr + "Total";
            // Insert Q1 sum column after month "03" of that year (or at group start if no month columns present for Q1)
            int q1Index;
            if (dt.Columns.Contains(yr + "03"))
            {
                q1Index = dt.Columns.IndexOf(yr + "03") + 1;
            }
            else
            {
                // find first column of this year that comes after Q1 period (month >=04)
                q1Index = FindInsertIndex(dt, yr, 1, 3);
                if (q1Index < FixedColumnCount) q1Index = FixedColumnCount;
            }
            dt.Columns.Add(q1Col, typeof(decimal)).SetOrdinal(q1Index);

            // Insert Q2 sum after month "06"
            int q2Index;
            if (dt.Columns.Contains(yr + "06"))
            {
                q2Index = dt.Columns.IndexOf(yr + "06") + 1;
            }
            else
            {
                q2Index = FindInsertIndex(dt, yr, 4, 6);
                if (q2Index < FixedColumnCount) q2Index = dt.Columns.Count;
            }
            dt.Columns.Add(q2Col, typeof(decimal)).SetOrdinal(q2Index);

            // Insert Q3 sum after month "09"
            int q3Index;
            if (dt.Columns.Contains(yr + "09"))
            {
                q3Index = dt.Columns.IndexOf(yr + "09") + 1;
            }
            else
            {
                q3Index = FindInsertIndex(dt, yr, 7, 9);
                if (q3Index < FixedColumnCount) q3Index = dt.Columns.Count;
            }
            dt.Columns.Add(q3Col, typeof(decimal)).SetOrdinal(q3Index);

            // Insert Q4 sum after month "12"
            int q4Index;
            if (dt.Columns.Contains(yr + "12"))
            {
                q4Index = dt.Columns.IndexOf(yr + "12") + 1;
            }
            else
            {
                // if no Dec column, insert at end of that year group (before next year's data or at end of table)
                q4Index = FindYearGroupEndIndex(dt, yr);
            }
            dt.Columns.Add(q4Col, typeof(decimal)).SetOrdinal(q4Index);

            // Insert Total column after Q4 sum
            dt.Columns.Add(totalCol, typeof(decimal));
            int totIndex = dt.Columns.IndexOf(q4Col) + 1;
            dt.Columns[totalCol].SetOrdinal(totIndex);

            // Calculate values for the new columns for each row
            foreach (DataRow dr in dt.Rows)
            {
                decimal sumQ1 = 0, sumQ2 = 0, sumQ3 = 0, sumQ4 = 0;
                // Q1: Jan(01) - Mar(03)
                for (int m = 1; m <= 3; m++)
                {
                    string mCol = yr + m.ToString("D2");
                    if (dt.Columns.Contains(mCol) && dr[mCol] != DBNull.Value)
                        sumQ1 += Convert.ToDecimal(dr[mCol]);
                }
                dr[q1Col] = sumQ1;
                // Q2: Apr(04) - Jun(06)
                for (int m = 4; m <= 6; m++)
                {
                    string mCol = yr + m.ToString("D2");
                    if (dt.Columns.Contains(mCol) && dr[mCol] != DBNull.Value)
                        sumQ2 += Convert.ToDecimal(dr[mCol]);
                }
                dr[q2Col] = sumQ2;
                // Q3: Jul(07) - Sep(09)
                for (int m = 7; m <= 9; m++)
                {
                    string mCol = yr + m.ToString("D2");
                    if (dt.Columns.Contains(mCol) && dr[mCol] != DBNull.Value)
                        sumQ3 += Convert.ToDecimal(dr[mCol]);
                }
                dr[q3Col] = sumQ3;
                // Q4: Oct(10) - Dec(12)
                for (int m = 10; m <= 12; m++)
                {
                    string mCol = yr + m.ToString("D2");
                    if (dt.Columns.Contains(mCol) && dr[mCol] != DBNull.Value)
                        sumQ4 += Convert.ToDecimal(dr[mCol]);
                }
                dr[q4Col] = sumQ4;
                // Year total = sum of four quarters
                dr[totalCol] = sumQ1 + sumQ2 + sumQ3 + sumQ4;
            }
        }

        // Remove unwanted columns based on query mode
        if (mode == "Q")
        {
            // Remove all monthly columns, keep only quarter sums and totals
            List<string> removeCols = new List<string>();
            for (int i = FixedColumnCount; i < dt.Columns.Count; i++)
            {
                string name = dt.Columns[i].ColumnName;
                if (!(name.Contains("Q") || name.Contains("Total")))
                {
                    removeCols.Add(name);
                }
            }
            foreach (string colName in removeCols)
            {
                if (dt.Columns.Contains(colName)) dt.Columns.Remove(colName);
            }
        }
        else if (mode == "Y")
        {
            // Remove all monthly and quarterly columns, keep only annual Total columns
            List<string> removeCols = new List<string>();
            for (int i = FixedColumnCount; i < dt.Columns.Count; i++)
            {
                string name = dt.Columns[i].ColumnName;
                if (!name.Contains("Total"))
                {
                    removeCols.Add(name);
                }
            }
            foreach (string colName in removeCols)
            {
                if (dt.Columns.Contains(colName)) dt.Columns.Remove(colName);
            }
        }

        // Prepare dynamic column names list and data for client-side
        List<string> dynCols = new List<string>();
        for (int i = FixedColumnCount; i < dt.Columns.Count; i++)
        {
            dynCols.Add(dt.Columns[i].ColumnName);
        }
        // Prepare data rows (including fixed and dynamic columns)
        var dataList = new List<object>();
        foreach (DataRow dr in dt.Rows)
        {
            var rowData = new List<object>();
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                object val = dr[c];
                // For numeric values, you might want to format or cast as needed (here we keep raw)
                rowData.Add(val is DBNull ? null : val);
            }
            dataList.Add(rowData);
        }
        // Serialize to JSON
        string jsonCols = JsonConvert.SerializeObject(dynCols);
        string jsonData = JsonConvert.SerializeObject(dataList);
        // Register startup script to set data and build table on client side
        string script = "<script>" +
                        "var dynColumns = " + jsonCols + ";" +
                        "var data = " + jsonData + ";" +
                        "var queryMode = '" + mode + "';" +
                        "buildHeader();" +
                        "buildBody();" +
                        "updateTotals();" +
                        "</script>";
        ClientScript.RegisterStartupScript(this.GetType(), "buildReport", script, false);
    }

    protected void rdoQueryType_SelectedIndexChanged(object sender, EventArgs e)
    {
        // Automatically trigger search with the same filters when query type changes
        btnSearch_Click(sender, e);
    }

    /// <summary>
    /// Retrieves the annual report data for the given year.
    /// (This method should be implemented to fetch data from database or other source.)
    /// </summary>
    private DataTable GetAnnualReportData(string year)
    {
        // TODO: Implement data retrieval for the selected year.
        // This should return a DataTable with fixed columns and one column per month (YYYYMM) for the given year(s).
        // For example, columns: Customer, Inch, ..., Res3, 排程料號, 202101, 202102, ..., 202112 (if one year).
        throw new NotImplementedException();
    }

    /// <summary>
    /// Finds the insert index for a quarter sum column when monthly columns for that quarter may be missing.
    /// Returns the index to insert the quarter sum column for quarter defined by monthStart to monthEnd.
    /// </summary>
    private int FindInsertIndex(DataTable dt, string year, int monthStart, int monthEnd)
    {
        // Find the first dynamic column of the given year that comes after the specified quarter range.
        for (int j = FixedColumnCount; j < dt.Columns.Count; j++)
        {
            string name = dt.Columns[j].ColumnName;
            if (!name.StartsWith(year)) continue;
            if (name.Length == 6)
            {
                // Month column of same year
                int m = int.Parse(name.Substring(4));
                if (m >= monthEnd + 1)
                {
                    return j;
                }
            }
            else if (name.Contains("Q") || name.Contains("Total"))
            {
                // A quarter sum or total column of the same year (which would appear after all month columns of that year)
                if (name.StartsWith(year))
                {
                    return j;
                }
            }
        }
        // If no later column found, return table end index (which will append at end of current columns)
        return dt.Columns.Count;
    }

    /// <summary>
    /// Finds the end index of the column group for the specified year (position to insert Q4 sum or Total at end of that year).
    /// </summary>
    private int FindYearGroupEndIndex(DataTable dt, string year)
    {
        // Find the first column of the next year (to know where current year group ends)
        for (int j = FixedColumnCount; j < dt.Columns.Count; j++)
        {
            string name = dt.Columns[j].ColumnName;
            if (name.Length >= 4 && String.Compare(name.Substring(0, 4), year) > 0)
            {
                return j;
            }
        }
        // If this is the last year or no larger year found, return current end of table (will insert at end)
        return dt.Columns.Count;
    }
}
