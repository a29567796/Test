<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AnnualReport1.aspx.cs" Inherits="AnnualReport1" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Annual Report</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <style type="text/css">
        /* Example CSS for fixed columns/rows (adjust as needed) */
        .total-row { font-weight: bold; background-color: #EEF; }
        .total-col { font-weight: bold; background-color: #F0F0F0; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="query-panel">
            年份：
            <asp:DropDownList ID="ddlYear" runat="server"></asp:DropDownList>
            查詢方式：
            <asp:RadioButtonList ID="rdoQueryType" runat="server" RepeatDirection="Horizontal" AutoPostBack="true" 
                OnSelectedIndexChanged="rdoQueryType_SelectedIndexChanged">
                <asp:ListItem Text="月" Value="M" Selected="True" />
                <asp:ListItem Text="季" Value="Q" />
                <asp:ListItem Text="年" Value="Y" />
            </asp:RadioButtonList>
            <asp:Button ID="btnSearch" runat="server" Text="查詢" OnClick="btnSearch_Click" />
        </div>
        <!-- Table container for report -->
        <div id="reportContainer" style="height: 600px;">
            <!-- Fixed header table -->
            <table id="tblHeaderFixed" class="report-header-fixed" style="float: left;">
                <thead>
                    <tr>
                        <th>Customer</th>
                        <th>Inch</th>
                        <th>Cust Item</th>
                        <th>OA Item</th>
                        <th>Substrate</th>
                        <th>Price</th>
                        <th>Thk1</th>
                        <th>Res1</th>
                        <th>Thk2</th>
                        <th>Res2</th>
                        <th>Thk3</th>
                        <th>Res3</th>
                        <th>排程料號</th>
                    </tr>
                </thead>
            </table>
            <!-- Dynamic header table -->
            <table id="tblHeaderDyn" class="report-header-dyn" style="overflow: hidden;"></table>
            <!-- Body tables: fixed and dynamic, side by side -->
            <div style="clear: both; position: relative; height: 550px;">
                <!-- Fixed columns body (vertical scroll sync with dynamic) -->
                <div id="fixedContainer" style="float: left; overflow-y: auto; overflow-x: hidden; height: 100%;">
                    <table id="tblFixed" class="report-fixed">
                        <tbody>
                        </tbody>
                    </table>
                </div>
                <!-- Dynamic columns body (scrollable) -->
                <div id="dynamicContainer" style="overflow: auto; height: 100%;">
                    <table id="tblDynamic" class="report-dynamic">
                        <tbody>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <script type="text/javascript">
            // Global variables for data
            var fixedCount = 13;  // number of fixed columns
            var dynColumns = [];  // dynamic column name list
            var data = [];        // data rows (including fixed + dynamic values per row)
            var queryMode = 'M';  // current query mode ('M','Q','Y')

            function buildHeader() {
                // Build dynamic table header based on queryMode
                var thead;
                // Ensure the tblHeaderDyn has a THEAD element to append rows
                if (document.getElementById('tblHeaderDyn').getElementsByTagName('thead').length > 0) {
                    thead = document.getElementById('tblHeaderDyn').getElementsByTagName('thead')[0];
                } else {
                    thead = document.createElement('thead');
                    document.getElementById('tblHeaderDyn').appendChild(thead);
                }
                thead.innerHTML = '';  // clear existing header content
                if (queryMode === 'M') {
                    // Three-level header: Year, Quarter, Month+Subtotal
                    var rowYear = document.createElement('tr');
                    var rowQuarter = document.createElement('tr');
                    var rowMonth = document.createElement('tr');
                    // Group dynamic columns by year
                    var yearGroups = {};
                    for (var j = 0; j < dynColumns.length; j++) {
                        var col = dynColumns[j];
                        var year = col.substring(0, 4);
                        if (!yearGroups[year]) {
                            yearGroups[year] = [];
                        }
                        yearGroups[year].push(col);
                    }
                    for (var year in yearGroups) {
                        // Year header spanning all that year's dynamic columns
                        var cols = yearGroups[year];
                        var thYear = document.createElement('th');
                        thYear.colSpan = cols.length;
                        thYear.innerText = year;
                        rowYear.appendChild(thYear);
                        // Prepare quarter grouping within this year
                        var quarters = { 'Q1': [], 'Q2': [], 'Q3': [], 'Q4': [], 'Total': [] };
                        cols.forEach(function(colName) {
                            if (colName.indexOf('Total') !== -1) {
                                quarters['Total'].push(colName);
                            } else if (colName.indexOf('Q') !== -1 && colName.indexOf('Sum') !== -1) {
                                // Quarter subtotal column, e.g. "2021Q1Sum"
                                var qLabel = colName.substring(4, colName.indexOf('Sum'));  // "Q1"
                                quarters[qLabel].push(colName);
                            } else {
                                // Monthly column (format "YYYYMM")
                                var monthStr = colName.substring(4);  // e.g. "01", "02"
                                var monthNum = parseInt(monthStr, 10) % 100;
                                var qIndex = Math.floor((monthNum - 1) / 3) + 1;
                                quarters['Q' + qIndex].push(colName);
                            }
                        });
                        // Add quarter headers and month/subtotal headers for this year
                        var quarterLabels = ['Q1', 'Q2', 'Q3', 'Q4'];
                        quarterLabels.forEach(function(q) {
                            if (quarters[q] && quarters[q].length > 0) {
                                var thQ = document.createElement('th');
                                thQ.colSpan = quarters[q].length;
                                thQ.innerText = q;
                                rowQuarter.appendChild(thQ);
                                // Add month and subtotal columns under this quarter
                                quarters[q].forEach(function(colName) {
                                    var th = document.createElement('th');
                                    if (colName.indexOf('Sum') !== -1) {
                                        // Quarter subtotal column
                                        th.innerText = '合計';
                                    } else {
                                        // Month column - convert to month abbreviation
                                        var mStr = colName.substring(4);
                                        var mNum = parseInt(mStr, 10) % 100;
                                        var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                                        th.innerText = monthNames[mNum - 1] || mStr;
                                    }
                                    rowMonth.appendChild(th);
                                });
                            }
                        });
                        // Year total column (if exists) for this year group
                        if (quarters['Total'].length > 0) {
                            var thTotal = document.createElement('th');
                            thTotal.rowSpan = 2;
                            thTotal.innerText = '總計';
                            rowQuarter.appendChild(thTotal);
                            // Note: The year total column spans both quarter and month header rows (no separate month-level cell for it).
                        }
                    }
                    thead.appendChild(rowYear);
                    thead.appendChild(rowQuarter);
                    thead.appendChild(rowMonth);
                }
                else if (queryMode === 'Q') {
                    // Two-level header: Year, Quarter totals + Year Total
                    var rowYear = document.createElement('tr');
                    var rowQuarter = document.createElement('tr');
                    // Group columns by year
                    var yearGroups = {};
                    for (var k = 0; k < dynColumns.length; k++) {
                        var colName = dynColumns[k];
                        var yr = colName.substring(0, 4);
                        if (!yearGroups[yr]) {
                            yearGroups[yr] = [];
                        }
                        yearGroups[yr].push(colName);
                    }
                    for (var yr in yearGroups) {
                        var cols = yearGroups[yr].slice();
                        cols.sort();  // sort to ensure Q1..Q4 then Total
                        // Year header with colspan equal to number of columns for that year
                        var thYr = document.createElement('th');
                        thYr.colSpan = cols.length;
                        thYr.innerText = yr;
                        rowYear.appendChild(thYr);
                        // Quarter total and year total headers for this year
                        cols.forEach(function(colName) {
                            var th = document.createElement('th');
                            if (colName.indexOf('Total') !== -1) {
                                th.innerText = '總計';
                            } else if (colName.indexOf('Q') !== -1 && colName.indexOf('Sum') !== -1) {
                                var qLabel = colName.substring(4, colName.indexOf('Sum'));  // e.g. "Q1"
                                th.innerText = qLabel + '合計';
                            } else {
                                th.innerText = colName;
                            }
                            rowQuarter.appendChild(th);
                        });
                    }
                    thead.appendChild(rowYear);
                    thead.appendChild(rowQuarter);
                }
                else if (queryMode === 'Y') {
                    // One-level header: Year Total only
                    var row = document.createElement('tr');
                    if (dynColumns.length > 1) {
                        // Multiple years: label each column by year total
                        dynColumns.forEach(function(colName) {
                            var th = document.createElement('th');
                            if (colName.indexOf('Total') !== -1) {
                                var yr = colName.substring(0, 4);
                                th.innerText = yr + '總計';
                            } else {
                                th.innerText = colName;
                            }
                            row.appendChild(th);
                        });
                    } else if (dynColumns.length === 1) {
                        // Single column (one year total)
                        var th = document.createElement('th');
                        th.innerText = '總計';
                        row.appendChild(th);
                    }
                    thead.appendChild(row);
                }
            }

            function buildBody() {
                // Build table body rows (fixed and dynamic parts) and add total row
                var tbodyFixed = document.getElementById('tblFixed').getElementsByTagName('tbody')[0];
                var tbodyDyn = document.getElementById('tblDynamic').getElementsByTagName('tbody')[0];
                tbodyFixed.innerHTML = '';
                tbodyDyn.innerHTML = '';
                for (var i = 0; i < data.length; i++) {
                    var rowF = document.createElement('tr');
                    var rowD = document.createElement('tr');
                    // Fixed cells
                    for (var f = 0; f < fixedCount; f++) {
                        var cellF = document.createElement('td');
                        cellF.innerText = data[i][f] !== null ? data[i][f] : '';
                        rowF.appendChild(cellF);
                    }
                    // Dynamic cells
                    for (var j = 0; j < dynColumns.length; j++) {
                        var cellD = document.createElement('td');
                        var val = data[i][fixedCount + j];
                        cellD.innerText = val !== null ? val : '';
                        if (dynColumns[j].indexOf('Total') !== -1) {
                            cellD.className = 'total-col';
                        }
                        rowD.appendChild(cellD);
                    }
                    tbodyFixed.appendChild(rowF);
                    tbodyDyn.appendChild(rowD);
                }
                // Append totals row
                var totalRowF = document.createElement('tr');
                var totalRowD = document.createElement('tr');
                for (var ff = 0; ff < fixedCount; ff++) {
                    var cellF = document.createElement('td');
                    if (ff === 0) {
                        cellF.innerText = '總計';
                    } else {
                        cellF.innerText = '';
                    }
                    totalRowF.appendChild(cellF);
                }
                for (var jj = 0; jj < dynColumns.length; jj++) {
                    var cellD = document.createElement('td');
                    if (dynColumns[jj].indexOf('Total') !== -1) {
                        cellD.className = 'total-col';
                    }
                    cellD.innerText = '';
                    totalRowD.appendChild(cellD);
                }
                totalRowF.className = 'total-row';
                totalRowD.className = 'total-row';
                tbodyFixed.appendChild(totalRowF);
                tbodyDyn.appendChild(totalRowD);
            }

            function updateTotals() {
                // Calculate column totals and update the totals row
                var tbodyDyn = document.getElementById('tblDynamic').getElementsByTagName('tbody')[0];
                var rows = tbodyDyn.getElementsByTagName('tr');
                if (rows.length === 0) return;
                var lastRowIndex = rows.length - 1;
                var totalCells = rows[lastRowIndex].getElementsByTagName('td');
                if (totalCells.length !== dynColumns.length) return;
                // Initialize sums
                var sums = new Array(dynColumns.length).fill(0);
                // Sum each dynamic column across all data rows (excluding total row)
                for (var i = 0; i < data.length; i++) {
                    for (var j = 0; j < dynColumns.length; j++) {
                        var val = parseFloat(data[i][fixedCount + j]);
                        if (!isNaN(val)) {
                            sums[j] += val;
                        }
                    }
                }
                // Update total row cells
                for (var k = 0; k < dynColumns.length; k++) {
                    totalCells[k].innerText = (sums[k] !== 0 ? sums[k] : 0);
                }
            }

            // (Optional) Synchronize scrolling of fixed and dynamic containers
            document.getElementById('dynamicContainer').onscroll = function() {
                document.getElementById('fixedContainer').scrollTop = this.scrollTop;
                document.getElementById('tblHeaderDyn').scrollLeft = this.scrollLeft;
            };
        </script>
    </form>
</body>
</html>