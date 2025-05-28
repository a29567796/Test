<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AnnualReport.aspx.cs" Inherits="CRM.AnnualReport" %>

<!DOCTYPE html>
<html lang="zh-TW">
<head runat="server">
    <meta charset="utf-8" />
    <title>年度查詢報表</title>

    <!-- Bootstrap & jQuery -->
    <link  rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css"/>
    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>

    <!-- jQuery UI：日期選擇器 -->
    <link  rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css"/>
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

    <style>
        /* ---------- Sticky 表頭 ---------- */
        thead tr.totals-row   th{position:sticky;top:0;   z-index:4;background:#fff;color:#007bff;}
        thead tr:nth-child(2) th{position:sticky;top:48px;z-index:3;background:#f2f2f2;}
        thead tr:nth-child(3) th{position:sticky;top:96px;z-index:2;background:#fafafa;}
        thead tr:nth-child(4) th{position:sticky;top:144px;z-index:1;background:#fff;}

        /* 固定欄位：空白 / 標題 */
        .fc-blank{background:#fff!important; border:0px!important;}
        .fc      {background:#f2f2f2!important;}

        /* 年度色塊 */
        .year-0{background:#92D050 !important;}
        .year-1{background:#00B0F0 !important;}
        .year-2{background:#FFC000 !important;}
        .year-3{background:#A9A9A9 !important;}

        /* 捲動區 */
        .table-responsive{max-height:700px;overflow:auto;}

        /* 固定寬 */
        .w120{width:120px;}.w100{width:100px;}.w80{width:80px;}.w60{width:60px;}

        .table thead th{vertical-align:bottom;border-bottom:0;}
        .table-bordered td,.table-bordered th{border:0;}
        .table-sm td,.table-sm th{padding:.75rem;}

        .table tbody tr{border-bottom:1px solid;}
    </style>
</head>
<body>
<form id="form1" runat="server" class="container-fluid my-3">

<!-- ===== 查詢條件 ===== -->
<div class="card mb-3">
  <div class="card-header">查詢條件</div>
  <div class="card-body">

    <!-- 第一列 -->
    <div class="form-row">
      <div class="form-group col-md-2">
        <label>Department</label>
        <select id="ddlDept" class="form-control">
          <option value="">All</option><option>S100</option><option>S200</option>
        </select>
      </div>
      <div class="form-group col-md-2">
        <label>Business</label>
        <select id="ddlBiz" class="form-control">
          <option value="">All</option><option>Alice</option><option>Bob</option>
        </select>
      </div>
      <div class="form-group col-md-2">
        <label>Region</label>
        <select id="ddlRegion" class="form-control">
          <option value="">All</option>
          <option value="TW">Taiwan</option><option value="CN">China + HK</option><option value="JP">Japan</option>
        </select>
      </div>
      <div class="form-group col-md-3">
        <label>Year (多選)</label>
        <div>
          <label class="mr-2"><input type="checkbox" name="year" value="2022"/>2022</label>
          <label class="mr-2"><input type="checkbox" name="year" value="2023"/>2023</label>
          <label class="mr-2"><input type="checkbox" name="year" value="2024"/>2024</label>
          <label class="mr-2"><input type="checkbox" name="year" value="2025"/>2025</label>
        </div>
      </div>
      <div class="form-group col-md-3">
        <label>Customer</label>
        <select id="ddlCustomer" class="form-control">
          <option value="">All</option><option>VI</option><option>EPS</option>
        </select>
      </div>
    </div>

    <!-- 第二列 -->
    <div class="form-row">
      <div class="form-group col-md-2">
        <label>Cust&nbsp;Item</label><input id="txtCustItem" class="form-control"/>
      </div>
      <div class="form-group col-md-2">
        <label>Substrate</label>
        <select id="ddlSubstrate" class="form-control">
          <option value="">All</option><option>SubA</option><option>SubB</option>
        </select>
      </div>
      <div class="form-group col-md-2">
        <label>Products&nbsp;type</label>
        <select id="ddlProductType" class="form-control">
          <option value="">All</option><option>GaN</option><option>SiC</option>
        </select>
      </div>
      <div class="form-group col-md-2">
        <label>Ship&nbsp;to</label><input id="txtShipTo" class="form-control" placeholder="e.g. Taiwan"/>
      </div>
      <div class="form-group col-md-2">
        <label>Real&nbsp;Date</label><input id="txtRealDate" class="form-control datepicker" placeholder="yyyy/MM/dd"/>
      </div>
      <div class="form-group col-md-1">
        <label>Status</label>
        <select id="ddlStatus" class="form-control">
          <option value="">All</option><option>Open</option><option>Closed</option>
        </select>
      </div>
      <div class="form-group col-md-1">
        <label>Order&nbsp;Type</label>
        <select id="ddlOrderType" class="form-control">
          <option value="">All</option><option>3x01</option><option>3x02</option>
        </select>
      </div>
    </div>

    <!-- 第三列：月 / 季 / 年 -->
    <div class="form-row mb-2">
      <div class="form-group col-md-12">
        <label class="mr-2">顯示範圍：</label>
        <label class="mr-2"><input type="radio" name="viewType" value="month" checked/>月</label>
        <label class="mr-2"><input type="radio" name="viewType" value="quarter"/>季</label>
        <label class="mr-2"><input type="radio" name="viewType" value="year"/>年</label>
      </div>
    </div>

    <!-- 操作按鈕 -->
    <button id="btnQuery" type="button" class="btn btn-primary mr-2">查詢</button>
    <asp:Button ID="btnExport" runat="server" CssClass="btn btn-success" Text="匯出 Excel" OnClick="btnExport_Click" />

    <!-- Thk/Res 欄位切換 -->
    <div class="mt-3">
      <label class="mr-2">顯示欄位：</label>
      <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk1" checked/>Thk1</label>
      <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk2" checked/>Thk2</label>
      <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk3" checked/>Thk3</label>
      <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res1" checked/>Res1</label>
      <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res2" checked/>Res2</label>
      <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res3" checked/>Res3</label>
    </div>

  </div>
</div>

<!-- ===== 報表 ===== -->
<div class="table-responsive">
  <table id="resultTable" class="table table-bordered table-sm">
    <thead></thead>
    <tbody></tbody>
  </table>
</div>
</form>

<script>
    $(function () {
        /* 日期 */
        $(".datepicker").datepicker({ dateFormat: "yy/mm/dd" });

        /* 預選今年 */
        const y = new Date().getFullYear().toString();
        $(`input[name=year][value=${y}]`).prop("checked", true);

        /* 事件 */
        $("#btnQuery").on("click", queryReport);
        $(".col-toggle").on("change", function () {
            const cls = $(this).val();
            $("th." + cls + ", td." + cls).toggle($(this).is(":checked"));
        });
    });

    /* ---------- AJAX ---------- */
    function queryReport() {
        const years = $("input[name=year]:checked").map((_, el) => +el.value).get();
        if (!years.length) { alert("請至少選一年"); return; }
        const viewType = $("input[name=viewType]:checked").val(); // month / quarter / year

        $.ajax({
            type: "POST",
            url: "AnnualReport.aspx/GetReportData",
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            data: JSON.stringify({
                department: $("#ddlDept").val(),
                business: $("#ddlBiz").val(),
                region: $("#ddlRegion").val(),
                yearList: years,
                customer: $("#ddlCustomer").val(),
                custItem: $("#txtCustItem").val(),
                substrate: $("#ddlSubstrate").val(),
                productType: $("#ddlProductType").val(),
                shipTo: $("#txtShipTo").val(),
                realDate: $("#txtRealDate").val(),
                status: $("#ddlStatus").val(),
                orderType: $("#ddlOrderType").val()
            }),
            success: r => {
                const data = r.d || r;
                buildHeader(years, viewType);
                buildBody(data, years, viewType);
                updateTotals();
                $(".col-toggle").each(function () {
                    if (!$(this).is(":checked"))
                        $("th." + this.value + ", td." + this.value).hide();
                });
            },
            error: e => { console.error(e); alert("查詢失敗"); }
        });
    }

    /* ---------- 固定欄 ---------- */
    const fixedCols = [
        { t: "Customer", c: "col-customer", w: "w120" },
        { t: "Inch", c: "col-inch", w: "w60" },
        { t: "Cust&nbsp;Item", c: "col-custitem", w: "w120" },
        { t: "OA&nbsp;Item", c: "col-oaitem", w: "w120" },
        { t: "Substrate", c: "col-substrate", w: "w100" },
        { t: "Price", c: "col-price", w: "w80" },
        { t: "Thk1", c: "col-thk1", w: "w80" },
        { t: "Res1", c: "col-res1", w: "w80" },
        { t: "Thk2", c: "col-thk2", w: "w80" },
        { t: "Res2", c: "col-res2", w: "w80" },
        { t: "Thk3", c: "col-thk3", w: "w80" },
        { t: "Res3", c: "col-res3", w: "w80" }
    ];

    /* ---------- 表頭 ---------- */
    function buildHeader(years, viewType) {
        const $thead = $("#resultTable thead").empty();

        /* row0：Totals (空) */
        let row0 = "<tr class='totals-row'>";
        fixedCols.forEach(fc => row0 += `<th class="${fc.c} fc-blank ${fc.w}"></th>`);
        const colPerYear = (viewType === "month") ? 17 : (viewType === "quarter") ? 5 : 1;
        years.forEach(() => { for (let i = 0; i < colPerYear; i++)row0 += "<th></th>"; });
        row0 += "</tr>";
        $thead.append(row0);

        const fcBlank = () => fixedCols.map(fc => `<th class="${fc.c} fc-blank ${fc.w}"></th>`).join("");
        const fcTitle = () => fixedCols.map(fc => `<th class="${fc.c} fc ${fc.w}">${fc.t}</th>`).join("");

        /* 月 / 季 / 年 個別處理 */
        if (viewType === "month") {
            /* row1：年 */
            let r1 = "<tr>" + fcBlank();
            years.forEach((y, i) => r1 += `<th colspan="17" class="year-${i % 4}">${y}</th>`);
            r1 += "</tr>";
            $thead.append(r1);

            /* row2：季 */
            let r2 = "<tr>" + fcBlank();
            years.forEach(() => r2 += "<th colspan='4'>Q1</th><th colspan='4'>Q2</th><th colspan='4'>Q3</th><th colspan='4'>Q4</th><th rowspan='2'>Total</th>");
            r2 += "</tr>";
            $thead.append(r2);

            /* row3：固定欄 + 月 */
            let r3 = "<tr>" + fcTitle();
            years.forEach(() => r3 += `
      <th>Jan</th><th>Feb</th><th>Mar</th><th>Q1&nbsp;T</th>
      <th>Apr</th><th>May</th><th>Jun</th><th>Q2&nbsp;T</th>
      <th>Jul</th><th>Aug</th><th>Sep</th><th>Q3&nbsp;T</th>
      <th>Oct</th><th>Nov</th><th>Dec</th><th>Q4&nbsp;T</th>`);
            r3 += "</tr>";
            $thead.append(r3);

        } else if (viewType === "quarter") {
            /* row1：年 */
            let r1 = "<tr>" + fcBlank();
            years.forEach((y, i) => r1 += `<th colspan="5" class="year-${i % 4}">${y}</th>`);
            r1 += "</tr>";
            $thead.append(r1);

            /* row2：固定欄 + 季總 */
            let r2 = "<tr>" + fcTitle();
            years.forEach(() => r2 += "<th>Q1&nbsp;T</th><th>Q2&nbsp;T</th><th>Q3&nbsp;T</th><th>Q4&nbsp;T</th><th>Total</th>");
            r2 += "</tr>";
            $thead.append(r2);

        } else { // year
            /* row1：固定欄 + 年總 */
            let r1 = "<tr>" + fcTitle();
            years.forEach((y, i) => r1 += `<th class="year-${i % 4}">${y}&nbsp;Total</th>`);
            r1 += "</tr>";
            $thead.append(r1);
        }
    }

    /* ---------- 表體 ---------- */
    function buildBody(list, years, viewType) {
        const $tb = $("#resultTable tbody").empty();
        if (!list.length) {
            $tb.html("<tr><td colspan='999' class='text-center'>-- 無資料 --</td></tr>");
            return;
        }
        const sb = [];
        list.forEach(r => {
            sb.push("<tr>");
            /* 固定欄 */
            sb.push(`<td class="col-customer">${r.Customer || ""}</td>`);
            sb.push(`<td class="col-inch">${r.Inch || ""}</td>`);
            sb.push(`<td class="col-custitem">${r.CustItem || ""}</td>`);
            sb.push(`<td class="col-oaitem">${r.OAItem || ""}</td>`);
            sb.push(`<td class="col-substrate">${r.Substrate || ""}</td>`);
            sb.push(`<td class="col-price">${r.Price || ""}</td>`);
            sb.push(`<td class="col-thk1">${r.Thk1 || ""}</td>`);
            sb.push(`<td class="col-res1">${r.Res1 || ""}</td>`);
            sb.push(`<td class="col-thk2">${r.Thk2 || ""}</td>`);
            sb.push(`<td class="col-res2">${r.Res2 || ""}</td>`);
            sb.push(`<td class="col-thk3">${r.Thk3 || ""}</td>`);
            sb.push(`<td class="col-res3">${r.Res3 || ""}</td>`);

            /* 動態欄 */
            years.forEach(y => {
                const p = y + "_";
                if (viewType === "month") {
                    sb.push(`
          <td>${r[p + "Jan"].toLocaleString()}</td><td>${r[p + "Feb"].toLocaleString()}</td><td>${r[p + "Mar"].toLocaleString()}</td><td>${r[p + "Q1_Total"].toLocaleString()}</td>
          <td>${r[p + "Apr"].toLocaleString()}</td><td>${r[p + "May"].toLocaleString()}</td><td>${r[p + "Jun"].toLocaleString()}</td><td>${r[p + "Q2_Total"].toLocaleString()}</td>
          <td>${r[p + "Jul"].toLocaleString()}</td><td>${r[p + "Aug"].toLocaleString()}</td><td>${r[p + "Sep"].toLocaleString()}</td><td>${r[p + "Q3_Total"].toLocaleString()}</td>
          <td>${r[p + "Oct"].toLocaleString()}</td><td>${r[p + "Nov"].toLocaleString()}</td><td>${r[p + "Dec"].toLocaleString()}</td><td>${r[p + "Q4_Total"].toLocaleString()}</td>
          <td>${r[p + "Total"].toLocaleString()}</td>`);
                } else if (viewType === "quarter") {
                    sb.push(`<td>${r[p + "Q1_Total"].toLocaleString()}</td><td>${r[p + "Q2_Total"].toLocaleString()}</td><td>${r[p + "Q3_Total"].toLocaleString()}</td><td>${r[p + "Q4_Total"].toLocaleString()}</td><td>${r[p + "Total"].toLocaleString()}</td>`);
                } else {
                    sb.push(`<td>${r[p + "Total"].toLocaleString()}</td>`);
                }
            });
            sb.push("</tr>");
        });
        $tb.html(sb.join(""));
    }

    /* ---------- 合計 ---------- */
    function updateTotals() {
        const $rows = $("#resultTable tbody tr");
        if (!$rows.length) return;
        const colCnt = $rows.first().children().length;
        const sums = Array(colCnt).fill(0);
        $rows.each(function () {
            $(this).children().each(function (i) {
                const v = parseFloat($(this).text().replace(/,/g, ""));
                if (!isNaN(v)) sums[i] += v;
            });
        });
        const $tot = $("#resultTable thead tr.totals-row");
        $tot.children().each(function (i) {
            if (i === 0) { $(this).text("總計"); }
            else if (i < 12) { $(this).text(""); }
            else {
                const v = sums[i];
                $(this).text(v ? v.toLocaleString() : "");
            }
        });
    }
</script>
</body>
</html>
