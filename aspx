<%@ Page Language="C#" AutoEventWireup="true"
    CodeBehind="AnnualReport.aspx.cs"
    Inherits="WebApp.AnnualReport" %>

<!DOCTYPE html>
<html lang="zh-TW">
<head runat="server">
    <meta charset="utf-8" />
    <title>年度查詢報表</title>

    <!-- Bootstrap & jQuery -->
    <link rel="stylesheet"
          href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" />
    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>

    <!-- jQuery UI（日期選擇器） -->
    <link rel="stylesheet"
          href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css" />
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

    <style>
        /* ---- 固定三層表頭 ---- */
        thead tr:nth-child(1) th { position: sticky; top: 0;   z-index: 3; background:#f2f2f2; }
        thead tr:nth-child(2) th { position: sticky; top: 40px; z-index: 2; background:#fafafa; }
        thead tr:nth-child(3) th { position: sticky; top: 80px; z-index: 1; background:#fff; }

        /* 年度色塊（可自行調整） */
        .year-0 { background:#92D050; }
        .year-1 { background:#00B0F0; }
        .year-2 { background:#FFC000; }
        .year-3 { background:#A9A9A9; }

        /* 資料過多時保持表格可捲動 */
        .table-responsive { max-height: 700px; overflow:auto; }
    </style>
</head>
<body>
    <form id="form1" runat="server" class="container-fluid my-3">

        <!-- ===== 查詢條件區塊 ===== -->
        <div class="card mb-3">
            <div class="card-header">查詢條件</div>
            <div class="card-body">

                <!-- 第一列 -->
                <div class="form-row">
                    <div class="form-group col-md-2">
                        <label>Department</label>
                        <select id="ddlDept" class="form-control">
                            <option value="">All</option>
                            <option>S100</option><option>S200</option>
                        </select>
                    </div>
                    <div class="form-group col-md-2">
                        <label>Business</label>
                        <select id="ddlBiz" class="form-control">
                            <option value="">All</option>
                            <option>Alice</option><option>Bob</option>
                        </select>
                    </div>
                    <div class="form-group col-md-2">
                        <label>Region</label>
                        <select id="ddlRegion" class="form-control">
                            <option value="">All</option>
                            <option value="TW">Taiwan</option>
                            <option value="CN">China + HK</option>
                            <option value="JP">Japan</option>
                        </select>
                    </div>
                    <div class="form-group col-md-3">
                        <label>Year (多選)</label>
                        <div>
                            <%-- 近四年 --%>
                            <label class="mr-2"><input type="checkbox" name="year" value="2022" />2022</label>
                            <label class="mr-2"><input type="checkbox" name="year" value="2023" />2023</label>
                            <label class="mr-2"><input type="checkbox" name="year" value="2024" />2024</label>
                            <label class="mr-2"><input type="checkbox" name="year" value="2025" />2025</label>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <label>Customer</label>
                        <select id="ddlCustomer" class="form-control">
                            <option value="">All</option>
                            <option>VI</option><option>EPS</option>
                        </select>
                    </div>
                </div>

                <!-- 第二列 -->
                <div class="form-row">
                    <div class="form-group col-md-2">
                        <label>Cust&nbsp;Item</label>
                        <input id="txtCustItem" class="form-control" />
                    </div>
                    <div class="form-group col-md-2">
                        <label>Substrate</label>
                        <select id="ddlSubstrate" class="form-control">
                            <option value="">All</option>
                            <option>SubA</option><option>SubB</option>
                        </select>
                    </div>
                    <div class="form-group col-md-2">
                        <label>Products&nbsp;type</label>
                        <select id="ddlProductType" class="form-control">
                            <option value="">All</option>
                            <option>GaN</option><option>SiC</option>
                        </select>
                    </div>
                    <div class="form-group col-md-2">
                        <label>Ship&nbsp;to</label>
                        <input id="txtShipTo" class="form-control" placeholder="e.g. Taiwan" />
                    </div>
                    <div class="form-group col-md-2">
                        <label>Real&nbsp;Date</label>
                        <input id="txtRealDate" class="form-control datepicker" placeholder="yyyy/MM/dd" />
                    </div>
                    <div class="form-group col-md-1">
                        <label>Status</label>
                        <select id="ddlStatus" class="form-control">
                            <option value="">All</option>
                            <option>Open</option><option>Closed</option>
                        </select>
                    </div>
                    <div class="form-group col-md-1">
                        <label>Order&nbsp;Type</label>
                        <select id="ddlOrderType" class="form-control">
                            <option value="">All</option>
                            <option>3x01</option><option>3x02</option>
                        </select>
                    </div>
                </div>

                <!-- 操作按鈕 -->
                <button id="btnQuery" type="button" class="btn btn-primary mr-2">查詢</button>
                <asp:Button ID="btnExport" runat="server"
                    CssClass="btn btn-success"
                    Text="匯出 Excel" OnClick="btnExport_Click" />

                <!-- Thk/Res 顯示切換 -->
                <div class="mt-3">
                    <label class="mr-2">顯示欄位：</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk1" checked />Thk1</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk2" checked />Thk2</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk3" checked />Thk3</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res1" checked />Res1</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res2" checked />Res2</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res3" checked />Res3</label>
                </div>

            </div>
        </div>

        <!-- ===== 報表結果表格 ===== -->
        <div class="table-responsive">
            <table id="resultTable" class="table table-bordered table-sm">
                <thead></thead>
                <tbody></tbody>
            </table>
        </div>
    </form>

    <!-- -------- 前端腳本 -------- -->
    <script>
        /* ---------- UI 初始化 ---------- */
        $(function () {
            /* 日期選擇器 */
            $(".datepicker").datepicker({ dateFormat: "yy/mm/dd" });

            /* 預設勾選當前年份 */
            const y = new Date().getFullYear().toString();
            $(`input[name=year][value=${y}]`).prop("checked", true);

            /* 查詢按鈕 */
            $("#btnQuery").on("click", queryReport);

            /* Thk/Res 欄位顯示切換 */
            $(".col-toggle").on("change", function () {
                const cls = $(this).val();
                const $cells = $("th." + cls + ", td." + cls);
                $(this).is(":checked") ? $cells.show() : $cells.hide();
            });
        });

        /* ---------- 呼叫 AJAX 取得資料 ---------- */
        function queryReport() {
            // 收集多選年份
            const years = $("input[name=year]:checked").map((_, el) => parseInt(el.value)).get();
            if (years.length === 0) { alert("請至少選擇一個年份"); return; }

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
                success: res => {
                    const data = res.d || res;
                    buildHeader(years);
                    buildBody(data, years);
                    // 初始同步核取方塊狀態
                    $(".col-toggle").each(function () {
                        if (!$(this).is(":checked")) {
                            $("th." + this.value + ", td." + this.value).hide();
                        }
                    });
                },
                error: err => {
                    console.error(err);
                    alert("查詢失敗，請檢查參數或稍後再試！");
                }
            });
        }

        /* ---------- 動態表頭 ---------- */
        function buildHeader(years) {
            const $thead = $("#resultTable thead").empty();
            // 第一列：年度區塊 + Thk/Res
            let row1 = "<tr>";
            ["thk1", "thk2", "thk3", "res1", "res2", "res3"].forEach(k =>
                row1 += `<th rowspan="3" class="col-${k.toLowerCase()}">${k.toUpperCase()}</th>`);
            years.forEach((y, idx) =>
                row1 += `<th colspan="17" class="year-${idx % 4}">${y}</th>`);
            row1 += "</tr>";
            // 第二列：Q1~Q4 + 年Total (rowspan=2)
            let row2 = "<tr>";
            years.forEach(() => {
                for (let i = 0; i < 4; i++) row2 += "<th colspan='4'>Q" + (i + 1) + "</th>";
                row2 += "<th rowspan='2'>Total</th>";
            });
            row2 += "</tr>";
            // 第三列：月份 + 季合計
            let row3 = "<tr>";
            years.forEach(() => {
                row3 += `
                  <th>Jan</th><th>Feb</th><th>Mar</th><th>Q1&nbsp;T</th>
                  <th>Apr</th><th>May</th><th>Jun</th><th>Q2&nbsp;T</th>
                  <th>Jul</th><th>Aug</th><th>Sep</th><th>Q3&nbsp;T</th>
                  <th>Oct</th><th>Nov</th><th>Dec</th><th>Q4&nbsp;T</th>`;
            });
            row3 += "</tr>";
            $thead.append(row1, row2, row3);
        }

        /* ---------- 產生表體 ---------- */
        function buildBody(list, years) {
            const $tbody = $("#resultTable tbody").empty();
            if (!list || list.length === 0) {
                $tbody.append("<tr><td colspan='100%' class='text-center'>-- 無資料 --</td></tr>");
                return;
            }
            const sb = [];
            list.forEach(r => {
                sb.push("<tr>");
                sb.push(`<td class="col-thk1">${r.Thk1 || ""}</td>`);
                sb.push(`<td class="col-thk2">${r.Thk2 || ""}</td>`);
                sb.push(`<td class="col-thk3">${r.Thk3 || ""}</td>`);
                sb.push(`<td class="col-res1">${r.Res1 || ""}</td>`);
                sb.push(`<td class="col-res2">${r.Res2 || ""}</td>`);
                sb.push(`<td class="col-res3">${r.Res3 || ""}</td>`);
                years.forEach(y => {
                    const p = y + "_"; // 前綴
                    sb.push(`
                      <td>${r[p+"Jan"]||""}</td><td>${r[p+"Feb"]||""}</td><td>${r[p+"Mar"]||""}</td><td>${r[p+"Q1_Total"]||""}</td>
                      <td>${r[p+"Apr"]||""}</td><td>${r[p+"May"]||""}</td><td>${r[p+"Jun"]||""}</td><td>${r[p+"Q2_Total"]||""}</td>
                      <td>${r[p+"Jul"]||""}</td><td>${r[p+"Aug"]||""}</td><td>${r[p+"Sep"]||""}</td><td>${r[p+"Q3_Total"]||""}</td>
                      <td>${r[p+"Oct"]||""}</td><td>${r[p+"Nov"]||""}</td><td>${r[p+"Dec"]||""}</td><td>${r[p+"Q4_Total"]||""}</td>
                      <td>${r[p+"Total"]||""}</td>`);
                });
                sb.push("</tr>");
            });
            $tbody.html(sb.join(""));
        }
    </script>
</body>
</html>