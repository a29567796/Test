<%@ Page Language="C#" AutoEventWireup="true"
    CodeBehind="AnnualReport.aspx.cs"
    Inherits="WebApp.AnnualReport" %>

<!DOCTYPE html>
<html lang="zh-TW">
<head runat="server">
    <meta charset="utf-8" />
    <title>年度查詢報表</title>

    <!-- Bootstrap + jQuery -->
    <link rel="stylesheet"
          href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" />
    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>

    <!-- jQuery-UI：日期選擇器 -->
    <link rel="stylesheet"
          href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css" />
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

    <style>
        /* === Sticky 三層表頭 === */
        thead tr:nth-child(1) th { position:sticky; top:0;   z-index:3; background:#f2f2f2; }
        thead tr:nth-child(2) th { position:sticky; top:30px; z-index:2; background:#fafafa; }
        thead tr:nth-child(3) th { position:sticky; top:60px; z-index:1; background:#ffffff; }

        /* 年度色塊 (循環四色) */
        .year-0 { background:#92D050; }
        .year-1 { background:#00B0F0; }
        .year-2 { background:#FFC000; }
        .year-3 { background:#A9A9A9; }

        /* 透明空白表頭 (遮灰) */
        .blank { background:#ffffff!important; border:none!important; }

        /* 固定欄寬 */
        .w-120 { width:120px; }
        .w-90  { width:90px;  }
        .w-70  { width:70px;  }
        .w-60  { width:60px;  }

        /* 捲動區域 */
        .table-responsive { max-height:680px; overflow:auto; }

        /* 表格文字不換行 */
        th,td { white-space:nowrap; }
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
                            <option value="CN">China&nbsp;+&nbsp;HK</option>
                            <option value="JP">Japan</option>
                        </select>
                    </div>
                    <div class="form-group col-md-3">
                        <label>Year (多選)</label><br />
                        <%-- 核取方塊 4 年 --%>
                        <label class="mr-2"><input type="checkbox" name="year" value="2022">2022</label>
                        <label class="mr-2"><input type="checkbox" name="year" value="2023">2023</label>
                        <label class="mr-2"><input type="checkbox" name="year" value="2024">2024</label>
                        <label class="mr-2"><input type="checkbox" name="year" value="2025">2025</label>
                    </div>
                    <div class="form-group col-md-3">
                        <label>Customer</label>
                        <input id="ddlCustomer" class="form-control" placeholder="VI / EPS / ..." />
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

                <!-- 操作 -->
                <button id="btnQuery" type="button" class="btn btn-primary mr-2">查詢</button>
                <asp:Button ID="btnExport" runat="server"
                    CssClass="btn btn-success" Text="匯出 Excel"
                    OnClick="btnExport_Click" />

                <!-- Thk / Res 切換 -->
                <div class="mt-3">
                    <label class="mr-2">顯示欄位：</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk1" checked>Thk1</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res1" checked>Res1</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk2" checked>Thk2</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res2" checked>Res2</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk3" checked>Thk3</label>
                    <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res3" checked>Res3</label>
                </div>
            </div>
        </div>

        <!-- ===== 報表表格 ===== -->
        <div class="table-responsive">
            <table id="resultTable" class="table table-bordered table-sm">
                <thead></thead>
                <tbody></tbody>
            </table>
        </div>
    </form>

    <!-- ========= JavaScript ========= -->
    <script>
        /* === 初始化 === */
        $(function () {
            $(".datepicker").datepicker({ dateFormat: "yy/mm/dd" });
            const y = new Date().getFullYear().toString();
            $(`input[name=year][value=${y}]`).prop("checked", true);

            $("#btnQuery").on("click", queryReport);

            /* Thk / Res 切換 */
            $(".col-toggle").on("change", function () {
                const c = $(this).val();
                const $cells = $("th." + c + ", td." + c);
                $(this).is(":checked") ? $cells.show() : $cells.hide();
            });
        });

        /* === 查詢 === */
        function queryReport() {
            const years = $("input[name=year]:checked").map((i, el) => parseInt(el.value)).get();
            if (years.length === 0) { alert("請至少選擇一年"); return; }

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
                    const list = r.d || r;
                    buildHeader(years);
                    buildBody(list, years);
                    /* 同步核取方塊狀態 */
                    $(".col-toggle").each(function () {
                        if (!$(this).is(":checked"))
                            $("th." + this.value + ", td." + this.value).hide();
                    });
                },
                error: e => { alert("查詢失敗！"); console.error(e); }
            });
        }

        /* === 動態表頭 === */
        const fixedHeaders = [
            { text: "Customer", cls: "w-120" },
            { text: "Inch",     cls: "w-70"  },
            { text: "Cust Item",cls: "w-120" },
            { text: "OA Item",  cls: "w-120" },
            { text: "Substrate",cls: "w-90"  },
            { text: "Price",    cls: "w-90"  },
            { text: "Thk1",     cls: "w-70 col-thk1" },
            { text: "Res1",     cls: "w-70 col-res1" },
            { text: "Thk2",     cls: "w-70 col-thk2" },
            { text: "Res2",     cls: "w-70 col-res2" },
            { text: "Thk3",     cls: "w-70 col-thk3" },
            { text: "Res3",     cls: "w-70 col-res3" }
        ];

        function buildHeader(years) {
            const $thead = $("#resultTable thead").empty();

            /* ---------- 第一列 ---------- */
            let r1 = "<tr>";
            /* 空白 12 欄 (透明) */
            fixedHeaders.forEach(() => r1 += `<th class="blank"></th>`);
            years.forEach((y, idx) => r1 += `<th colspan="17" class="year-${idx % 4}">${y}</th>`);
            r1 += "</tr>";

            /* ---------- 第二列 ---------- */
            let r2 = "<tr>";
            fixedHeaders.forEach(() => r2 += `<th class="blank"></th>`);
            years.forEach(() => {
                for (let q = 1; q <= 4; q++) r2 += `<th colspan="4">Q${q}</th>`;
                r2 += "<th rowspan='2'>Total</th>";
            });
            r2 += "</tr>";

            /* ---------- 第三列 ---------- */
            let r3 = "<tr>";
            /* 固定 12 欄 */
            fixedHeaders.forEach(h => r3 += `<th class="${h.cls}">${h.text}</th>`);
            /* 月份 + 季 Total */
            years.forEach(() => {
                r3 += `
                   <th>Jan</th><th>Feb</th><th>Mar</th><th>Q1&nbsp;T</th>
                   <th>Apr</th><th>May</th><th>Jun</th><th>Q2&nbsp;T</th>
                   <th>Jul</th><th>Aug</th><th>Sep</th><th>Q3&nbsp;T</th>
                   <th>Oct</th><th>Nov</th><th>Dec</th><th>Q4&nbsp;T</th>`;
            });
            r3 += "</tr>";

            $thead.append(r1, r2, r3);
        }

        /* === 表體 === */
        function buildBody(list, years) {
            const $tbody = $("#resultTable tbody").empty();
            if (!list || list.length === 0) {
                $tbody.append("<tr><td class='text-center' colspan='100%'>-- 無資料 --</td></tr>");
                return;
            }
            const sb = [];
            list.forEach(r => {
                sb.push("<tr>");
                /* 固定 12 欄 */
                sb.push(`<td>${r.Customer}</td>`);
                sb.push(`<td>${r.Inch}</td>`);
                sb.push(`<td>${r["Cust Item"]}</td>`);
                sb.push(`<td>${r["OA Item"]}</td>`);
                sb.push(`<td>${r.Substrate}</td>`);
                sb.push(`<td>${r.Price}</td>`);
                sb.push(`<td class="col-thk1">${r.Thk1}</td><td class="col-res1">${r.Res1}</td>`);
                sb.push(`<td class="col-thk2">${r.Thk2}</td><td class="col-res2">${r.Res2}</td>`);
                sb.push(`<td class="col-thk3">${r.Thk3}</td><td class="col-res3">${r.Res3}</td>`);

                /* 動態年份欄 */
                years.forEach(y => {
                    const p = y + "_";
                    sb.push(`
                        <td>${r[p+"Jan"]}</td><td>${r[p+"Feb"]}</td><td>${r[p+"Mar"]}</td><td>${r[p+"Q1_Total"]}</td>
                        <td>${r[p+"Apr"]}</td><td>${r[p+"May"]}</td><td>${r[p+"Jun"]}</td><td>${r[p+"Q2_Total"]}</td>
                        <td>${r[p+"Jul"]}</td><td>${r[p+"Aug"]}</td><td>${r[p+"Sep"]}</td><td>${r[p+"Q3_Total"]}</td>
                        <td>${r[p+"Oct"]}</td><td>${r[p+"Nov"]}</td><td>${r[p+"Dec"]}</td><td>${r[p+"Q4_Total"]}</td>
                        <td>${r[p+"Total"]}</td>`);
                });
                sb.push("</tr>");
            });
            $tbody.html(sb.join(""));
        }
    </script>
</body>
</html>
