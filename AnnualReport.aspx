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

    <!-- jQuery UI：日期選擇器 -->
    <link rel="stylesheet"
          href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css" />
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

    <style>
        /* ---------- Sticky 三層表頭 ---------- */
        thead tr:nth-child(1) th { position: sticky; top: 0;    z-index: 3; background:#f2f2f2; }
        thead tr:nth-child(2) th { position: sticky; top: 40px; z-index: 2; background:#fafafa; }
        thead tr:nth-child(3) th { position: sticky; top: 80px; z-index: 1; background:#fff; }

        /* 固定欄位在 row1/row2 的空白 th：改為白底去大灰塊 */
        .fc-blank            { background:#fff!important; }

        /* row3 固定欄位標題灰底 */
        .fc                  { background:#f2f2f2!important; }

        /* 年度色塊(最多 4 年) */
        .year-0 { background:#92D050; }
        .year-1 { background:#00B0F0; }
        .year-2 { background:#FFC000; }
        .year-3 { background:#A9A9A9; }

        /* 表格捲動區 */
        .table-responsive { max-height:700px; overflow:auto; }

        /* 固定欄寬 */
        .w120{width:120px;} .w100{width:100px;} .w80{width:80px;} .w60{width:60px;}
    </style>
</head>
<body>
    <form id="form1" runat="server" class="container-fluid my-3">

        <!-- ===== 查詢條件（保留原架構，略） ===== -->
        <!-- --- 內容與之前版本相同，略 --- -->

        <!-- ===== 操作按鈕與 Thk/Res 切換 ===== -->
        <!-- 同先前版本，略 -->

        <!-- ===== 報表結果表格 ===== -->
        <div class="table-responsive">
            <table id="resultTable" class="table table-bordered table-sm">
                <thead></thead>
                <tbody></tbody>
            </table>
        </div>
    </form>

<!-- ========================= 前端腳本 ========================= -->
<script>
$(function () {
    /* 日期選擇器 */
    $(".datepicker").datepicker({ dateFormat: "yy/mm/dd" });

    /* 預設勾選當前年份 */
    const y = new Date().getFullYear().toString();
    $(`input[name=year][value=${y}]`).prop("checked", true);

    /* 查詢 */
    $("#btnQuery").on("click", queryReport);

    /* Thk/Res 欄切換 */
    $(".col-toggle").on("change", function () {
        const cls = $(this).val();
        $("th."+cls+", td."+cls).toggle( $(this).is(":checked") );
    });
});

/* ---------- AJAX 取得資料 ---------- */
function queryReport(){
    const years = $("input[name=year]:checked").map((_,el)=>+el.value).get();
    if(!years.length){ alert("請至少選一年"); return; }

    $.ajax({
        type:"POST",
        url:"AnnualReport.aspx/GetReportData",
        contentType:"application/json; charset=utf-8",
        dataType:"json",
        data: JSON.stringify({
            /* 依 UI 收集參數 —— 省略，其餘同前版 */
            yearList: years
        }),
        success: r=>{
            const data = r.d || r;
            buildHeader(years);
            buildBody(data, years);
            // 同步核取方塊顯示
            $(".col-toggle").each(function(){
                if(!$(this).is(":checked")){
                    $("th."+this.value+", td."+this.value).hide();
                }
            });
        },
        error:e=>{
            console.error(e); alert("查詢失敗");
        }
    });
}

/* ---------- 動態表頭 ---------- */
const fixedCols = [
    { t:"Customer" , c:"col-customer" , w:"w120" },
    { t:"Inch"     , c:"col-inch"     , w:"w60"  },
    { t:"Cust&nbsp;Item", c:"col-custitem", w:"w120"},
    { t:"OA&nbsp;Item"  , c:"col-oaitem"  , w:"w120"},
    { t:"Substrate", c:"col-substrate", w:"w100" },
    { t:"Price"    , c:"col-price"    , w:"w80"  },
    { t:"Thk1"     , c:"col-thk1"     , w:"w80"  },
    { t:"Res1"     , c:"col-res1"     , w:"w80"  },
    { t:"Thk2"     , c:"col-thk2"     , w:"w80"  },
    { t:"Res2"     , c:"col-res2"     , w:"w80"  },
    { t:"Thk3"     , c:"col-thk3"     , w:"w80"  },
    { t:"Res3"     , c:"col-res3"     , w:"w80"  }
];

function buildHeader(years){
    const $thead=$("#resultTable thead").empty();

    /* ---- row1 ---- */
    let r1="<tr>";
    fixedCols.forEach(fc=> r1 += `<th class="${fc.c} fc-blank ${fc.w}"></th>`);
    years.forEach((y,i)=> r1 += `<th colspan="17" class="year-${i%4}">${y}</th>` );
    r1+="</tr>";

    /* ---- row2 ---- */
    let r2="<tr>";
    fixedCols.forEach(fc=> r2 += `<th class="${fc.c} fc-blank ${fc.w}"></th>`);
    years.forEach(()=> r2 += "<th colspan='4'>Q1</th><th colspan='4'>Q2</th><th colspan='4'>Q3</th><th colspan='4'>Q4</th><th rowspan='2'>Total</th>");
    r2+="</tr>";

    /* ---- row3 ---- */
    let r3="<tr>";
    fixedCols.forEach(fc=> r3 += `<th class="${fc.c} fc ${fc.w}">${fc.t}</th>`);

    years.forEach(()=>{
        r3+=`
          <th>Jan</th><th>Feb</th><th>Mar</th><th>Q1&nbsp;T</th>
          <th>Apr</th><th>May</th><th>Jun</th><th>Q2&nbsp;T</th>
          <th>Jul</th><th>Aug</th><th>Sep</th><th>Q3&nbsp;T</th>
          <th>Oct</th><th>Nov</th><th>Dec</th><th>Q4&nbsp;T</th>`;
    });
    r3+="</tr>";

    $thead.append(r1,r2,r3);
}

/* ---------- 產生表體 ---------- */
function buildBody(list, years){
    const $tb=$("#resultTable tbody").empty();
    if(!list.length){ $tb.html("<tr><td colspan='999' class='text-center'>-- 無資料 --</td></tr>"); return; }

    const sb=[];
    list.forEach(r=>{
        sb.push("<tr>");
        sb.push(`<td class="col-customer">${r.Customer||""}</td>`);
        sb.push(`<td class="col-inch">${r.Inch||""}</td>`);
        sb.push(`<td class="col-custitem">${r.CustItem||""}</td>`);
        sb.push(`<td class="col-oaitem">${r.OAItem||""}</td>`);
        sb.push(`<td class="col-substrate">${r.Substrate||""}</td>`);
        sb.push(`<td class="col-price">${r.Price||""}</td>`);
        sb.push(`<td class="col-thk1">${r.Thk1||""}</td>`);
        sb.push(`<td class="col-res1">${r.Res1||""}</td>`);
        sb.push(`<td class="col-thk2">${r.Thk2||""}</td>`);
        sb.push(`<td class="col-res2">${r.Res2||""}</td>`);
        sb.push(`<td class="col-thk3">${r.Thk3||""}</td>`);
        sb.push(`<td class="col-res3">${r.Res3||""}</td>`);

        years.forEach(y=>{
            const p=y+"_";
            sb.push(`
              <td>${r[p+"Jan"]}</td><td>${r[p+"Feb"]}</td><td>${r[p+"Mar"]}</td><td>${r[p+"Q1_Total"]}</td>
              <td>${r[p+"Apr"]}</td><td>${r[p+"May"]}</td><td>${r[p+"Jun"]}</td><td>${r[p+"Q2_Total"]}</td>
              <td>${r[p+"Jul"]}</td><td>${r[p+"Aug"]}</td><td>${r[p+"Sep"]}</td><td>${r[p+"Q3_Total"]}</td>
              <td>${r[p+"Oct"]}</td><td>${r[p+"Nov"]}</td><td>${r[p+"Dec"]}</td><td>${r[p+"Q4_Total"]}</td>
              <td>${r[p+"Total"]}</td>`);
        });
        sb.push("</tr>");
    });
    $tb.html(sb.join(""));
}
</script>
</body>
</html>
