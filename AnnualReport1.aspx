<%@ Page Language="C#" AutoEventWireup="true"
    CodeBehind="AnnualReport1.aspx.cs"
    Inherits="CRM.AnnualReport1" %>

<!DOCTYPE html>
<html lang="zh-TW">
<head runat="server">
    <meta charset="utf-8" />
    <title>年度查詢報表</title>

    <!-- Bootstrap & jQuery -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" />
    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>

    <!-- jQuery-UI：日期選擇器 -->
    <link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css" />
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

    <style>
        /* ---- Sticky 4 層表頭 (totals + 3 原層) ---- */
        thead tr.totals-row  th { position:sticky; top:0;    z-index:4; background:#fff; color:#007bff; }
        thead tr:nth-child(2) th{ position:sticky; top:48px; z-index:3; background:#f2f2f2; }
        thead tr:nth-child(3) th{ position:sticky; top:96px; z-index:2; background:#fafafa; }
        thead tr:nth-child(4) th{ position:sticky; top:144px;z-index:1; background:#fff; }

        /* 去除固定欄 row1/row2 灰塊與框線 */
        .fc-blank{ background:#fff!important; border:none!important; }

        /* row3 固定欄灰底 */
        .fc{ background:#f2f2f2!important; }

        /* 年度色塊 (最多 4 年) */
        .year-0{background:#92D050;} .year-1{background:#00B0F0;}
        .year-2{background:#FFC000;} .year-3{background:#A9A9A9;}

        /* 捲動區 */
        .table-responsive{ max-height:700px; overflow:auto; }

        /* 統一欄寬 (min 100 / max 200) */
        #resultTable th, #resultTable td{
            min-width:100px; max-width:200px;
            white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
        }

        /* 調整表格線與 padding */
        .table thead th{ vertical-align:bottom; border-bottom:0; }
        .table-bordered td, .table-bordered th{ border:0; }
        .table-sm td, .table-sm th{ padding:.75rem; }
    </style>
</head>
<body>
<form id="form1" runat="server" class="container-fluid my-3">

    <!-- ===== 查詢條件 (略，與記憶版本一致) ===== -->
    <!-- 省略：Department、Business、Region、Year... 內容與之前相同 -->

    <!-- ===== 操作按鈕 ===== -->
    <!-- 同上一版，略 -->

    <!-- ===== 報表 ===== -->
    <div class="table-responsive">
        <table id="resultTable" class="table table-bordered table-sm">
            <thead></thead>
            <tbody></tbody>
        </table>
    </div>
</form>

<!-- =================== JS =================== -->
<script>
$(function () {
    $(".datepicker").datepicker({ dateFormat:"yy/mm/dd" });
    const y=new Date().getFullYear().toString();
    $(`input[name=year][value=${y}]`).prop("checked",true);

    $("#btnQuery").on("click",queryReport);
    $(".col-toggle").on("change",function(){
        $("th."+this.value+", td."+this.value).toggle(this.checked);
    });
});

/* ---------------- AJAX 取資料 ---------------- */
function queryReport(){
    const years=$("input[name=year]:checked").map((_,x)=>+x.value).get();
    if(!years.length){ alert("請選年份"); return; }

    $.ajax({
        type:"POST", url:"AnnualReport1.aspx/GetReportData",
        contentType:"application/json; charset=utf-8", dataType:"json",
        data:JSON.stringify({ yearList:years }),
        success:r=>{
            const data=r.d||r;
            buildHeader(years);
            buildBody(data,years);
            updateTotals();   // 產生後即計算合計
            $(".col-toggle").each(function(){ if(!this.checked) $("."+this.value).hide(); });
        },
        error:e=>{ console.error(e); alert("查詢失敗"); }
    });
}

/* --------- 固定欄定義 --------- */
const fixed=[
  {t:"Customer" ,c:"col-customer"},
  {t:"Inch"     ,c:"col-inch"},
  {t:"Cust&nbsp;Item",c:"col-custitem"},
  {t:"OA&nbsp;Item" ,c:"col-oaitem"},
  {t:"Substrate",c:"col-substrate"},
  {t:"Price"    ,c:"col-price"},
  {t:"Thk1"     ,c:"col-thk1"},
  {t:"Res1"     ,c:"col-res1"},
  {t:"Thk2"     ,c:"col-thk2"},
  {t:"Res2"     ,c:"col-res2"},
  {t:"Thk3"     ,c:"col-thk3"},
  {t:"Res3"     ,c:"col-res3"}
];

/* --------- 建立表頭 --------- */
function buildHeader(years){
    const $th=$("#resultTable thead").empty();

    /* ── totals row (空值，稍後填) ── */
    let rt="<tr class='totals-row'>";
    fixed.forEach(()=> rt+="<th></th>");
    years.forEach(()=>{ for(let i=0;i<17;i++) rt+="<th></th>"; });
    rt+="</tr>";
    $th.append(rt);

    /* row1-row3 (沿用既有) */
    let r1="<tr>";
    fixed.forEach(f=> r1+=`<th class="${f.c} fc-blank"></th>`);
    years.forEach((y,i)=> r1+=`<th colspan="17" class="year-${i%4}">${y}</th>`);
    r1+="</tr>";

    let r2="<tr>";
    fixed.forEach(f=> r2+=`<th class="${f.c} fc-blank"></th>`);
    years.forEach(()=> r2+=
      "<th colspan='4'>Q1</th><th colspan='4'>Q2</th>"+
      "<th colspan='4'>Q3</th><th colspan='4'>Q4</th>"+
      "<th rowspan='2'>Total</th>");
    r2+="</tr>";

    let r3="<tr>";
    fixed.forEach(f=> r3+=`<th class="${f.c} fc">${f.t}</th>`);
    years.forEach(()=> r3+=`
        <th>Jan</th><th>Feb</th><th>Mar</th><th>Q1&nbsp;T</th>
        <th>Apr</th><th>May</th><th>Jun</th><th>Q2&nbsp;T</th>
        <th>Jul</th><th>Aug</th><th>Sep</th><th>Q3&nbsp;T</th>
        <th>Oct</th><th>Nov</th><th>Dec</th><th>Q4&nbsp;T</th>` );
    r3+="</tr>";

    $th.append(r1,r2,r3);
}

/* --------- 表體 --------- */
function buildBody(list,years){
    const $tb=$("#resultTable tbody").empty();
    if(!list.length){ $tb.html("<tr><td colspan='999' class='text-center'>-- 無資料 --</td></tr>"); return;}

    const sb=[];
    list.forEach(r=>{
        sb.push("<tr>");
        sb.push(`<td class="col-customer">${r.Customer}</td>`);
        sb.push(`<td class="col-inch">${r.Inch}</td>`);
        sb.push(`<td class="col-custitem">${r.CustItem}</td>`);
        sb.push(`<td class="col-oaitem">${r.OAItem}</td>`);
        sb.push(`<td class="col-substrate">${r.Substrate}</td>`);
        sb.push(`<td class="col-price">${r.Price}</td>`);
        sb.push(`<td class="col-thk1">${r.Thk1}</td><td class="col-res1">${r.Res1}</td>`);
        sb.push(`<td class="col-thk2">${r.Thk2}</td><td class="col-res2">${r.Res2}</td>`);
        sb.push(`<td class="col-thk3">${r.Thk3}</td><td class="col-res3">${r.Res3}</td>`);
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

/* --------- 計算合計列 --------- */
function updateTotals(){
    const $rows=$("#resultTable tbody tr");
    if(!$rows.length) return;

    const colCnt=$rows.first().children().length;
    const sums=Array(colCnt).fill(0);

    $rows.each(function(){
        $(this).children().each(function(idx){
            const v=parseFloat($(this).text().replace(/,/g,""));
            if(!isNaN(v)) sums[idx]+=v;
        });
    });

    const $totRow=$("#resultTable thead tr.totals-row");
    $totRow.children().each(function(idx){
        if(idx===0){ $(this).text("總計"); }
        else{
            const val=sums[idx];
            $(this).text(val? val.toLocaleString(): "");
        }
    });
}
</script>
</body>
</html>
