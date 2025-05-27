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
    <!-- jQuery-UI（日期） -->
    <link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css" />
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

    <style>
        /* ── Sticky 4 層 (totals + 動態 1~3) ── */
        thead tr.totals-row th{position:sticky;top:0;   z-index:5;background:#fff;color:#007bff;}
        thead tr.lv1        th{position:sticky;top:48px;z-index:4;background:#f2f2f2;}
        thead tr.lv2        th{position:sticky;top:96px;z-index:3;background:#fafafa;}
        thead tr.lv3        th{position:sticky;top:144px;z-index:2;background:#fff;}

        /* 固定欄空白 */
        .fc-blank{background:#fff!important;border:none!important;}
        .fc      {background:#f2f2f2!important;}

        /* 年度色 */
        .year-0{background:#92D050;}
        .year-1{background:#00B0F0;}
        .year-2{background:#FFC000;}
        .year-3{background:#A9A9A9;}

        .table-responsive{max-height:700px;overflow:auto;}

        /* 欄寬 100~200 */
        #resultTable th,#resultTable td{
            min-width:100px;max-width:200px;
            white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
        }
        .table thead th{vertical-align:bottom;border-bottom:0;}
        .table-bordered td,.table-bordered th{border:0;}
        .table-sm td,.table-sm th{padding:.75rem;}
    </style>
</head>
<body>
<form id="form1" runat="server" class="container-fluid my-3">

    <!-- === 查詢條件 (簡化示範) === -->
    <div class="mb-2">
        <label class="mr-2 font-weight-bold">View：</label>
        <label class="mr-2"><input type="radio" name="view" value="month"  checked />月</label>
        <label class="mr-2"><input type="radio" name="view" value="quarter"      />季</label>
        <label class="mr-2"><input type="radio" name="view" value="year"         />年</label>

        <label class="ml-4 mr-2 font-weight-bold">Year：</label>
        <label class="mr-1"><input type="checkbox" name="year" value="2023">2023</label>
        <label class="mr-1"><input type="checkbox" name="year" value="2024">2024</label>
        <label class="mr-1"><input type="checkbox" name="year" value="2025">2025</label>

        <button id="btnQuery" type="button" class="btn btn-primary btn-sm ml-3">查詢</button>
        <asp:Button ID="btnExport" runat="server" Text="匯出 Excel" CssClass="btn btn-success btn-sm ml-2" OnClick="btnExport_Click" />
    </div>

    <!-- Thk/Res 切換 (保留) -->
    <div class="mb-3">
        <label class="mr-2">顯示欄位：</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk1" checked />Thk1</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk2" checked />Thk2</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk3" checked />Thk3</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res1" checked />Res1</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res2" checked />Res2</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res3" checked />Res3</label>
    </div>

    <!-- === 表格 === -->
    <div class="table-responsive">
        <table id="resultTable" class="table table-bordered table-sm">
            <thead></thead>
            <tbody></tbody>
        </table>
    </div>
</form>

<!-- ========= JS ========= -->
<script>
$(function () {
    $(".datepicker").datepicker({ dateFormat:"yy/mm/dd" });
    const y=new Date().getFullYear().toString();
    $(`input[name=year][value=${y}]`).prop("checked",true);

    $("#btnQuery").on("click",queryReport);
    $(".col-toggle").on("change",function(){ $("."+this.value).toggle(this.checked); });
});

/* 固定欄 (12) */
const fixed=[
  {t:"Customer",c:"col-customer"},
  {t:"Inch",c:"col-inch"},
  {t:"Cust&nbsp;Item",c:"col-custitem"},
  {t:"OA&nbsp;Item",c:"col-oaitem"},
  {t:"Substrate",c:"col-substrate"},
  {t:"Price",c:"col-price"},
  {t:"Thk1",c:"col-thk1"},
  {t:"Res1",c:"col-res1"},
  {t:"Thk2",c:"col-thk2"},
  {t:"Res2",c:"col-res2"},
  {t:"Thk3",c:"col-thk3"},
  {t:"Res3",c:"col-res3"}
];

/* -------------- AJAX 取得 -------------- */
function queryReport(){
    const years=$("input[name=year]:checked").map((_,x)=>+x.value).get();
    const view=$("input[name=view]:checked").val();   // month / quarter / year
    if(!years.length){ alert("請選年份"); return; }

    $.ajax({
        type:"POST", url:"AnnualReport1.aspx/GetReportData",
        contentType:"application/json; charset=utf-8", dataType:"json",
        data:JSON.stringify({ yearList:years }),
        success:r=>{
            const data=r.d||r;
            buildHeader(years,view);
            buildBody(data,years,view);
            updateTotals();
            $(".col-toggle").each(function(){ if(!this.checked) $("."+this.value).hide(); });
        },
        error:e=>{ console.error(e); alert("查詢失敗"); }
    });
}

/* -------------- 動態表頭 -------------- */
function buildHeader(years,view){
    const $thead=$("#resultTable thead").empty();

    /* totals-row (空，稍後填值) */
    let trT="<tr class='totals-row'>";
    fixed.forEach(()=> trT+="<th></th>");
    const colPerYear = view==="month"?17 : view==="quarter"?5 : 1;
    years.forEach(()=>{ for(let i=0;i<colPerYear;i++) trT+="<th></th>"; });
    trT+="</tr>";
    $thead.append(trT);

    if(view==="month"){           // === 月：3 層表頭 ===
        // lv1 年
        let r1="<tr class='lv1'>";
        fixed.forEach(f=> r1+=`<th class='${f.c} fc-blank'></th>`);
        years.forEach((y,i)=> r1+=`<th colspan='17' class='year-${i%4}'>${y}</th>`);
        r1+="</tr>";

        // lv2 季
        let r2="<tr class='lv2'>";
        fixed.forEach(f=> r2+=`<th class='${f.c} fc-blank'></th>`);
        years.forEach(()=> r2+="<th colspan='4'>Q1</th><th colspan='4'>Q2</th><th colspan='4'>Q3</th><th colspan='4'>Q4</th><th rowspan='2'>Total</th>");
        r2+="</tr>";

        // lv3 月
        let r3="<tr class='lv3'>";
        fixed.forEach(f=> r3+=`<th class='${f.c} fc'>${f.t}</th>`);
        years.forEach(()=> r3+=`
            <th>Jan</th><th>Feb</th><th>Mar</th><th>Q1&nbsp;T</th>
            <th>Apr</th><th>May</th><th>Jun</th><th>Q2&nbsp;T</th>
            <th>Jul</th><th>Aug</th><th>Sep</th><th>Q3&nbsp;T</th>
            <th>Oct</th><th>Nov</th><th>Dec</th><th>Q4&nbsp;T</th>` );
        r3+="</tr>";

        $thead.append(r1,r2,r3);
    }
    else if(view==="quarter"){    // === 季：2 層 ===
        // lv1 年
        let r1="<tr class='lv1'>";
        fixed.forEach(f=> r1+=`<th class='${f.c} fc-blank'></th>`);
        years.forEach((y,i)=> r1+=`<th colspan='5' class='year-${i%4}'>${y}</th>`);
        r1+="</tr>";

        // lv2 Q1T~Q4T+Total
        let r2="<tr class='lv2'>";
        fixed.forEach(f=> r2+=`<th class='${f.c} fc'>${f.t}</th>`);
        years.forEach(()=> r2+="<th>Q1&nbsp;T</th><th>Q2&nbsp;T</th><th>Q3&nbsp;T</th><th>Q4&nbsp;T</th><th>Total</th>");
        r2+="</tr>";

        $thead.append(r1,r2);
    }
    else{                         // === 年：1 層 ===
        let r1="<tr class='lv1'>";
        fixed.forEach(f=> r1+=`<th class='${f.c} fc'>${f.t}</th>`);
        years.forEach((y,i)=> r1+=`<th class='year-${i%4}'>${y}&nbsp;Total</th>`);
        r1+="</tr>";
        $thead.append(r1);
    }
}

/* -------------- 表體 -------------- */
function buildBody(list,years,view){
    const $tb=$("#resultTable tbody").empty();
    if(!list.length){ $tb.html("<tr><td colspan='999' class='text-center'>-- 無資料 --</td></tr>"); return;}

    const html=[];
    list.forEach(r=>{
        html.push("<tr>");
        html.push(`<td class="col-customer">${r.Customer}</td>`);
        html.push(`<td class="col-inch">${r.Inch}</td>`);
        html.push(`<td class="col-custitem">${r.CustItem}</td>`);
        html.push(`<td class="col-oaitem">${r.OAItem}</td>`);
        html.push(`<td class="col-substrate">${r.Substrate}</td>`);
        html.push(`<td class="col-price">${r.Price}</td>`);
        html.push(`<td class="col-thk1">${r.Thk1}</td><td class="col-res1">${r.Res1}</td>`);
        html.push(`<td class="col-thk2">${r.Thk2}</td><td class="col-res2">${r.Res2}</td>`);
        html.push(`<td class="col-thk3">${r.Thk3}</td><td class="col-res3">${r.Res3}</td>`);

        years.forEach(y=>{
            const p=y+"_";
            if(view==="month"){
                html.push(`
                  <td>${r[p+"Jan"]}</td><td>${r[p+"Feb"]}</td><td>${r[p+"Mar"]}</td><td>${r[p+"Q1_Total"]}</td>
                  <td>${r[p+"Apr"]}</td><td>${r[p+"May"]}</td><td>${r[p+"Jun"]}</td><td>${r[p+"Q2_Total"]}</td>
                  <td>${r[p+"Jul"]}</td><td>${r[p+"Aug"]}</td><td>${r[p+"Sep"]}</td><td>${r[p+"Q3_Total"]}</td>
                  <td>${r[p+"Oct"]}</td><td>${r[p+"Nov"]}</td><td>${r[p+"Dec"]}</td><td>${r[p+"Q4_Total"]}</td>
                  <td>${r[p+"Total"]}</td>`);
            }
            else if(view==="quarter"){
                html.push(`
                  <td>${r[p+"Q1_Total"]}</td><td>${r[p+"Q2_Total"]}</td><td>${r[p+"Q3_Total"]}</td><td>${r[p+"Q4_Total"]}</td>
                  <td>${r[p+"Total"]}</td>`);
            }
            else{ // year
                html.push(`<td>${r[p+"Total"]}</td>`);
            }
        });
        html.push("</tr>");
    });
    $tb.html(html.join(""));
}

/* -------------- 合計列 -------------- */
function updateTotals(){
    const $rows=$("#resultTable tbody tr");
    if(!$rows.length)return;

    const sums=[];
    $rows.each(function(){
        $(this).children().each(function(i){
            const v=parseFloat($(this).text().replace(/,/g,""));
            if(!isNaN(v)) sums[i]=(sums[i]||0)+v;
        });
    });
    const $t=$("#resultTable thead tr.totals-row").children();
    $t.each(function(i){
        if(i===0) $(this).text("總計");
        else{
            const v=sums[i];
            $(this).text(v? v.toLocaleString(): "");
        }
    });
}
</script>
</body>
</html>