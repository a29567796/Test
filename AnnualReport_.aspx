<%@ Page Language="C#" AutoEventWireup="true"
    CodeBehind="AnnualReport.aspx.cs"
    Inherits="WebApp.AnnualReport" %>

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
    <script      src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

    <style>
        /* ---------- Sticky 三層表頭 ---------- */
        thead tr:nth-child(1) th { position:sticky; top:0;    z-index:3; background:#f2f2f2; }
        thead tr:nth-child(2) th { position:sticky; top:40px; z-index:2; background:#fafafa; }
        thead tr:nth-child(3) th { position:sticky; top:80px; z-index:1; background:#fff; }

        /* 去除固定欄上方灰塊 & 框線 */
        .fc-blank{ background:#fff!important; border:none!important; }

        /* row3 固定欄標題 (淺灰) */
        .fc{ background:#f2f2f2!important; }

        /* 年度色塊 (至多四年) */
        .year-0{background:#92D050;} .year-1{background:#00B0F0;}
        .year-2{background:#FFC000;} .year-3{background:#A9A9A9;}

        /* 捲動區 */
        .table-responsive{ max-height:700px; overflow:auto; }

        /* 統一欄寬 */
        #resultTable th,#resultTable td{
            min-width:100px; max-width:200px;
            white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
        }
    </style>
</head>
<body>
<form id="form1" runat="server" class="container-fluid my-3">

    <!-- ===== 查詢條件 (省略其餘欄位, 如需調整可自行擴充) ===== -->
    <div class="mb-2">
        <label class="mr-2">Year：</label>
        <label class="mr-1"><input type="checkbox" name="year" value="2022">2022</label>
        <label class="mr-1"><input type="checkbox" name="year" value="2023">2023</label>
        <label class="mr-1"><input type="checkbox" name="year" value="2024">2024</label>
        <label class="mr-1"><input type="checkbox" name="year" value="2025">2025</label>

        <button id="btnQuery" type="button" class="btn btn-primary btn-sm ml-3">查詢</button>
        <asp:Button ID="btnExport" runat="server" Text="匯出 Excel"
            CssClass="btn btn-success btn-sm ml-2" OnClick="btnExport_Click" />
    </div>

    <!-- Thk/Res 切換 -->
    <div class="mb-3">
        <label class="mr-2">顯示欄位：</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk1" checked />Thk1</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res1" checked />Res1</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk2" checked />Thk2</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res2" checked />Res2</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk3" checked />Thk3</label>
        <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res3" checked />Res3</label>
    </div>

    <!-- ===== 報表表格 ===== -->
    <div class="table-responsive">
        <table id="resultTable" class="table table-bordered table-sm">
            <thead></thead>
            <tbody></tbody>
        </table>
    </div>

</form>

<!-- ================= 前端腳本 ================= -->
<script>
$(function(){
    /* 預設勾選今年 */
    const y = new Date().getFullYear().toString();
    $(`input[name=year][value=${y}]`).prop("checked",true);

    /* 查詢 */
    $("#btnQuery").on("click", queryReport);

    /* 切換固定欄顯示 */
    $(".col-toggle").on("change", function(){
        $("th."+this.value+", td."+this.value).toggle(this.checked);
    });
});

/* ------ AJAX 取數 ------ */
function queryReport(){
    const years = $("input[name=year]:checked").map((_,x)=>+x.value).get();
    if(!years.length){ alert("請選年份"); return; }

    $.ajax({
        type:"POST", url:"AnnualReport.aspx/GetReportData",
        contentType:"application/json; charset=utf-8", dataType:"json",
        data:JSON.stringify({ yearList:years }),
        success:r=>{
            const data = r.d || r;
            buildHeader(years);
            buildBody(data,years);
            // 同步核取
            $(".col-toggle").each(function(){
                if(!this.checked) $("th."+this.value+", td."+this.value).hide();
            });
        },
        error:e=>{ console.error(e); alert("查詢失敗"); }
    });
}

/* 固定欄定義 */
const fixed = [
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

/* ------ 表頭 ------ */
function buildHeader(years){
    const $th = $("#resultTable thead").empty();

    /* row1 */
    let r1="<tr>";
    fixed.forEach(f=> r1+=`<th class="${f.c} fc-blank"></th>`);
    years.forEach((y,i)=> r1+=`<th colspan="17" class="year-${i%4}">${y}</th>`);
    r1+="</tr>";

    /* row2 */
    let r2="<tr>";
    fixed.forEach(f=> r2+=`<th class="${f.c} fc-blank"></th>`);
    years.forEach(()=> r2+=
      "<th colspan='4'>Q1</th><th colspan='4'>Q2</th>"+
      "<th colspan='4'>Q3</th><th colspan='4'>Q4</th>"+
      "<th rowspan='2'>Total</th>");
    r2+="</tr>";

    /* row3 */
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

/* ------ 表體 ------ */
function buildBody(list, years){
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
            html.push(`
              <td>${r[p+"Jan"]}</td><td>${r[p+"Feb"]}</td><td>${r[p+"Mar"]}</td><td>${r[p+"Q1_Total"]}</td>
              <td>${r[p+"Apr"]}</td><td>${r[p+"May"]}</td><td>${r[p+"Jun"]}</td><td>${r[p+"Q2_Total"]}</td>
              <td>${r[p+"Jul"]}</td><td>${r[p+"Aug"]}</td><td>${r[p+"Sep"]}</td><td>${r[p+"Q3_Total"]}</td>
              <td>${r[p+"Oct"]}</td><td>${r[p+"Nov"]}</td><td>${r[p+"Dec"]}</td><td>${r[p+"Q4_Total"]}</td>
              <td>${r[p+"Total"]}</td>`);
        });
        html.push("</tr>");
    });
    $tb.html(html.join(""));
}
</script>
</body>
</html>
