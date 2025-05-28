<%@ Page Language="C#" AutoEventWireup="true"
    CodeBehind="AnnualReport1.aspx.cs"
    Inherits="CRM.AnnualReport1" %>

<!DOCTYPE html>
<html lang="zh-TW">
<head runat="server">
<meta charset="utf-8"/>
<title>年度查詢報表</title>

<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css"/>
<script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css"/>
<script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

<style>
/* sticky：總計 + 最多三層表頭 */
thead tr.totals-row th{position:sticky;top:0;   z-index:4;background:#fff;color:#007bff;}
thead tr.h1         th{position:sticky;top:48px;z-index:3;background:#f2f2f2;}
thead tr.h2         th{position:sticky;top:96px;z-index:2;background:#fafafa;}
thead tr.h3         th{position:sticky;top:144px;z-index:1;background:#fff;}

.fc-blank{background:#fff!important;border:none!important;}
.fc{background:#f2f2f2!important;}

.year-0{background:#92D050;} .year-1{background:#00B0F0;}
.year-2{background:#FFC000;} .year-3{background:#A9A9A9;}

.table-responsive{max-height:700px;overflow:auto;}
#resultTable th,#resultTable td{min-width:100px;max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}

.table thead th{vertical-align:bottom;border-bottom:0;}
.table-bordered td,.table-bordered th{border:0;}
.table-sm td,.table-sm th{padding:.75rem;}
</style>
</head>
<body>
<form id="form1" runat="server" class="container-fluid my-3">

<!-- ===== 查詢條件 (簡化版：年份複選 + 顯示層級) ===== -->
<div class="mb-2">
  <label class="mr-2">年份：</label>
  <label class="mr-1"><input type="checkbox" name="year" value="2023">2023</label>
  <label class="mr-1"><input type="checkbox" name="year" value="2024">2024</label>
  <label class="mr-1"><input type="checkbox" name="year" value="2025">2025</label>

  <label class="ml-4 mr-1">顯示：</label>
  <select id="ddlLevel" class="custom-select custom-select-sm w-auto">
      <option value="Month" selected>月</option>
      <option value="Season">季</option>
      <option value="Year">年</option>
  </select>

  <button id="btnQuery" type="button" class="btn btn-primary btn-sm ml-3">查詢</button>
  <asp:Button ID="btnExport" runat="server" CssClass="btn btn-success btn-sm ml-2"
      Text="匯出 Excel" OnClick="btnExport_Click"/>
</div>

<!-- ===== 顯示 / 隱藏欄位 ===== -->
<div class="mb-3">
  <label class="mr-2">顯示欄位：</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk1" checked/>Thk1</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk2" checked/>Thk2</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk3" checked/>Thk3</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res1" checked/>Res1</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res2" checked/>Res2</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res3" checked/>Res3</label>
</div>

<!-- ===== 報表 ===== -->
<div class="table-responsive">
  <table id="resultTable" class="table table-bordered table-sm">
      <thead></thead><tbody></tbody>
  </table>
</div>
</form>

<script>
/* ===== 初始事件 ===== */
$(function(){
  const y=new Date().getFullYear().toString();
  $(`input[name=year][value=${y}]`).prop("checked",true);

  $("#btnQuery").on("click",queryReport);
  $(".col-toggle").on("change",function(){
      $("th."+this.value+", td."+this.value).toggle(this.checked);
  });
});

/* ===== 固定欄定義 ===== */
const fixed=[
 {t:"Customer",c:"col-customer"},
 {t:"Inch"    ,c:"col-inch"},
 {t:"Cust&nbsp;Item",c:"col-custitem"},
 {t:"OA&nbsp;Item" ,c:"col-oaitem"},
 {t:"Substrate",c:"col-substrate"},
 {t:"Price"   ,c:"col-price"},
 {t:"Thk1"    ,c:"col-thk1"},
 {t:"Res1"    ,c:"col-res1"},
 {t:"Thk2"    ,c:"col-thk2"},
 {t:"Res2"    ,c:"col-res2"},
 {t:"Thk3"    ,c:"col-thk3"},
 {t:"Res3"    ,c:"col-res3"}
];

/* ===== AJAX 查詢 ===== */
function queryReport(){
  const years=$("input[name=year]:checked").map((_,x)=>+x.value).get();
  if(!years.length){alert("請選年份");return;}

  $.ajax({
    type:"POST",url:"AnnualReport1.aspx/GetReportData",
    contentType:"application/json; charset=utf-8",dataType:"json",
    data:JSON.stringify({yearList:years}),
    success:r=>{
      const level=$("#ddlLevel").val();            // Month / Season / Year
      const data=r.d||r;
      buildHeader(years,level);
      buildBody(data,years,level);
      updateTotals();
      $(".col-toggle").each(function(){ if(!this.checked) $("."+this.value).hide(); });
    },
    error:e=>{console.error(e);alert("查詢失敗");}
  });
}

/* ===== 產生表頭 ===== */
function buildHeader(years,level){
  const $thead=$("#resultTable thead").empty();

  /* 上層總計列 (空白待填) */
  let totalRow="<tr class='totals-row'>";
  fixed.forEach(()=>totalRow+="<th></th>");
  years.forEach(()=>{
      const n=(level==="Month")?17:(level==="Season")?5:1;
      for(let i=0;i<n;i++) totalRow+="<th></th>";
  });
  totalRow+="</tr>";
  $thead.append(totalRow);

  if(level==="Month"){                      /* ==== 月 === */
      // row 年
      let rY="<tr class='h1'>";
      fixed.forEach(f=>rY+=`<th class='fc-blank'></th>`);
      years.forEach((y,i)=>rY+=`<th colspan='17' class='year-${i%4}'>${y}</th>`);
      rY+="</tr>";

      // row 季
      let rQ="<tr class='h2'>";
      fixed.forEach(f=>rQ+=`<th class='fc-blank'></th>`);
      years.forEach(()=>rQ+=
        "<th colspan='4'>Q1</th><th colspan='4'>Q2</th>"+
        "<th colspan='4'>Q3</th><th colspan='4'>Q4</th>"+
        "<th rowspan='1'>Total</th>");
      rQ+="</tr>";

      // row 月
      let rM="<tr class='h3'>";
      fixed.forEach(f=>rM+=`<th class='fc'>${f.t}</th>`);
      years.forEach(()=>rM+=`
        <th>Jan</th><th>Feb</th><th>Mar</th><th>Q1&nbsp;T</th>
        <th>Apr</th><th>May</th><th>Jun</th><th>Q2&nbsp;T</th>
        <th>Jul</th><th>Aug</th><th>Sep</th><th>Q3&nbsp;T</th>
        <th>Oct</th><th>Nov</th><th>Dec</th><th>Q4&nbsp;T</th>`);
      rM+="</tr>";

      $thead.append(rY,rQ,rM);
  }
  else if(level==="Season"){               /* ==== 季 === */
      // row 年
      let rY="<tr class='h1'>";
      fixed.forEach(f=>rY+=`<th class='fc-blank'></th>`);
      years.forEach((y,i)=>rY+=`<th colspan='5' class='year-${i%4}'>${y}</th>`);
      rY+="</tr>";

      // row Q1T..Total
      let rS="<tr class='h3'>";
      fixed.forEach(f=>rS+=`<th class='fc'>${f.t}</th>`);
      years.forEach(()=>rS+="<th>Q1T</th><th>Q2T</th><th>Q3T</th><th>Q4T</th><th>Total</th>");
      rS+="</tr>";
      $thead.append(rY,rS);
  }
  else{                                    /* ==== 年 === */
      // 只有 Total 列
      let rY="<tr class='h3'>";
      fixed.forEach(f=>rY+=`<th class='fc'>${f.t}</th>`);
      years.forEach((y,i)=>rY+=`<th class='year-${i%4}'>${y} Total</th>`);
      rY+="</tr>";
      $thead.append(rY);
  }
}

/* ===== 產生表體 ===== */
function buildBody(list,years,level){
  const $tb=$("#resultTable tbody").empty();
  if(!list.length){$tb.html("<tr><td colspan='999' class='text-center'>-- 無資料 --</td></tr>");return;}

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
      if(level==="Month"){
          html.push(`
            <td>${r[p+"Jan"]}</td><td>${r[p+"Feb"]}</td><td>${r[p+"Mar"]}</td><td>${r[p+"Q1_Total"]}</td>
            <td>${r[p+"Apr"]}</td><td>${r[p+"May"]}</td><td>${r[p+"Jun"]}</td><td>${r[p+"Q2_Total"]}</td>
            <td>${r[p+"Jul"]}</td><td>${r[p+"Aug"]}</td><td>${r[p+"Sep"]}</td><td>${r[p+"Q3_Total"]}</td>
            <td>${r[p+"Oct"]}</td><td>${r[p+"Nov"]}</td><td>${r[p+"Dec"]}</td><td>${r[p+"Q4_Total"]}</td>
            <td>${r[p+"Total"]}</td>`);
      }else if(level==="Season"){
          html.push(`<td>${r[p+"Q1_Total"]}</td><td>${r[p+"Q2_Total"]}</td><td>${r[p+"Q3_Total"]}</td><td>${r[p+"Q4_Total"]}</td><td>${r[p+"Total"]}</td>`);
      }else{ // Year
          html.push(`<td>${r[p+"Total"]}</td>`);
      }
    });
    html.push("</tr>");
  });
  $tb.html(html.join(""));
}

/* ===== 合計列 ===== */
function updateTotals(){
  const $rows=$("#resultTable tbody tr");
  if(!$rows.length) return;
  const sums=Array($rows.first().children().length).fill(0);
  $rows.each(function(){
     $(this).children().each((i,td)=>{
        const v=parseFloat(td.textContent.replace(/,/g,""));
        if(!isNaN(v)) sums[i]+=v;
     });
  });
  $("#resultTable thead tr.totals-row th").each((i,th)=>{
      th.textContent=i===0?"總計":(sums[i]?sums[i].toLocaleString():"");
  });
}
</script>
</body>
</html>