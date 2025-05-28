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
/* totals 列永遠置頂 */
thead tr.totals-row th{position:sticky;top:0;z-index:5;background:#fff;color:#007bff;}
/* 其餘頂部由 JS 動態設定 top */

/* 固定欄樣式 */
.fc-blank{background:#fff!important;border:none!important;}
.fc{background:#f2f2f2!important;}

/* 年度色塊 (至多四年循環) */
.year-0{background:#92D050;} .year-1{background:#00B0F0;}
.year-2{background:#FFC000;} .year-3{background:#A9A9A9;}

.table-responsive{max-height:700px;overflow:auto;}

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

<!-- ===== 查詢列 ===== -->
<div class="mb-2">
  <!-- 年份多選 -->
  <label class="mr-2">年份：</label>
  <label class="mr-1"><input type="checkbox" name="year" value="2023"/>2023</label>
  <label class="mr-1"><input type="checkbox" name="year" value="2024"/>2024</label>
  <label class="mr-1"><input type="checkbox" name="year" value="2025"/>2025</label>

  <!-- 顯示維度 -->
  <label class="ml-4 mr-1">顯示：</label>
  <select id="ddlLevel" class="custom-select custom-select-sm w-auto">
    <option value="Month" selected>Month</option>
    <option value="Season">Season</option>
    <option value="Year">Year</option>
  </select>

  <!-- 按鈕 -->
  <button id="btnQuery" type="button" class="btn btn-primary btn-sm ml-3">查詢</button>
  <asp:Button ID="btnExport" runat="server" Text="匯出 Excel"
              CssClass="btn btn-success btn-sm ml-2" OnClick="btnExport_Click"/>
</div>

<!-- Thk/Res 顯示切換 -->
<div class="mb-3">
  <label class="mr-2">顯示欄位：</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk1" checked/>Thk1</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk2" checked/>Thk2</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-thk3" checked/>Thk3</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res1" checked/>Res1</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res2" checked/>Res2</label>
  <label class="mr-2"><input type="checkbox" class="col-toggle" value="col-res3" checked/>Res3</label>
</div>

<!-- ===== 報表表格 ===== -->
<div class="table-responsive">
  <table id="resultTable" class="table table-bordered table-sm">
    <thead></thead><tbody></tbody>
  </table>
</div>
</form>

<script>
/* === 初始化 === */
$(function(){
  // 預設勾選今年
  const y=new Date().getFullYear().toString();
  $(`input[name=year][value=${y}]`).prop('checked',true);

  $("#btnQuery").on('click',queryReport);
  $(".col-toggle").on('change',function(){
     $("."+this.value).toggle(this.checked);
  });
});

/* === 固定欄定義 === */
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

/* === 查詢 === */
function queryReport(){
  const years=$("input[name=year]:checked").map((_,x)=>+x.value).get();
  if(!years.length){alert("請選年份");return;}
  const level=$("#ddlLevel").val(); // Month / Season / Year

  $.ajax({
    type:"POST",url:"AnnualReport1.aspx/GetReportData",
    contentType:"application/json; charset=utf-8",dataType:"json",
    data:JSON.stringify({yearList:years}),
    success:r=>{
      const data=r.d||r;
      buildHeader(years,level);
      buildBody(data,years,level);
      updateTotals();
      $(".col-toggle").each(function(){ if(!this.checked) $("."+this.value).hide();});
      applySticky(); // 動態計算 top
    },
    error:e=>{console.error(e);alert("查詢失敗");}
  });
}

/* === 動態表頭 === */
function buildHeader(years,level){
  const $thead=$("#resultTable thead").empty();

  /* 0 : totals */
  let tot="<tr class='totals-row'>";
  fixed.forEach(()=>tot+="<th></th>");
  years.forEach(()=>{ tot+=genBlankCells(level); });
  tot+="</tr>"; $thead.append(tot);

  if(level==="Month"){
      /* 年 (h1) */
      let h1="<tr>";
      fixed.forEach(()=>h1+="<th class='fc-blank'></th>");
      years.forEach((y,i)=>h1+=`<th colspan='17' class='year-${i%4}'>${y}</th>`);
      h1+="</tr>"; $thead.append($(h1).addClass('h1'));

      /* 季群組 (h2) */
      let h2="<tr>";
      fixed.forEach(()=>h2+="<th class='fc-blank'></th>");
      years.forEach(()=>h2+=
        "<th colspan='4'>Q1</th><th colspan='4'>Q2</th>"+
        "<th colspan='4'>Q3</th><th colspan='4'>Q4</th>"+
        "<th rowspan='1'>Total</th>");
      h2+="</tr>"; $thead.append($(h2).addClass('h2'));

      /* 月/季Total 列 (h3) */
      let h3="<tr>";
      fixed.forEach(f=>h3+=`<th class='fc'>${f.t}</th>`);
      years.forEach(()=>h3+=`
        <th>Jan</th><th>Feb</th><th>Mar</th><th>Q1T</th>
        <th>Apr</th><th>May</th><th>Jun</th><th>Q2T</th>
        <th>Jul</th><th>Aug</th><th>Sep</th><th>Q3T</th>
        <th>Oct</th><th>Nov</th><th>Dec</th><th>Q4T</th>
        <th>Total</th>`);
      h3+="</tr>"; $thead.append($(h3).addClass('h3'));
  }
  else if(level==="Season"){
      /* 年列 (h1) */
      let h1="<tr>";
      fixed.forEach(()=>h1+="<th class='fc-blank'></th>");
      years.forEach((y,i)=>h1+=`<th colspan='5' class='year-${i%4}'>${y}</th>`);
      h1+="</tr>"; $thead.append($(h1).addClass('h1'));

      /* Q1T..Total (h3) */
      let h3="<tr>";
      fixed.forEach(f=>h3+=`<th class='fc'>${f.t}</th>`);
      years.forEach(()=>h3+="<th>Q1T</th><th>Q2T</th><th>Q3T</th><th>Q4T</th><th>Total</th>");
      h3+="</tr>"; $thead.append($(h3).addClass('h3'));
  }
  else{ /* Year */
      let h3="<tr>";
      fixed.forEach(f=>h3+=`<th class='fc'>${f.t}</th>`);
      years.forEach((y,i)=>h3+=`<th class='year-${i%4}'>${y} Total</th>`);
      h3+="</tr>"; $thead.append($(h3).addClass('h3'));
  }
}

/* 依 level 給空白列數 */
function genBlankCells(level){
  return (level==="Month")?"<th></th>".repeat(17):
         (level==="Season")?"<th></th>".repeat(5):
         "<th></th>";
}

/* === 建立表體 === */
function buildBody(list,years,level){
  const $tb=$("#resultTable tbody").empty();
  if(!list.length){$tb.html("<tr><td colspan='999' class='text-center'>-- 無資料 --</td></tr>");return;}

  const sb=[];
  list.forEach(r=>{
    sb.push("<tr>");
    // 固定欄
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
       if(level==="Month"){
         sb.push(`
           <td>${r[p+"Jan"]}</td><td>${r[p+"Feb"]}</td><td>${r[p+"Mar"]}</td><td>${r[p+"Q1_Total"]}</td>
           <td>${r[p+"Apr"]}</td><td>${r[p+"May"]}</td><td>${r[p+"Jun"]}</td><td>${r[p+"Q2_Total"]}</td>
           <td>${r[p+"Jul"]}</td><td>${r[p+"Aug"]}</td><td>${r[p+"Sep"]}</td><td>${r[p+"Q3_Total"]}</td>
           <td>${r[p+"Oct"]}</td><td>${r[p+"Nov"]}</td><td>${r[p+"Dec"]}</td><td>${r[p+"Q4_Total"]}</td>
           <td>${r[p+"Total"]}</td>`);
       }else if(level==="Season"){
         sb.push(`<td>${r[p+"Q1_Total"]}</td><td>${r[p+"Q2_Total"]}</td><td>${r[p+"Q3_Total"]}</td><td>${r[p+"Q4_Total"]}</td><td>${r[p+"Total"]}</td>`);
       }else{
         sb.push(`<td>${r[p+"Total"]}</td>`);
       }
    });
    sb.push("</tr>");
  });
  $tb.html(sb.join(""));
}

/* === 合計列 === */
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
    th.textContent= i===0? "總計": (sums[i]?sums[i].toLocaleString():"");
  });
}

/* === 動態設定 sticky top === */
function applySticky(){
  const rows=$("#resultTable thead tr:not(.totals-row)");
  rows.each(function(idx){
     const offset=(idx+1)*48; // totals-row 已占 0
     $(this).find("th").css({position:"sticky",top:offset+"px", "z-index":4-idx});
  });
}
</script>
</body>
</html>