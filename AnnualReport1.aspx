<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AnnualReport1.aspx.cs" Inherits="CRM.AnnualReport1" %>

<!DOCTYPE html>
<html lang="zh-TW">
<head runat="server">
<meta charset="utf-8" />
<title>年度查詢報表</title>

<!-- Bootstrap & jQuery -->
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" />
<script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>

<!-- jQuery UI：日期選擇器 -->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css" />
<script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>

<style>
    /* Sticky 表頭 */
    thead tr.totals-row   th{position:sticky;top:0;   z-index:4;background:#fff;color:#007bff;}
    thead tr:nth-child(2) th{position:sticky;top:48px;z-index:3;background:#f2f2f2;}
    thead tr:nth-child(3) th{position:sticky;top:96px;z-index:2;background:#fafafa;}
    thead tr:nth-child(4) th{position:sticky;top:144px;z-index:1;background:#fff;}

    /* 固定欄位色 */
    .fc-blank{background:#fff!important;}
    .fc      {background:#f2f2f2!important;}

    /* 年度色塊 */
    .year-0{background:#92D050;}
    .year-1{background:#00B0F0;}
    .year-2{background:#FFC000;}
    .year-3{background:#A9A9A9;}

    /* 捲動 */
    .table-responsive{max-height:700px;overflow-y:auto;}

    /* 固定寬 */
    .w120{width:120px;} .w100{width:100px;} .w80{width:80px;} .w60{width:60px;}

    /* 表格外觀 */
    .table thead th{vertical-align:bottom;border-bottom:0;}
    .table-bordered td,.table-bordered th{border:0;}
    .table-sm td,.table-sm th{padding:.75rem;}

    /* 兩表使用相同佈局 */
    #headerTable,#bodyTable{table-layout:fixed;width:100%;}

    /* (B) 每格 60~150px，可換行 */
    .table-wrapper th,
    .table-wrapper td{
        min-width:60px;
        max-width:150px;
        white-space:normal;
        word-break:break-word;
    }
</style>
</head>
<body>
<form id="form1" runat="server" class="container-fluid my-3">

<!-- ===== 查詢條件 (略‧內容同前) ===== -->
<!-- ……(上方條件區塊與之前版本相同，為節省篇幅未改動，故省略)…… -->

<!-- ===== 表頭 + 表身 ===== -->
<div class="table-wrapper">
  <table id="headerTable" class="table table-bordered table-sm mb-0">
    <thead></thead>
  </table>
  <div class="table-responsive">
    <table id="bodyTable" class="table table-bordered table-sm">
      <tbody></tbody>
    </table>
  </div>
</div>
</form>

<script>
$(function(){
    $(".datepicker").datepicker({dateFormat:"yy/mm/dd"});
    const y=new Date().getFullYear().toString();
    $(`input[name=year][value=${y}]`).prop("checked",true);

    $("#btnQuery").on("click",queryReport);

    $(".col-toggle").on("change",function(){
        const cls=$(this).val();
        $("th."+cls+", td."+cls).toggle($(this).is(":checked"));
        syncHeaderWidths();               // A-1: 欄位顯示切換後重新對齊
    });

    $(window).on("resize",syncHeaderWidths); // A-2: 視窗變更時保持對齊
});

/* ---------- AJAX 取得資料 ---------- */
function queryReport(){
    const years=$("input[name=year]:checked").map((_,el)=>+el.value).get();
    if(!years.length){alert("請至少選一年");return;}
    const viewType=$("input[name=viewType]:checked").val();

    $.ajax({
        type:"POST",
        url:"AnnualReport1.aspx/GetReportData",
        contentType:"application/json; charset=utf-8",
        dataType:"json",
        data:JSON.stringify({
            department:$("#ddlDept").val(),
            business:$("#ddlBiz").val(),
            region:$("#ddlRegion").val(),
            yearList:years,
            customer:$("#ddlCustomer").val(),
            custItem:$("#txtCustItem").val(),
            substrate:$("#ddlSubstrate").val(),
            productType:$("#ddlProductType").val(),
            shipTo:$("#txtShipTo").val(),
            realDate:$("#txtRealDate").val(),
            status:$("#ddlStatus").val(),
            orderType:$("#ddlOrderType").val()
        }),
        success:r=>{
            const data=r.d||r;
            buildHeader(years,viewType);
            buildBody(data,years,viewType);
            updateTotals();
            syncHeaderWidths();            // A-3: 產生完成後立即對齊
            $(".col-toggle").each(function(){
                if(!$(this).is(":checked"))
                    $("th."+this.value+", td."+this.value).hide();
            });
        },
        error:e=>{console.error(e);alert("查詢失敗");}
    });
}

/* ---------- 固定欄資料 (同前) ---------- */
const fixedCols=[ /* …省略，與前版相同… */ ];

/* ---------- 表頭 & 表身 (buildHeader / buildBody 與前版一致) ---------- */
/* ……為簡潔起見，這兩隻函式內容完全沿用前一版，無需修改…… */

/* ---------- (A) 同步兩表欄寬 ---------- */
function syncHeaderWidths(){
    const $bodyRow=$("#bodyTable tbody tr:first");
    if(!$bodyRow.length) return;

    // 1) 依第一筆資料列寬度套用到所有 thead cell
    const widths=[];
    $bodyRow.children().each(function(i){widths[i]=$(this).outerWidth();});
    $("#headerTable thead tr").each(function(){
        $(this).children().each(function(i){
            $(this).css("width",widths[i]);
        });
    });

    // 2) 考量垂直 scrollbar：讓 headerTable 右邊留出滾動條寬度
    const $wrap=$("#bodyTable").closest(".table-responsive")[0];
    const scrollbar=$wrap.offsetWidth-$wrap.clientWidth; // 0 或 ~17px
    $("#headerTable").css("margin-right",scrollbar+"px");
}

/* ---------- 合計 (updateTotals) ---------- */
function updateTotals(){
    const $rows=$("#bodyTable tbody tr");
    if(!$rows.length) return;
    const cnt=$rows.first().children().length;
    const sums=Array(cnt).fill(0);
    $rows.each(function(){
        $(this).children().each(function(i){
            const v=parseFloat($(this).text().replace(/,/g,""));
            if(!isNaN(v)) sums[i]+=v;
        });
    });
    $("#headerTable thead tr.totals-row").children().each(function(i){
        if(i===0) $(this).text("總計");
        else{
            const v=sums[i];
            $(this).text(v? v.toLocaleString():"");
        }
    });
}
</script>
</body>
</html>
