<!-- ... 上方省略的查詢條件表單 ... -->
<div class="table-responsive">
  <table id="resultTable" class="table table-bordered table-striped">
    <thead></thead>
    <tbody></tbody>
  </table>
</div>

<script>
$(function(){
  // 綁定日期選擇器
  $(".datepicker").datepicker({ dateFormat: "yy/mm/dd" });
  
  // 綁定查詢按鈕事件
  $("#btnQuery").click(function(e){
    e.preventDefault();  // 防止表單提交
    // 收集選擇的年份 (假設 #ddlYear 是 <select multiple>)
    let selectedYears = $("#ddlYear").val() || [];
    // 調用前述 AJAX 函式
    $.ajax({
      type: "POST",
      url: "AnnualReport.aspx/GetReportData",
      data: JSON.stringify({
        department: $("#ddlDept").val(),
        business: $("#ddlBiz").val(),
        region: $("#ddlRegion").val(),
        yearList: selectedYears.map(y => parseInt(y)), // 傳整數陣列
        customer: $("#txtCustomer").val(),
        custItem: $("#txtCustItem").val(),
        substrate: $("#ddlSubstrate").val(),
        productType: $("#ddlProductType").val(),
        shipTo: $("#txtShipTo").val(),
        realDate: $("#txtRealDate").val(),
        status: $("#ddlStatus").val(),
        orderType: $("#ddlOrderType").val()
      }),
      contentType: "application/json; charset=utf-8",
      dataType: "json",
      success: function(res) {
        let data = res.d || res;
        // 動態建立表頭
        buildTableHeader(selectedYears);
        // 將資料填入表格
        buildTableBody(data, selectedYears);
        // 顯示表頭下方資料（若有必要初始化隱藏某些欄位）
        // 根據當前核取方塊狀態隱藏相應欄位
        $(".col-toggle").each(function(){
          if (!$(this).is(":checked")) {
            $("th."+this.value+", td."+this.value).hide();
          }
        });
      },
      error: function(err) {
        alert("查詢失敗:" + err.responseText);
      }
    });
  });
  
  // 定義動態建立表頭的函式
  function buildTableHeader(years) {
    const thead = $("#resultTable thead");
    thead.empty();
    // 第一行表頭
    let headerRow1 = "<tr>";
    headerRow1 += '<th rowspan="3" class="col-thk1">Thk1</th>'
               +  '<th rowspan="3" class="col-thk2">Thk2</th>'
               +  '<th rowspan="3" class="col-thk3">Thk3</th>'
               +  '<th rowspan="3" class="col-res1">Res1</th>'
               +  '<th rowspan="3" class="col-res2">Res2</th>'
               +  '<th rowspan="3" class="col-res3">Res3</th>';
    years.forEach((yr, idx) => {
      // 加上年度表頭，附帶背景色類別
      headerRow1 += `<th colspan="17" class="year-header year-${yr}">${yr}</th>`;
    });
    headerRow1 += "</tr>";
    // 第二行表頭 (季度 + 年Total)
    let headerRow2 = "<tr>";
    years.forEach(yr => {
      headerRow2 += '<th colspan="4">Q1</th>'
                 +  '<th colspan="4">Q2</th>'
                 +  '<th colspan="4">Q3</th>'
                 +  '<th colspan="4">Q4</th>'
                 +  '<th rowspan="2">Total</th>';
    });
    headerRow2 += "</tr>";
    // 第三行表頭 (月份 + 季Total)
    let headerRow3 = "<tr>";
    years.forEach(yr => {
      headerRow3 += 
         '<th>Jan</th><th>Feb</th><th>Mar</th><th>Q1 Total</th>'
       + '<th>Apr</th><th>May</th><th>Jun</th><th>Q2 Total</th>'
       + '<th>Jul</th><th>Aug</th><th>Sep</th><th>Q3 Total</th>'
       + '<th>Oct</th><th>Nov</th><th>Dec</th><th>Q4 Total</th>';
      // 年Total不在此行
    });
    headerRow3 += "</tr>";
    // 將三行表頭插入
    thead.append(headerRow1).append(headerRow2).append(headerRow3);
  }
  
  // 定義動態建立表格主體的函式
  function buildTableBody(data, years) {
    const tbody = $("#resultTable tbody");
    tbody.empty();
    if (!data || data.length === 0) {
      return; // 無資料，不處理
    }
    let rowsHtml = "";
    data.forEach(item => {
      rowsHtml += "<tr>";
      // 可選欄位值
      rowsHtml += `<td class="col-thk1">${item.Thk1}</td>`
               +  `<td class="col-thk2">${item.Thk2}</td>`
               +  `<td class="col-thk3">${item.Thk3}</td>`
               +  `<td class="col-res1">${item.Res1}</td>`
               +  `<td class="col-res2">${item.Res2}</td>`
               +  `<td class="col-res3">${item.Res3}</td>`;
      // 年度欄位值
      years.forEach(yr => {
        // 注意：物件的鍵以"年_月"等格式對應 DataTable 欄名
        rowsHtml += 
           `<td>${item[yr + '_Jan']}</td><td>${item[yr + '_Feb']}</td><td>${item[yr + '_Mar']}</td><td>${item[yr + '_Q1_Total']}</td>`
         + `<td>${item[yr + '_Apr']}</td><td>${item[yr + '_May']}</td><td>${item[yr + '_Jun']}</td><td>${item[yr + '_Q2_Total']}</td>`
         + `<td>${item[yr + '_Jul']}</td><td>${item[yr + '_Aug']}</td><td>${item[yr + '_Sep']}</td><td>${item[yr + '_Q3_Total']}</td>`
         + `<td>${item[yr + '_Oct']}</td><td>${item[yr + '_Nov']}</td><td>${item[yr + '_Dec']}</td><td>${item[yr + '_Q4_Total']}</td>`
         + `<td>${item[yr + '_Total']}</td>`;  // 年度總計
      });
      rowsHtml += "</tr>";
    });
    tbody.html(rowsHtml);
  }

  // 綁定欄位顯示切換 (checkbox) - 前述已定義，可重複利用
  $(".col-toggle").change(function() {
    let colClass = $(this).val();
    if ($(this).is(":checked")) {
      $("th."+colClass+", td."+colClass).show();
    } else {
      $("th."+colClass+", td."+colClass).hide();
    }
  });
});
</script>