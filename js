<script type="text/javascript">
$(document).ready(function () {
    // 頁面載入時預設勾選當前年份
    var currentYear = new Date().getFullYear().toString();
    $("input[name='year'][value='" + currentYear + "']").prop("checked", true);

    // 查詢按鈕點擊事件
    $("#btnSearch").on("click", function () {
        // 收集查詢條件值
        var dept = $("#deptSelect").val();
        var business = $("#businessSelect").val();
        var region = $("#regionSelect").val();
        // 收集多選年份
        var years = [];
        $("input[name='year']:checked").each(function () {
            years.push(parseInt(this.value));
        });
        if (years.length === 0) {
            alert("請選擇至少一個年份");
            return;
        }
        var customer = $("#customerSelect").val();
        var custItem = $("#itemSelect").val();
        var substrate = $("#substrateSelect").val();
        // 發出 AJAX 請求至後端 WebMethod
        $.ajax({
            type: "POST",
            url: "<%= ResolveUrl(\"~/AnnualReport.aspx/GetReportData\") %>",  // WebForm 靜態方法網址
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            data: JSON.stringify({ 
                dept: dept, business: business, region: region, 
                years: years, customer: customer, custItem: custItem, substrate: substrate 
            }),
            beforeSend: function () {
                // 可以在這裡顯示載入動畫等
            },
            success: function (response) {
                // 從回傳結果中解析 JSON 資料
                var resultData = JSON.parse(response.d);  // 解析 JSON 字串為物件陣列5
                // 呼叫函式渲染表格
                renderReportTable(resultData, years);
            },
            error: function (xhr, status, error) {
                console.error("查詢失敗: " + status + " - " + error);
            }
        });
    });

    // 顯示欄位核取方塊變更事件：切換欄位顯示
    $("#chkThk1, #chkThk2, #chkThk3, #chkRes1, #chkRes2, #chkRes3").on("change", function () {
        var colClass = "";
        if (this.id === "chkThk1") colClass = ".col-thk1";
        if (this.id === "chkThk2") colClass = ".col-thk2";
        if (this.id === "chkThk3") colClass = ".col-thk3";
        if (this.id === "chkRes1") colClass = ".col-res1";
        if (this.id === "chkRes2") colClass = ".col-res2";
        if (this.id === "chkRes3") colClass = ".col-res3";
        if ($(this).prop("checked")) {
            // 顯示欄位（使用 CSS 顯示元素）
            $(colClass).show();
        } else {
            // 隱藏欄位
            $(colClass).hide();
        }
    });
});

/**
 * 根據查詢資料和年份清單渲染報表表格
 * @param {Array} data 後端回傳的資料列陣列，每項為物件包含所有欄位值
 * @param {Array} years 選取的年份 (數值陣列)
 */
function renderReportTable(data, years) {
    // 先對年份排序，確保從小到大顯示
    years.sort();
    var colors = ["#92D050", "#00B0F0", "#FFC000", "#A9A9A9"];  // 預設最多四年的顏色
    var $thead = $("#reportTable thead").empty();
    var $tbody = $("#reportTable tbody").empty();

    // 建立表頭三列元素
    var $yearRow = $("<tr></tr>");
    var $quarterRow = $("<tr></tr>");
    var $monthRow = $("<tr></tr>");

    // 靜態欄位列表（含Thk、Res欄位）
    var staticHeaders = ["部門", "業務", "區域", "客戶", "Cust Item", "Substrate", 
                         "Thk1", "Thk2", "Thk3", "Res1", "Res2", "Res3"];
    // 生成靜態欄位表頭 (rowspan=3 固定三列)
    staticHeaders.forEach(function(header) {
        var th = $("<th></th>").text(header).addClass("sticky-header first-header");
        th.attr("rowspan", 3);
        // 如果是可切換顯示的欄位，加對應類別以控制顯示
        var lower = header.toLowerCase();
        if (lower.startsWith("thk") || lower.startsWith("res")) {
            th.addClass("col-" + lower);  // 例如 "col-thk1"
        }
        $yearRow.append(th);
    });

    // 生成年份欄位組表頭 (第一列)
    years.forEach(function(year, index) {
        var thYear = $("<th></th>").text(year).addClass("sticky-header first-header");
        thYear.attr("colspan", 17);  // 每年份欄位組佔17欄 (4季*4 + 年總計)
        thYear.css("background-color", colors[index] || "#ddd");  // 標示年份色塊
        $yearRow.append(thYear);
    });

    // 生成季度群組表頭 (第二列) 及 年總計欄位 (亦在第二列)
    years.forEach(function(year) {
        // 四個季度標頭 (各 colspan=4)
        for (var q = 1; q <= 4; q++) {
            var thQ = $("<th></th>").text("Q" + q).addClass("sticky-header second-header");
            thQ.attr("colspan", 4);
            $quarterRow.append(thQ);
        }
        // 年度總計欄 (rowspan=2 覆蓋第二、三列)
        var thYearTotal = $("<th></th>").text("總計").addClass("sticky-header second-header");
        thYearTotal.attr("rowspan", 2);
        $quarterRow.append(thYearTotal);
    });

    // 生成月份及季度合計表頭 (第三列)
    years.forEach(function(year) {
        // 各季度下的月份與合計
        var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        // Q1: 1~3月
        $monthRow.append( $("<th></th>").text("Jan").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("Feb").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("Mar").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("合計").addClass("sticky-header third-header") );
        // Q2: 4~6月
        $monthRow.append( $("<th></th>").text("Apr").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("May").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("Jun").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("合計").addClass("sticky-header third-header") );
        // Q3: 7~9月
        $monthRow.append( $("<th></th>").text("Jul").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("Aug").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("Sep").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("合計").addClass("sticky-header third-header") );
        // Q4: 10~12月
        $monthRow.append( $("<th></th>").text("Oct").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("Nov").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("Dec").addClass("sticky-header third-header") );
        $monthRow.append( $("<th></th>").text("合計").addClass("sticky-header third-header") );
        // （年度總計欄在第二列已佔位，無需在第三列添加空欄位）
    });

    // 組裝三列表頭進入 thead
    $thead.append($yearRow).append($quarterRow).append($monthRow);

    // 生成資料列
    data.forEach(function(item) {
        var $tr = $("<tr></tr>");
        // 插入靜態欄位資料 (依序與 staticHeaders 對應)
        staticHeaders.forEach(function(col) {
            var colKey = col; // 鍵名與表頭文字相同
            // 注意：若資料鍵名不同，需要在後端統一鍵名。此處假定 data 物件的屬性名稱與表頭一致
            var td = $("<td></td>").text(item[colKey] !== undefined ? item[colKey] : "");
            // 可切換欄位加對應class
            var lower = col.toLowerCase();
            if (lower.startsWith("thk") || lower.startsWith("res")) {
                td.addClass("col-" + lower);
            }
            $tr.append(td);
        });
        // 插入動態年份欄位資料
        years.forEach(function(year) {
            // 逐月欄位
            $tr.append( $("<td></td>").text(item["Jan_" + year]).addClass("col-jan_"+year) );
            $tr.append( $("<td></td>").text(item["Feb_" + year]).addClass("col-feb_"+year) );
            $tr.append( $("<td></td>").text(item["Mar_" + year]).addClass("col-mar_"+year) );
            $tr.append( $("<td></td>").text(item["Q1_" + year]).addClass("col-q1_"+year) );   // 第一季合計
            $tr.append( $("<td></td>").text(item["Apr_" + year]).addClass("col-apr_"+year) );
            $tr.append( $("<td></td>").text(item["May_" + year]).addClass("col-may_"+year) );
            $tr.append( $("<td></td>").text(item["Jun_" + year]).addClass("col-jun_"+year) );
            $tr.append( $("<td></td>").text(item["Q2_" + year]).addClass("col-q2_"+year) );   // 第二季合計
            $tr.append( $("<td></td>").text(item["Jul_" + year]).addClass("col-jul_"+year) );
            $tr.append( $("<td></td>").text(item["Aug_" + year]).addClass("col-aug_"+year) );
            $tr.append( $("<td></td>").text(item["Sep_" + year]).addClass("col-sep_"+year) );
            $tr.append( $("<td></td>").text(item["Q3_" + year]).addClass("col-q3_"+year) );   // 第三季合計
            $tr.append( $("<td></td>").text(item["Oct_" + year]).addClass("col-oct_"+year) );
            $tr.append( $("<td></td>").text(item["Nov_" + year]).addClass("col-nov_"+year) );
            $tr.append( $("<td></td>").text(item["Dec_" + year]).addClass("col-dec_"+year) );
            $tr.append( $("<td></td>").text(item["Q4_" + year]).addClass("col-q4_"+year) );   // 第四季合計
            // 年度總計
            $tr.append( $("<td></td>").text(item["Total_" + year]).addClass("col-total_"+year) );
        });
        $tbody.append($tr);
    });

    // 初始化欄位顯示：依據當前核取方塊狀態決定哪些欄位隱藏
    ["thk1","thk2","thk3","res1","res2","res3"].forEach(function(key) {
        if (!$("#chk" + key.charAt(0).toUpperCase() + key.slice(1)).prop("checked")) {
            $(".col-" + key).hide();
        }
    });
}
</script>
