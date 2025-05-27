<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AnnualReport.aspx.cs" Inherits="YourNamespace.AnnualReport" %>
<!DOCTYPE html>
<html lang="zh-TW">
<head runat="server">
    <meta charset="utf-8" />
    <title>年度查詢報表</title>
    <!-- 引入 Bootstrap CSS 與 jQuery 庫（可使用CDN） -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" />
    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
    <style>
        /* 固定表頭：使用 position: sticky 讓 <th> 固定在頂端2*/
        th.sticky-header { position: sticky; top: 0; background: #ccc; z-index: 10; }
        th.first-header { top: 0; z-index: 30; }    /* 第一列表頭 */
        th.second-header { top: 50px; z-index: 20; } /* 第二列表頭 (約略偏移第一列高度) */
        th.third-header { top: 100px; z-index: 10; } /* 第三列表頭 */
        /* 調整表頭文字置中對齊 */
        th { text-align: center; }
    </style>
</head>
<body>
    <form id="form1" runat="server" class="container-fluid mt-3">
        <!-- 查詢條件區塊 -->
        <div class="form-row">
            <div class="form-group col-md-3">
                <label for="deptSelect">部門：</label>
                <select id="deptSelect" name="deptSelect" class="form-control">
                    <option value="All">All</option>
                    <option value="S100">S100</option>
                    <option value="S200">S200</option>
                </select>
            </div>
            <div class="form-group col-md-3">
                <label for="businessSelect">業務：</label>
                <select id="businessSelect" name="businessSelect" class="form-control">
                    <option value="All">All</option>
                    <option value="Alice">Alice</option><!-- 範例業務名 -->
                </select>
            </div>
            <div class="form-group col-md-3">
                <label for="regionSelect">區域：</label>
                <select id="regionSelect" name="regionSelect" class="form-control">
                    <option value="All">All</option>
                    <option value="TW">Taiwan (TW)</option>
                    <option value="CN">China+Hong Kong (CN)</option>
                    <option value="JP">Japan (JP)</option>
                </select>
            </div>
            <div class="form-group col-md-3">
                <label>查詢年份：</label><br/>
                <!-- 最近四年的複選框，預設當前年勾選 -->
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" name="year" id="year2022" value="2022">
                    <label class="form-check-label" for="year2022">2022</label>
                </div>
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" name="year" id="year2023" value="2023">
                    <label class="form-check-label" for="year2023">2023</label>
                </div>
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" name="year" id="year2024" value="2024">
                    <label class="form-check-label" for="year2024">2024</label>
                </div>
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" name="year" id="year2025" value="2025">
                    <label class="form-check-label" for="year2025">2025</label>
                </div>
            </div>
        </div>
        <div class="form-row">
            <div class="form-group col-md-2">
                <label for="customerSelect">客戶：</label>
                <select id="customerSelect" name="customerSelect" class="form-control">
                    <option value="All">All</option>
                    <option value="VI">VI</option>
                    <option value="EPS">EPS</option>
                </select>
            </div>
            <div class="form-group col-md-2">
                <label for="itemSelect">Cust Item：</label>
                <select id="itemSelect" name="itemSelect" class="form-control">
                    <option value="All">All</option>
                    <option value="Internal">Internal</option>
                    <option value="External">External</option>
                </select>
            </div>
            <div class="form-group col-md-2">
                <label for="substrateSelect">Substrate：</label>
                <select id="substrateSelect" name="substrateSelect" class="form-control">
                    <option value="All">All</option>
                    <option value="SubA">SubA</option>
                    <option value="SubB">SubB</option>
                </select>
            </div>
            <!-- 顯示欄位控制核取方塊 -->
            <div class="form-group col-md-6">
                <label class="mr-2">顯示欄位：</label>
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" id="chkThk1" checked>
                    <label class="form-check-label" for="chkThk1">Thk1</label>
                </div>
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" id="chkThk2" checked>
                    <label class="form-check-label" for="chkThk2">Thk2</label>
                </div>
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" id="chkThk3" checked>
                    <label class="form-check-label" for="chkThk3">Thk3</label>
                </div>
                <div class="form-check form-check-inline ml-4">
                    <input class="form-check-input" type="checkbox" id="chkRes1" checked>
                    <label class="form-check-label" for="chkRes1">Res1</label>
                </div>
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" id="chkRes2" checked>
                    <label class="form-check-label" for="chkRes2">Res2</label>
                </div>
                <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" id="chkRes3" checked>
                    <label class="form-check-label" for="chkRes3">Res3</label>
                </div>
                <!-- 查詢和匯出按鈕 -->
                <button type="button" id="btnSearch" class="btn btn-primary ml-4">查詢</button>
                <asp:Button ID="btnExport" runat="server" CssClass="btn btn-success ml-2" Text="匯出 Excel" />
            </div>
        </div>

        <!-- 報表表格 -->
        <table id="reportTable" class="table table-bordered table-striped">
            <thead></thead>
            <tbody></tbody>
        </table>
    </form>
    <!-- 引入 Bootstrap JS (如果需要) -->
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
</body>
</html>
