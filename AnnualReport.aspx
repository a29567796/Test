<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AnnualReport.aspx.cs" Inherits="AnnualReport" %>
<!DOCTYPE html>
<html>
<head runat="server">
    <title>Annual Report</title>
    <style>
        /* 2. Uniform column width for all table cells */
        table#reportTable th, table#reportTable td {
            width: 100px;
        }
        /* 3. Header background: remove gray from first two rows, keep on third */
        table#reportTable thead tr:nth-child(1) th,
        table#reportTable thead tr:nth-child(2) th {
            background-color: transparent !important;
        }
        table#reportTable thead tr:nth-child(3) th {
            background-color: #ccc !important;
        }
    </style>
</head>
<body>
<form id="form1" runat="server">
    <!-- Search/Filter Panel (query inputs remain unchanged) -->
    <div class="search-panel">
        Year: <asp:TextBox ID="txtYear" runat="server" Width="60" />
        <asp:Button ID="btnQuery" runat="server" Text="查詢" OnClick="btnQuery_Click" />
        <asp:Button ID="btnExport" runat="server" Text="匯出Excel" OnClick="btnExport_Click" />
        <asp:CheckBox ID="chkThk" runat="server" Text="Show Thk" Checked="true"
                      OnClientClick="toggleColumns('thk', this.checked); return true;" />
        <asp:CheckBox ID="chkRes" runat="server" Text="Show Res" Checked="true"
                      OnClientClick="toggleColumns('res', this.checked); return true;" />
    </div>

    <!-- UpdatePanel for AJAX refresh on query (preserves existing AJAX behavior) -->
    <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
    <asp:UpdatePanel ID="updPanel" runat="server">
        <ContentTemplate>
            <table id="reportTable" runat="server" border="1" cellspacing="0" cellpadding="2">
                <thead>
                    <!-- First header row: Year label spans all dynamic columns; blank spans fixed cols -->
                    <tr>
                        <th colspan="13"></th>
                        <th colspan="17"><asp:Literal ID="litYear" runat="server"></asp:Literal></th>
                    </tr>
                    <!-- Second header row: Quarter labels span their 3 months + quarter total; blank for fixed; year total spans two rows -->
                    <tr>
                        <th colspan="13"></th>
                        <th colspan="4">Q1</th>
                        <th colspan="4">Q2</th>
                        <th colspan="4">Q3</th>
                        <th colspan="4">Q4</th>
                        <th rowspan="2"><asp:Literal ID="litYearTotal" runat="server"></asp:Literal></th>
                    </tr>
                    <!-- Third header row: Fixed column headers + month headers + quarter total labels -->
                    <tr>
                        <th>Customer</th>
                        <th>Inch</th>
                        <th>Cust Item</th>
                        <th>OA Item</th>
                        <th>Substrate</th>
                        <th>Price</th>
                        <th class="thk-col">Thk1</th>
                        <th class="res-col">Res1</th>
                        <th class="thk-col">Thk2</th>
                        <th class="res-col">Res2</th>
                        <th class="thk-col">Thk3</th>
                        <th class="res-col">Res3</th>
                        <th>排程列表</th>
                        <th><asp:Literal ID="litM01" runat="server"></asp:Literal></th>
                        <th><asp:Literal ID="litM02" runat="server"></asp:Literal></th>
                        <th><asp:Literal ID="litM03" runat="server"></asp:Literal></th>
                        <th>合計</th>
                        <th><asp:Literal ID="litM04" runat="server"></asp:Literal></th>
                        <th><asp:Literal ID="litM05" runat="server"></asp:Literal></th>
                        <th><asp:Literal ID="litM06" runat="server"></asp:Literal></th>
                        <th>合計</th>
                        <th><asp:Literal ID="litM07" runat="server"></asp:Literal></th>
                        <th><asp:Literal ID="litM08" runat="server"></asp:Literal></th>
                        <th><asp:Literal ID="litM09" runat="server"></asp:Literal></th>
                        <th>合計</th>
                        <th><asp:Literal ID="litM10" runat="server"></asp:Literal></th>
                        <th><asp:Literal ID="litM11" runat="server"></asp:Literal></th>
                        <th><asp:Literal ID="litM12" runat="server"></asp:Literal></th>
                        <th>合計</th>
                        <!-- Year total header cell is merged above (no separate th here) -->
                    </tr>
                </thead>
                <tbody>
                    <asp:Repeater ID="rptData" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td><%# Eval("Customer") %></td>
                                <td><%# Eval("Inch") %></td>
                                <td><%# Eval("CustItem") %></td>
                                <td><%# Eval("OAItem") %></td>
                                <td><%# Eval("Substrate") %></td>
                                <td><%# Eval("Price", "{0:N2}") %></td>
                                <td class="thk-col"><%# Eval("Thk1") %></td>
                                <td class="res-col"><%# Eval("Res1") %></td>
                                <td class="thk-col"><%# Eval("Thk2") %></td>
                                <td class="res-col"><%# Eval("Res2") %></td>
                                <td class="thk-col"><%# Eval("Thk3") %></td>
                                <td class="res-col"><%# Eval("Res3") %></td>
                                <td><%# Eval("Schedule") %></td>
                                <td><%# Eval("M01") %></td>
                                <td><%# Eval("M02") %></td>
                                <td><%# Eval("M03") %></td>
                                <td><%# Eval("Q1Sum") %></td>
                                <td><%# Eval("M04") %></td>
                                <td><%# Eval("M05") %></td>
                                <td><%# Eval("M06") %></td>
                                <td><%# Eval("Q2Sum") %></td>
                                <td><%# Eval("M07") %></td>
                                <td><%# Eval("M08") %></td>
                                <td><%# Eval("M09") %></td>
                                <td><%# Eval("Q3Sum") %></td>
                                <td><%# Eval("M10") %></td>
                                <td><%# Eval("M11") %></td>
                                <td><%# Eval("M12") %></td>
                                <td><%# Eval("Q4Sum") %></td>
                                <td><%# Eval("YearTotal") %></td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </tbody>
            </table>
        </ContentTemplate>
        <!-- Define triggers: Query button causes partial postback, Export button does full postback -->
        <Triggers>
            <asp:AsyncPostBackTrigger ControlID="btnQuery" EventName="Click" />
            <asp:PostBackTrigger ControlID="btnExport" />
        </Triggers>
    </asp:UpdatePanel>

    <!-- Existing toggle script for Thk/Res columns (unchanged) -->
    <script type="text/javascript">
        function toggleColumns(type, show) {
            var elements = document.querySelectorAll('.' + type + '-col');
            for (var i = 0; i < elements.length; i++) {
                elements[i].style.display = show ? '' : 'none';
            }
            return false;
        }
    </script>
</form>
</body>
</html>
