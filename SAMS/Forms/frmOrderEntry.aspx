<%@ Page Language="C#" MasterPageFile="~/Forms/PageMaster.master" AutoEventWireup="true"
    MaintainScrollPositionOnPostback="true" CodeFile="frmOrderEntry.aspx.cs" Inherits="Forms_frmOrderEntry"
    Title="SAMS :: Order/Invoice Step 2" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" runat="server" ContentPlaceHolderID="cphPage">
    <script type="text/javascript" src="../AjaxLibrary/jquery.searchabledropdown-1.0.8.min.js"></script>

    <script language="JavaScript" type="text/javascript">
        function pageLoad() {
            $("select").searchable();

            $('#<%=ddlSKuCde.ClientID %>').change(function () {
                document.getElementById('<%=txtStock.ClientID%>').value = '0-0';
                var selectedVal = $('#<%=ddlSKuCde.ClientID %> option:selected').text();
                var selectedArry = selectedVal.split(':');
                document.getElementById('<%=txtStock.ClientID%>').value = selectedArry[3];
                document.getElementById('<%=txtUnitRate.ClientID%>').value = selectedArry[2];
                document.getElementById('<%=ddlSKuCde.ClientID%>').focus();
            });

        }

        function ValidateForm() {
            var str;
            str = document.getElementById('<%=txtQuantity.ClientID%>').value;
            var str2 = document.getElementById('<%=txtCtn.ClientID%>').value;
            var str3 = document.getElementById('<%=txtFreeSku.ClientID%>').value;
            if ((str == null || str.length == 0) && (str2 == null || str2.length == 0) && (str3 == null || str3.length == 0)) {
                alert('Must Enter Quantity');
                return false;
            }

            return true;
        }
    </script>
    <script type="text/javascript" language="javascript">
        Sys.WebForms.PageRequestManager.getInstance().add_beginRequest(BeginRequestHandler);
        function BeginRequestHandler(sender, args) {
            var oControl = args.get_postBackElement();
            var ddlName = oControl.name;
            if (ddlName == "ctl00$ctl00$mainCopy$cphPage$ddlPrincipal") {
                document.getElementById('<%=ddlSKuCde.ClientID%>').disabled = true;
            }
            else if (oControl.value == "Save Order" || oControl.value == "Update Order"
                    || oControl.value == "Save Invoice" || oControl.value == "Sale Return"
                    || oControl.value == "Update Invoice"
                    || oControl.value == "Calculate"
                    || oControl.value == "Update Return") {
                oControl.value = "Wait...";
                oControl.disabled = true;
            }
        }
    </script>

    <div id="right_data">
        <table width="100%">
            <tr>
                <td style="width: 45%">
                    <div>
                        <span class="heading">Order/Invoice Step 2</span>
                    </div>
                </td>
                <td style="width: 40%" align="left">
                    <div>
                        <span class="heading">Working Date:
                            <asp:Label ID="lblOrderDate" runat="server"></asp:Label>
                        </span>
                    </div>
                </td>

            </tr>
        </table>
        <div>
            <table width="100%">
                <tr>
                    <td align="left">
                        <div style="left: 120px; position: absolute; top: 190px;">
                        </div>
                        <asp:UpdatePanel ID="UpdatePanel4" runat="server">
                            <ContentTemplate>
                                <table style="width: 100%">
                                    <tbody>
                                        <tr>
                                            <td colspan="5">
                                                <strong>
                                                    <asp:Label ID="lblBillBook" runat="server" Text="Bill Book No:" Visible="false"></asp:Label></strong>
                                                <asp:TextBox ID="txtBillBookNo" runat="server" CssClass="uppercase" MaxLength="10"
                                                    Visible="false" Width="22px"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td valign="top" align="left" style="width: 90px;">
                                                <strong>Invoice #</strong>
                                            </td>
                                            <td valign="top" align="left" colspan="2" style="width: 30px">
                                                <asp:DropDownList ID="drpDocumentNo" runat="server" Width="232px" TabIndex="0"
                                                    OnSelectedIndexChanged="drpDocumentNo_SelectedIndexChanged" AutoPostBack="True">
                                                </asp:DropDownList>

                                            </td>
                                            <td style="width: 26px">
                                                <asp:CheckBox ID="ChbDiscount" runat="server" Width="90px" Text="Promotion" AutoPostBack="True"
                                                    Checked="false" Visible="false" Style="margin-left: 0px"></asp:CheckBox>
                                            </td>
                                            <td style="width: auto;"></td>
                                        </tr>
                                        <tr>
                                            <td valign="middle" align="left" style="width: 90px;">
                                                <strong>Invoice Type</strong>
                                            </td>
                                            <td valign="top" align="left" colspan="2" style="width: 30px">
                                                <asp:RadioButtonList ID="RblPayMode" runat="server" Width="228px" Height="1px" RepeatDirection="Horizontal">
                                                    <asp:ListItem Selected="True" Value="214">Cash</asp:ListItem>
                                                    <asp:ListItem Value="215">Credit</asp:ListItem>
                                                    <asp:ListItem Value="216">Advance</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                            <td valign="top" align="left" style="width: 26px">
                                                <asp:CheckBox ID="ChbBatchNo" runat="server" Text="Batch No" AutoPostBack="True"
                                                    Visible="false"></asp:CheckBox>
                                            </td>
                                            <td rowspan="4" style="width: 800px" align="left">
                                                <asp:Panel ID="promotionData" BorderWidth="1px" Width="220px" Height="100px" runat="server">
                                                    <table>
                                                        <tr>
                                                            <td colspan="3">
                                                                <h2>Last Invoice Discount</h2>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td style="width: 68px">
                                                                <strong>Invoice No:</strong> </td>
                                                            <td style="width: 50px">
                                                                <strong>
                                                                    <asp:Label ID="lblinvoice" runat="server" Text="0"></asp:Label>
                                                                </strong>
                                                            </td>
                                                            <td style="width: 78px">
                                                                <strong>
                                                                    <asp:Label ID="lblInvoiceDate" runat="server" Text='<%# Eval("Date", "{dd MMM yyyy}") %>'></asp:Label>
                                                                </strong>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td style="width: 68px">
                                                                <strong>Sale Qty:</strong> </td>
                                                            <td style="width: 50px">
                                                                <strong>
                                                                    <asp:Label ID="lblQty" runat="server" Text="0.00"></asp:Label>
                                                                </strong>
                                                            </td>
                                                            <td style="width: 78px">
                                                                <strong>
                                                                    <asp:Label ID="lblQtyTime" runat="server" Text='<%# Eval("Date", "{dd MMM yyyy}") %>'></asp:Label>
                                                                </strong>
                                                            </td>
                                                        </tr>


                                                        <tr>
                                                            <td style="width: 68px">
                                                                <strong>Discount:</strong> </td>
                                                            <td style="width: 50px">
                                                                <strong>
                                                                    <asp:Label ID="lblDiscountValue" runat="server" Text="0.00"></asp:Label>
                                                                </strong>
                                                            </td>
                                                            <td style="width: 78px">
                                                                <strong>
                                                                    <asp:Label ID="lblDiscountTime" runat="server" Text='<%# Eval("Date", "{dd MMM yyyy}") %>'></asp:Label>
                                                                </strong>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td style="width: 28px"><strong>Free Sku:</strong></td>
                                                            <td><strong>
                                                                <asp:Label ID="lblfreeskuValue" runat="server" Text="0"></asp:Label></strong></td>
                                                            <td><strong>
                                                                <asp:Label ID="lblfreeskutime" runat="server" Text='<%# Eval("Date", "{dd MMM yyyy}") %>'></asp:Label></strong></td>
                                                        </tr>
                                                    </table>
                                                </asp:Panel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td valign="middle" align="left" style="width: 90px;">
                                                <strong>Area/Saleman</strong>
                                            </td>
                                            <td valign="top" align="left" colspan="2" style="width: 30px">
                                                <asp:TextBox ID="txtprincipal" runat="server" Width="228px" CssClass="txtBox " ReadOnly="True"></asp:TextBox>
                                            </td>
                                            <td valign="top" align="left" style="width: 26px">
                                                <asp:TextBox ID="txtDeliveryMan" runat="server" Width="180px" CssClass="txtBox "
                                                    ReadOnly="True"></asp:TextBox>
                                            </td>

                                        </tr>
                                        <tr>
                                            <td style="height: 23px; width: 90px;" align="left">
                                                <strong>Customer</strong>
                                            </td>
                                            <td style="height: 23px; width: 30px;" align="left" colspan="2">
                                                <asp:DropDownList ID="drpCustomer" runat="server" AutoPostBack="true" TabIndex="1"
                                                    Width="232px" OnSelectedIndexChanged="drpCustomer_SelectedIndexChanged">
                                                </asp:DropDownList>
                                            </td>
                                            <td style="height: 23px; width: 26px;" align="left">
                                                <asp:TextBox ID="txtDiscountType" runat="server" Text="Manual Discount" Width="180px" CssClass="txtBox "
                                                    ReadOnly="True"></asp:TextBox>
                                        </tr>
                                        <tr>
                                            <td valign="middle" align="left" style="width: 90px;">
                                                <strong>
                                                    <label>
                                                        Remarks</label></strong>
                                            </td>
                                            <td valign="top" align="left" colspan="3">
                                                <asp:TextBox ID="txtRemarks" TextMode="MultiLine" TabIndex="2" runat="server" Width="225px"></asp:TextBox>
                                            </td>

                                        </tr>

                                        <tr>

                                            <td>
                                                <strong>Principal</strong>
                                            </td>
                                            <td colspan="2">
                                                <asp:DropDownList ID="ddlPrincipal" TabIndex="3" runat="server" AutoPostBack="true" Width="232px"
                                                    OnSelectedIndexChanged="ddlPrincipal_SelectedIndexChanged"
                                                    AccessKey="P">
                                                </asp:DropDownList>
                                            </td>
                                            <td colspan="" style="width: 26px">
                                                <strong>Available Stock:</strong>
                                                <asp:TextBox ID="txtStock" runat="server" Text="0-0" Enabled="false" Width="50px"></asp:TextBox>
                                            </td>
                                            <td>
                                                <asp:RadioButtonList ID="rbExtraDiscountType" runat="server" Width="94px" Height="1px"
                                                    RepeatDirection="Horizontal" AutoPostBack="true" OnSelectedIndexChanged="rbExtraDiscountType_SelectedIndexChanged">
                                                    <asp:ListItem Value="1">%</asp:ListItem>
                                                    <asp:ListItem Value="2" Selected="True">Value</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>

                                        </tr>

                                    </tbody>
                                </table>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                    <td></td>
                </tr>
                <tr>
                    <td align="left">
                        <asp:Panel ID="Panel5" runat="server" DefaultButton="btnSave">
                            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                <ContentTemplate>
                                    <table width="100%">
                                        <tr>
                                            <td style="width: 270px" class="lblDetail">
                                                <strong style="color: White;">SKU Description</strong>
                                            </td>
                                            <td style="width: 65px" class="lblDetail">
                                                <strong style="color: White;">Ctn</strong>
                                            </td>
                                            <td style="width: 65px" class="lblDetail">
                                                <strong style="color: White;">Unit</strong>
                                            </td>
                                            <td style="width: 65px" class="lblDetail">
                                                <strong style="color: White;">Free Unit</strong>

                                            </td>
                                            <td style="width: 65px" class="lblDetail">
                                                <strong style="color: White;">Unit Price</strong>
                                            </td>
                                            <td style="width: 65px" class="lblDetail">
                                                <strong style="color: White;">Amount</strong>
                                            </td>
                                            <td style="width: 65px" class="lblDetail">
                                                <strong style="color: White;">Disc.(%)</strong>
                                            </td>
                                            <td style="width: 65px" class="lblDetail">
                                                <strong style="color: White;">Disc.Value</strong>
                                            </td>
                                            <td style="width: 70px" class="lblDetail">
                                                <strong style="color: White;">Net Value</strong>
                                            </td>
                                            <td colspan="2"></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DropDownList ID="ddlSKuCde"
                                                    runat="server" Width="100%" AutoPostBack="true" TabIndex="4"
                                                    OnSelectedIndexChanged="ddlSKuCde_SelectedIndexChanged1">
                                                </asp:DropDownList>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtCtn" runat="server" TabIndex="5" Width="100%"></asp:TextBox>
                                                <ajaxToolkit:FilteredTextBoxExtender ID="FilteredTextBoxExtender5" runat="server"
                                                    FilterType="Custom" ValidChars="01234567890" TargetControlID="txtCtn">
                                                </ajaxToolkit:FilteredTextBoxExtender>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtQuantity" runat="server" TabIndex="6" Width="100%"></asp:TextBox>
                                                <ajaxToolkit:FilteredTextBoxExtender ID="FilteredTextBoxExtender6" runat="server"
                                                    FilterType="Custom" ValidChars="01234567890" TargetControlID="txtQuantity">
                                                </ajaxToolkit:FilteredTextBoxExtender>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtFreeSku" runat="server" TabIndex="7" Width="100%"></asp:TextBox>
                                                <ajaxToolkit:FilteredTextBoxExtender ID="FilteredTextBoxExtender7" runat="server"
                                                    FilterType="Custom" ValidChars="01234567890" TargetControlID="txtFreeSku">
                                                </ajaxToolkit:FilteredTextBoxExtender>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtUnitRate" runat="server" Width="100%" Enabled="False"></asp:TextBox>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtAmount" runat="server" Enabled="False" Width="100%"></asp:TextBox>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtExtDiscount" TabIndex="8" runat="server" Width="100%"></asp:TextBox>
                                                <ajaxToolkit:FilteredTextBoxExtender ID="FilteredTextBoxExtender3" runat="server"
                                                    FilterType="Custom" ValidChars="01234567890." TargetControlID="txtExtDiscount">
                                                </ajaxToolkit:FilteredTextBoxExtender>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtExtDiscountValue" TabIndex="9" runat="server" Width="100%" Enabled="false"></asp:TextBox>
                                                <ajaxToolkit:FilteredTextBoxExtender ID="FilteredTextBoxExtender4" runat="server"
                                                    FilterType="Custom" ValidChars="01234567890." TargetControlID="txtExtDiscountValue">
                                                </ajaxToolkit:FilteredTextBoxExtender>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtNetValue" runat="server" Width="100%" Enabled="false"></asp:TextBox>
                                            </td>
                                            <td>
                                                <asp:DropDownList runat="server" ID="drpReturnType" TabIndex="10" Visible="false">
                                                    <asp:ListItem Value="1" Text="Expiry"></asp:ListItem>
                                                    <asp:ListItem Value="2" Text="Damage"></asp:ListItem>
                                                    <asp:ListItem Selected="True" Value="3" Text="Saleable"></asp:ListItem>
                                                </asp:DropDownList>
                                            </td>
                                            <td colspan="2">
                                                <asp:UpdatePanel ID="upAdd" runat="server">
                                                    <ContentTemplate>
                                                        <asp:Button ID="btnSave" runat="server" TabIndex="11" Font-Size="8pt" OnClick="btnSave_Click" Text="Add"
                                                            ValidationGroup="vg" Width="100%" AccessKey="A" CssClass="Button" />
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="11">
                                                <asp:Panel ID="Panel2" runat="server" Height="130px" ScrollBars="Vertical" Width="100%"
                                                    BorderColor="Silver" BorderStyle="Groove" BorderWidth="1px">
                                                    <asp:GridView ID="GrdPurchase" runat="server" ForeColor="SteelBlue" BackColor="White"
                                                        HorizontalAlign="Center" AutoGenerateColumns="False" BorderColor="White" ShowHeader="False"
                                                        OnRowDeleting="GrdPurchase_RowDeleting" OnRowEditing="GrdPurchase_RowEditing"
                                                        Width="100%">
                                                        <Columns>
                                                            <asp:BoundField DataField="SKU_ID" HeaderText="SKU_ID">
                                                                <HeaderStyle CssClass="HidePanel"></HeaderStyle>
                                                                <ItemStyle CssClass="HidePanel"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SKU_CODE" HeaderText="SKU Code">
                                                                <ItemStyle HorizontalAlign="Left" CssClass="grdDetail" Width="70px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SKU_NAME" HeaderText="SKU Name">
                                                                <ItemStyle HorizontalAlign="Left" CssClass="grdDetail" Width="195px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="QUANTITY_CTN" HeaderText="Ctn">
                                                                <ItemStyle HorizontalAlign="Right" CssClass="grdDetail" Width="65px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="QUANTITY_UNIT" HeaderText="Quantity">
                                                                <ItemStyle HorizontalAlign="Right" CssClass="grdDetail" Width="65px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="FREE_SKU" HeaderText="Free_unit">
                                                                <ItemStyle HorizontalAlign="Right" CssClass="grdDetail" Width="65px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="UNIT_PRICE" HeaderText="PRICE" DataFormatString="{0:F2}">
                                                                <ItemStyle HorizontalAlign="Right" CssClass="grdDetail" Width="65px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Amount" HeaderText="Amount" DataFormatString="{0:F2}">
                                                                <ItemStyle HorizontalAlign="Right" CssClass="grdDetail" Width="65px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="EXTRA_DISCOUNT_PER" HeaderText="Ext. Discount" DataFormatString="{0:F2}">
                                                                <ItemStyle HorizontalAlign="Right" CssClass="grdDetail" Width="65px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="EXTRA_DISCOUNT" HeaderText="Ext. Dis Value" DataFormatString="{0:F2}">
                                                                <ItemStyle HorizontalAlign="Right" CssClass="grdDetail" Width="65px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="NET_AMOUNT" HeaderText="Net Value" DataFormatString="{0:F2}">
                                                                <ItemStyle HorizontalAlign="Right" CssClass="grdDetail" Width="75px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="QUANTITY_CTN2" HeaderText="QUANTITY_CTN2">
                                                                <HeaderStyle CssClass="HidePanel" />
                                                                <ItemStyle CssClass="HidePanel" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="QUANTITY_UNIT2" HeaderText="QUANTITY_UNIT2">
                                                                <HeaderStyle CssClass="HidePanel" />
                                                                <ItemStyle CssClass="HidePanel" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Stock" HeaderText="Stock">
                                                                <HeaderStyle CssClass="HidePanel" />
                                                                <ItemStyle CssClass="HidePanel" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="UNITS_IN_CASE" HeaderText="UNITS_IN_CASE">
                                                                <HeaderStyle CssClass="HidePanel" />
                                                                <ItemStyle CssClass="HidePanel" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="returnType" HeaderText="returnType">
                                                                <HeaderStyle CssClass="HidePanel" />
                                                                <ItemStyle CssClass="HidePanel" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PRINCIPAL_ID" HeaderText="PRINCIPAL_ID">
                                                                <HeaderStyle CssClass="HidePanel" />
                                                                <ItemStyle CssClass="HidePanel" />
                                                            </asp:BoundField>
                                                            <asp:CommandField ShowEditButton="True" HeaderText="Edit">
                                                                <ItemStyle HorizontalAlign="Center" CssClass="grdDetail" Width="45px"></ItemStyle>
                                                            </asp:CommandField>
                                                            <asp:TemplateField HeaderText="Delete">
                                                                <ItemTemplate>
                                                                    <asp:LinkButton ID="btnDelete" runat="server" Text="Delete" OnClientClick="javascript:return confirm('Are you sure you want to Delete?');return false;"
                                                                        CommandName="Delete"></asp:LinkButton>
                                                                </ItemTemplate>
                                                                <ItemStyle HorizontalAlign="Center" Width="50px" CssClass="grdDetail"></ItemStyle>
                                                            </asp:TemplateField>
                                                        </Columns>
                                                        <AlternatingRowStyle BackColor="#F2F2F2" CssClass="GridAlternateRowStyle" ForeColor="#333333" />
                                                    </asp:GridView>
                                                </asp:Panel>

                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="right">Total</td>
                                            <td style="border: 1px solid silver;" align="right">
                                                <asp:Label ID="lblCtn" runat="server" Width="100%" Text="0"></asp:Label>
                                            </td>
                                            <td style="border: 1px solid silver;" align="right">
                                                <asp:Label ID="lblUnit" runat="server" Width="100%" Text="0"></asp:Label>
                                            </td>
                                            <td style="border: 1px solid silver;" align="right">
                                                <asp:Label ID="lblFreeUnit" runat="server" Width="100%" Text="0"></asp:Label>
                                            </td>
                                            <td style="border: 1px solid silver;"></td>
                                            <td style="border: 1px solid silver;" align="right">
                                                <asp:Label ID="lblAmount" runat="server" Width="100%" Text="0"></asp:Label>
                                            </td>
                                            <td></td>
                                            <td style="border: 1px solid silver;" align="right">
                                                <asp:Label ID="lblExtraDiscount" runat="server" Width="100%" Text="0"></asp:Label>
                                            </td>
                                            <td style="border: 1px solid silver;" align="right">
                                                <asp:Label ID="lblNetValue" runat="server" Width="100%" Text="0"></asp:Label>
                                            </td>
                                            <td colspan="2"></td>
                                        </tr>
                                    </table>
                                    <hr />
                                </ContentTemplate>
                            </asp:UpdatePanel>
                        </asp:Panel>
                    </td>
                    <td>
                        <asp:HiddenField ID="hfBillBookNo" runat="server" />
                    </td>
                </tr>
            </table>
        </div>
        <div>
            <table width="100%">
                <tr>
                    <td align="left">
                        <asp:UpdatePanel ID="UpdatePanel3" runat="server">
                            <ContentTemplate>
                                <table>
                                    <tbody>
                                        <tr>
                                            <td valign="top" align="left" colspan="2" rowspan="7">
                                                <strong></strong>
                                                <%--<asp:Panel ID="Panel4" runat="server" Width="350px" Height="130px" BorderColor="Silver"
                                                    BorderStyle="Groove" ScrollBars="Vertical" BorderWidth="1px">
                                                    <asp:GridView ID="GrdFreeSKU" runat="server" Width="100%" ForeColor="Silver" CssClass="gridRow2"
                                                        BorderColor="White" BackColor="White" AutoGenerateColumns="False" HorizontalAlign="Center">

                                                        <RowStyle BorderColor="Silver" BorderWidth="1px" BorderStyle="Solid" ForeColor="Black" />

                                                        <Columns>
                                                            <asp:TemplateField>
                                                                <ItemTemplate>
                                                                    <asp:CheckBox ID="cbSelect" runat="server" Checked="true" />
                                                                </ItemTemplate>
                                                                <ItemStyle BorderColor="Silver" BorderStyle="Solid" BorderWidth="1px" HorizontalAlign="Center" />
                                                            </asp:TemplateField>
                                                            <asp:BoundField DataField="SKU_ID" HeaderText="SKU_ID">
                                                                <HeaderStyle CssClass="HidePanel"></HeaderStyle>
                                                                <ItemStyle CssClass="HidePanel"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SKU_Code" HeaderText="SKU Code">
                                                                <ItemStyle Width="80px" BorderColor="Silver" BorderStyle="Solid" BorderWidth="1px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="SKU_Name" HeaderText="SKU Name">
                                                                <ItemStyle Width="200px" BorderColor="Silver" BorderStyle="Solid" BorderWidth="1px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Quantity" HeaderText="Qty">
                                                                <ItemStyle Width="50px" BorderColor="Silver" BorderStyle="Solid" BorderWidth="1px"></ItemStyle>
                                                            </asp:BoundField>
                                                        </Columns>
                                                        <HeaderStyle CssClass="tblhead"></HeaderStyle>
                                                    </asp:GridView>
                                                </asp:Panel>--%>
                                            </td>
                                            <td style="height: 20px; width: 100px;" align="left"></td>
                                            <td style="height: 20px; width: 20%" align="left">

                                                <strong>Gross Sale</strong>
                                            </td>
                                            <td style="width: 7px; height: 20px">
                                                <asp:TextBox ID="txtGrossAmount" runat="server" Width="150px" ForeColor="Black" Font-Bold="True"
                                                    CssClass="txtBoxnNum" ReadOnly="True"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left" style="width: 100px"></td>
                                            <td align="left">
                                                <strong>Discount</strong>
                                            </td>
                                            <td style="width: 7px">
                                                <asp:TextBox ID="numtxtTotalExtraDiscnt" runat="server" Width="150px" ForeColor="Black"
                                                    ReadOnly="true" Font-Bold="True"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="height: 18px; width: 100px;" align="left"></td>
                                            <td style="height: 18px" align="left">
                                                <strong>Extra Discount</strong>
                                            </td>
                                            <td style="width: 7px; height: 18px">
                                                <asp:TextBox ID="numTxtTotalStndrdDiscnt" runat="server" Width="150px" ForeColor="Black"
                                                    Font-Bold="True" CssClass="txtBoxnNum" ReadOnly="false"></asp:TextBox>
                                                <ajaxToolkit:FilteredTextBoxExtender ID="FilteredTextBoxExtender" runat="server"
                                                    FilterType="Custom" ValidChars="01234567890." TargetControlID="numTxtTotalStndrdDiscnt">
                                                </ajaxToolkit:FilteredTextBoxExtender>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="width: 100px; height: 20px" valign="top"></td>
                                            <td style="width: 1px; height: 20px" align="left">
                                                <strong>Claimable Discount</strong>

                                            </td>
                                            <td style="width: 7px; height: 20px" align="right">
                                                <asp:TextBox ID="numtxtUnClaimabledist" runat="server" Width="150px" ForeColor="Black"
                                                    Font-Bold="True" CssClass="txtBoxnNum" ReadOnly="True"></asp:TextBox>
                                                <ajaxToolkit:FilteredTextBoxExtender ID="FilteredTextBoxExtender1" runat="server"
                                                    FilterType="Custom" ValidChars="01234567890." TargetControlID="numtxtUnClaimabledist">
                                                </ajaxToolkit:FilteredTextBoxExtender>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="width: 100px; height: 20px" valign="top"></td>
                                            <td style="width: 1px; height: 20px" align="left">
                                                <strong>GST Amount</strong>

                                            </td>
                                            <td style="width: 7px; height: 20px" align="right">
                                                <asp:TextBox ID="numTxtTotalGST" runat="server" Width="150px" ForeColor="Black" Font-Bold="True"
                                                    CssClass="txtBoxnNum" ReadOnly="True"></asp:TextBox>
                                                <ajaxToolkit:FilteredTextBoxExtender ID="FilteredTextBoxExtender2" runat="server"
                                                    FilterType="Custom" ValidChars="01234567890." TargetControlID="numTxtTotalGST">
                                                </ajaxToolkit:FilteredTextBoxExtender>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="width: 100px; height: 20px" valign="top"></td>
                                            <td style="width: 1px; height: 20px" align="left">
                                                <strong>TST Amount</strong>

                                            </td>
                                            <td style="width: 7px; height: 20px" align="right">
                                                <asp:TextBox ID="numTxtTotalTST" runat="server" Width="150px" ForeColor="Black" Font-Bold="True"
                                                    CssClass="txtBoxnNum" ReadOnly="True"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="width: 100px; height: 20px" valign="top"></td>
                                            <td style="width: 1px; height: 20px" align="left">
                                                <strong>Net Amount</strong>
                                            </td>
                                            <td style="width: 7px; height: 20px" align="right">
                                                <asp:TextBox ID="numTxtTotlAmnt" runat="server" Width="150px" ForeColor="Black" Font-Bold="True"
                                                    CssClass="txtBoxnNum" ReadOnly="True"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left" colspan="2">
                                                <table>
                                                    <tbody>
                                                        <tr>
                                                            <td style="width: 100px">
                                                                <asp:Button AccessKey="C" ID="btnCalculate" OnClick="btnCalculate_Click" runat="server"
                                                                    Width="100px" Font-Size="8pt" Text="Calculate" Enabled="False" CssClass="Button" />
                                                            </td>
                                                            <td style="width: 100px"  runat="server" id="tdInvButton" visible="false">
                                                                <asp:Button AccessKey="I" ID="btnSaveInvoice" OnClick="btnSaveInvoice_Click" runat="server"
                                                                    Width="100px" Font-Size="8pt" Text="Save Invocie" Enabled="False" CssClass="Button" />
                                                            </td>
                                                            <td style="width: 100px">
                                                                <asp:Button AccessKey="S" ID="btnSaveOrder" OnClick="btnSaveOrder_Click" runat="server"
                                                                    Width="110px" Font-Size="8pt" Text="Save Order" CssClass="Button" />
                                                            </td>
                                                            <td style="width: 100px">
                                                                <asp:Button AccessKey="H" ID="btnCancel" runat="server" Width="110px" Font-Size="8pt"
                                                                    Text="Home" OnClick="btnCancel_Click" CssClass="Button" />
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                            <td style="width: 100px" valign="top" align="left">&nbsp; &nbsp; &nbsp;&nbsp;
                                            </td>
                                            <td style="width: 1px" align="left">
                                                <strong>
                                                    <asp:Label ID="Label5" runat="server" Width="105px" Text="Cash Received" Visible="false"></asp:Label></strong>
                                            </td>
                                            <td style="width: 7px" align="right">
                                                <asp:TextBox ID="txtCashReceived" runat="server" Width="150px" ForeColor="Black"
                                                    Font-Bold="True" CssClass="txtBoxnNum" Visible="False"></asp:TextBox>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                        &nbsp;&nbsp;
                        <asp:TextBox ID="numTxtTotalSED" runat="server" CssClass="txtBox " Font-Bold="False"
                            ForeColor="Black" Width="139px" ReadOnly="True" Visible="False"></asp:TextBox>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>
        </div>
    </div>

</asp:Content>
