using System;
using System.Data;
using System.Collections;
using System.Web.UI;
using System.Web.UI.WebControls;
using SAMSBusinessLayer.Classes;
using SAMSCommon.Classes;

/// <summary>
/// Form To Take Order, Invoice And Sale Return(Step2)
/// </summary>
public partial class Forms_frmOrderEntry : System.Web.UI.Page
{
    #region Variables

    DataTable PurchaseSKU;
    DataTable dtFreeSKU;
    private static int mCustomerTypeId;
    private static int mCustomerVolClassId;
    private int mTownId;
    private static int RowId;
    private static int OrderNo;

    #endregion

    DataControl dc = new DataControl();

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            int OrderNo = (int)Session["OrderNo"];
            txtprincipal.Text = Session["Route"].ToString();
            txtDeliveryMan.Text = Session["SaleMan"].ToString();
            if (OrderNo == -1)
            {
                btnSaveOrder.Text = "Save Invoice";
                drpDocumentNo.Enabled = false;
            }
            else if (OrderNo == -2)
            {
                drpDocumentNo.Enabled = false;
                ChbDiscount.Visible = true;
                btnSaveOrder.Text = "Sale Return";
                txtUnitRate.Enabled = true;
                lblBillBook.Visible = false;
                txtBillBookNo.Visible = false;

                drpReturnType.Visible = true;
                //txtFreeSku.Enabled = false;
            }
            LoadPrincipal();
            LoadCustomerData();
            LoadSKUDetail();
            CreateTable();
            //LoadPromotion();
            //CreateFreeSKU();
            // LoadFreeGrid();
            LoadPendingOrder();
            LoadSKULastPromotionDiscount();
            btnSave.Attributes.Add("onclick", "return ValidateForm();");

            ddlSKuCde_SelectedIndexChanged(null, null);
            btnSaveOrder.Enabled = false;
            lblOrderDate.Text = Convert.ToDateTime(Session["OrderDate"]).ToString("dd-MMM-yyyy");
            rbExtraDiscountType_SelectedIndexChanged(null, null);

            ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
        }
    }

    private void CreateTable()
    {
        PurchaseSKU = new DataTable();
        PurchaseSKU.Columns.Add("SALE_ORDER_DETAIL_ID", typeof(long));
        PurchaseSKU.Columns.Add("DistributorId", typeof(int));
        PurchaseSKU.Columns.Add("SALE_ORDER_ID", typeof(int));
        PurchaseSKU.Columns.Add("SKU_ID", typeof(int));
        PurchaseSKU.Columns.Add("SKU_Code", typeof(string));
        PurchaseSKU.Columns.Add("SKU_Name", typeof(string));
        PurchaseSKU.Columns.Add("BATCH_NO", typeof(string));
        PurchaseSKU.Columns.Add("UNIT_PRICE", typeof(decimal));
        PurchaseSKU.Columns.Add("QUANTITY_UNIT", typeof(int));
        PurchaseSKU.Columns.Add("QUANTITY_UNIT2", typeof(int));
        PurchaseSKU.Columns.Add("FREE_SKU", typeof(int));
        PurchaseSKU.Columns.Add("AMOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("STANDARD_DISCOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("STANDARD_DISCOUNT_PER", typeof(decimal));
        PurchaseSKU.Columns.Add("EXTRA_DISCOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("EXTRA_DISCOUNT_PER", typeof(decimal));
        PurchaseSKU.Columns.Add("RETAIL_AMOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("GST_RATE", typeof(decimal));
        PurchaseSKU.Columns.Add("GST_AMOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("TST_AMOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("CLAIM_EXTRA_AMOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("CLAIM_STANDARD_DISCOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("CLAIM_PER", typeof(decimal));
        PurchaseSKU.Columns.Add("SED_AMOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("NET_AMOUNT", typeof(decimal));
        PurchaseSKU.Columns.Add("QUANTITY_CTN", typeof(decimal));
        PurchaseSKU.Columns.Add("QUANTITY_CTN2", typeof(decimal));
        PurchaseSKU.Columns.Add("IS_DELETED", typeof(bool));
        PurchaseSKU.Columns.Add("UNITS_IN_CASE", typeof(int));
        PurchaseSKU.Columns.Add("Stock", typeof(string));
        PurchaseSKU.Columns.Add("QUANTITY", typeof(decimal));
        PurchaseSKU.Columns.Add("returnType", typeof(int));
        PurchaseSKU.Columns.Add("PRINCIPAL_ID", typeof(int));
        Session.Add("PurchaseSKU", PurchaseSKU);
    }
    private void CreateFreeSKU()
    {
        dtFreeSKU = new DataTable();
        dtFreeSKU.Columns.Add("SKU_ID", typeof(int));
        dtFreeSKU.Columns.Add("SKU_Code", typeof(string));
        dtFreeSKU.Columns.Add("SKU_Name", typeof(string));
        dtFreeSKU.Columns.Add("UNIT_PRICE", typeof(decimal));
        dtFreeSKU.Columns.Add("Quantity", typeof(int));
        dtFreeSKU.Columns.Add("AMOUNT", typeof(decimal));
        dtFreeSKU.Columns.Add("GST_RATE", typeof(decimal));
        dtFreeSKU.Columns.Add("GST_AMOUNT", typeof(decimal));
        dtFreeSKU.Columns.Add("TST_AMOUNT", typeof(decimal));
        dtFreeSKU.Columns.Add("PROMOTION_ID", typeof(int));
        dtFreeSKU.Columns.Add("BASKET_ID", typeof(int));
        dtFreeSKU.Columns.Add("BASKET_DETAIL_ID", typeof(int));
        dtFreeSKU.Columns.Add("PROMOTION_OFFER_ID", typeof(int));
        Session.Add("dtFreeSKU", dtFreeSKU);
    }

    #region Load

    /// <summary>
    /// Loads Pending Order Nos To Docuemnt No Combo
    /// </summary>
    private void LoadPendingOrder()
    {
        OrderEntryController or = new OrderEntryController();
        drpDocumentNo.Items.Clear();
        int OrderNo = (int)Session["OrderNo"];
        if (OrderNo == -2)
        {
            DataTable dtOrder = or.SelectPendingOrder(int.Parse(Session["DistributorId"].ToString()), int.Parse(Session["AreaId"].ToString()), Constants.IntNullValue, int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), Constants.Order_Pending_Id, Constants.IntNullValue, int.Parse(Session["UserId"].ToString()), Convert.ToDateTime(Session["OrderDate"]), 2);
            drpDocumentNo.Items.Add(new clsListItems("New", Constants.LongNullValue.ToString()));

            clsWebFormUtil.FillDropDownList(drpDocumentNo, dtOrder, "SALE_INVOICE_ID", "SALE_INVOICE_ID2");

            Session.Add("dtOrder", dtOrder);
        }
        else
        {
            DataTable dtOrder = or.SelectPendingOrder(int.Parse(Session["DistributorId"].ToString()), int.Parse(Session["AreaId"].ToString()), Constants.IntNullValue, int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), Constants.Order_Pending_Id, Constants.IntNullValue, int.Parse(Session["UserId"].ToString()), Convert.ToDateTime(Session["OrderDate"]), 0);
            drpDocumentNo.Items.Add(new clsListItems("New", Constants.LongNullValue.ToString()));

            clsWebFormUtil.FillDropDownList(drpDocumentNo, dtOrder, "SALE_INVOICE_ID", "SALE_INVOICE_ID2");

            Session.Add("dtOrder", dtOrder);
        }

    }
    private void LoadCustomerData()
    {
        if (int.Parse(Session["DistributorId"].ToString()) > 0 && long.Parse(Session["AreaId"].ToString()) > 0)
        {
            drpCustomer.Items.Clear();
            CustomerDataController mController = new CustomerDataController();
            DataTable dtCustomer = mController.SelectPrincipalCustomer(int.Parse(Session["DistributorId"].ToString()), int.Parse(Session["AreaId"].ToString()), Constants.IntNullValue, Constants.IntNullValue);
            clsWebFormUtil.FillDropDownList(drpCustomer, dtCustomer, "CUSTOMER_ID", "CUSTOMER_DETAIL2", true);
            Session.Add("dtCustomer", dtCustomer);
        }
        else
        {
            drpCustomer.Items.Clear();
        }
    }
    private void LoadSKUDetail()
    {
        DataTable Dtsku_Price = (DataTable)Session["Dtsku_Price"];

        DataView dv = new DataView(Dtsku_Price);
        dv.RowFilter = "PRINCIPAL_ID = " + ddlPrincipal.SelectedValue;
        clsWebFormUtil.FillDropDownList(ddlSKuCde, dv.ToTable(), "Sku_ID", "SkuPriceDetail2", true);

    }
    private void LoadPromotion()
    {
        SKUPriceDetailController PController = new SKUPriceDetailController();
        DataTable dt = PController.SelectDataPrice(Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, int.Parse(Session["UserId"].ToString()), Constants.IntNullValue, 0, DateTime.Parse(Session["OrderDate"].ToString()));
        // if (bool.Parse(dt.Rows[0]["Is_ManualDiscount"].ToString()) == true)
        // {
        OrderEntryController orderController = new OrderEntryController();
        //    PromotionCollections_Controller Arpc = orderController.LoadSchemes(int.Parse(Session["DistributorId"].ToString()), Constants.IntNullValue, Convert.ToDateTime(Session["OrderDate"]));
        PromotionCollections_Controller Arpc = orderController.LoadSchemesOnlyScheme(int.Parse(Session["DistributorId"].ToString()), Constants.IntNullValue, Convert.ToDateTime(Session["OrderDate"]));

        Session.Add("Arpc", Arpc);
        txtDiscountType.Text = "Auto Scheme";

        //
        numTxtTotalStndrdDiscnt.ReadOnly = false;
        //
        numtxtUnClaimabledist.ReadOnly = true;
        ChbDiscount.Checked = true;
        //}
        //else
        //{
        //    txtDiscountType.Text = "Manual Discount";
        //    numTxtTotalStndrdDiscnt.ReadOnly = false;            
        //    numtxtUnClaimabledist.ReadOnly = false;
        //    ChbDiscount.Checked = false;
        //}
    }
    private void LoadGird()
    {
        PurchaseSKU = (DataTable)Session["PurchaseSKU"];
        GrdPurchase.DataSource = PurchaseSKU;
        GrdPurchase.DataBind();

        CalculateTotal(PurchaseSKU);
    }
    private void LoadFreeGrid()
    {
        //dtFreeSKU = (DataTable)Session["dtFreeSKU"];
        //GrdFreeSKU.DataSource = dtFreeSKU;
        //GrdFreeSKU.DataBind();
    }

    private void LoadPrincipal()
    {
        SKUPriceDetailController PController = new SKUPriceDetailController();
        DataTable m_dt = PController.SelectDataPrice(Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, int.Parse(Session["UserId"].ToString()), Constants.IntNullValue, 0, DateTime.Parse(Session["CurrentWorkDate"].ToString()));
        clsWebFormUtil.FillDropDownList(ddlPrincipal, m_dt, 0, 1);
    }
    #endregion

    private void CalculateTotal(DataTable dt)
    {
        lblCtn.Text = "0.00";
        lblUnit.Text = "0.00";
        lblFreeUnit.Text = "0.00";
        lblAmount.Text = "0.00";
        lblExtraDiscount.Text = "0.00";
        lblNetValue.Text = "0.00";
        decimal ctn = 0, free_sku = 0, unit = 0, amount = 0, extradiscount = 0, net = 0;
        if (dt != null)
        {
            foreach (DataRow dr in dt.Rows)
            {
                ctn += Convert.ToDecimal(dr["QUANTITY_CTN"]);
                unit += Convert.ToDecimal(dr["QUANTITY_UNIT"]);
                free_sku += Convert.ToDecimal(dr["FREE_SKU"]);


                amount += Convert.ToDecimal(dr["Amount"]);
                extradiscount += Convert.ToDecimal(dr["EXTRA_DISCOUNT"]);
                net += Convert.ToDecimal(dr["NET_AMOUNT"]);
            }
        }

        lblCtn.Text = ctn.ToString();
        lblUnit.Text = unit.ToString();
        lblFreeUnit.Text = free_sku.ToString();
        lblAmount.Text = String.Format("{0:0.00}", amount);
        lblExtraDiscount.Text = String.Format("{0:0.00}", extradiscount);
        lblNetValue.Text = String.Format("{0:0.00}", net);
    }

    #region Grid Operations

    protected void GrdPurchase_RowEditing(object sender, GridViewEditEventArgs e)
    {
        try
        {
            int stock = 0;
            stock += Convert.ToInt32(GrdPurchase.Rows[e.NewEditIndex].Cells[11].Text) * Convert.ToInt32(GrdPurchase.Rows[e.NewEditIndex].Cells[14].Text);
            stock += Convert.ToInt32(GrdPurchase.Rows[e.NewEditIndex].Cells[12].Text);

            string[] strstock = GrdPurchase.Rows[e.NewEditIndex].Cells[13].Text.Split('-');
            stock += Convert.ToInt32(strstock[0]) * Convert.ToInt32(GrdPurchase.Rows[e.NewEditIndex].Cells[14].Text);
            stock += Convert.ToInt32(strstock[1]);

            RowId = e.NewEditIndex;
            ddlPrincipal.SelectedValue = GrdPurchase.Rows[e.NewEditIndex].Cells[16].Text;
            this.LoadSKUDetail();
            ddlSKuCde.SelectedValue = GrdPurchase.Rows[e.NewEditIndex].Cells[0].Text;

            txtCtn.Text = GrdPurchase.Rows[e.NewEditIndex].Cells[3].Text;
            txtQuantity.Text = GrdPurchase.Rows[e.NewEditIndex].Cells[4].Text;
            txtFreeSku.Text = GrdPurchase.Rows[e.NewEditIndex].Cells[5].Text;
            txtExtDiscount.Text = GrdPurchase.Rows[e.NewEditIndex].Cells[8].Text;
            txtExtDiscountValue.Text = GrdPurchase.Rows[e.NewEditIndex].Cells[9].Text;
            txtUnitRate.Text = GrdPurchase.Rows[e.NewEditIndex].Cells[6].Text;
            txtStock.Text = (stock / Convert.ToInt32(GrdPurchase.Rows[e.NewEditIndex].Cells[14].Text)).ToString() + "-" + (stock % Convert.ToInt32(GrdPurchase.Rows[e.NewEditIndex].Cells[14].Text)).ToString();

            int OrderNo = (int)Session["OrderNo"];
            if (OrderNo == -2)//for Sale Return
            {
                drpReturnType.SelectedValue = GrdPurchase.Rows[e.NewEditIndex].Cells[15].Text;
            }
            ddlSKuCde.Enabled = false;

            btnSave.Text = "Update";
        }
        catch (Exception ex)
        {
            ddlSKuCde.Enabled = false;
            btnSave.Text = "Update";
            ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + ex.Message.ToString() + "')", true);
        }
    }
    protected void GrdPurchase_RowDeleting(object sender, GridViewDeleteEventArgs e)
    {
        PurchaseSKU = (DataTable)Session["PurchaseSKU"];
        PurchaseSKU.Rows.RemoveAt(e.RowIndex);
        Session.Add("PurchaseSKU", PurchaseSKU);

        LoadGird();
    }

    #endregion

    private bool CheckDublicateSKU()
    {
        PurchaseSKU = (DataTable)Session["PurchaseSKU"];
        DataRow[] foundRows = PurchaseSKU.Select("SKU_ID  = '" + ddlSKuCde.SelectedValue + "' and returnType = '" + drpReturnType.SelectedValue + "'");
        if (foundRows.Length == 0)
        {
            return true;
        }

        return false;

    }

    private int FreeSKUExist(DataTable dt, int Sku_id)
    {
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            if (int.Parse(dt.Rows[i]["SKU_ID"].ToString()) == Sku_id)
            {
                return i;
            }

        }
        return -1;
    }

    private void ExistenOrderDetail(long OrderId, int Type)
    {
        OrderEntryController ord = new OrderEntryController();
        PurchaseSKU = ord.SelectOrderDetail(int.Parse(Session["DistributorId"].ToString()), OrderId, Type);
        //  dtFreeSKU = ord.SelectOrderPromotion(int.Parse(Session["DistributorId"].ToString()), OrderId, Type);

        Session.Add("PurchaseSKU", PurchaseSKU);
        //  Session.Add("dtFreeSKU", dtFreeSKU);

        LoadGird();
        //  LoadFreeGrid();
    }

    private bool IsBillBookNoExist()
    {
        bool flag = false;
        if ((long.Parse(drpDocumentNo.SelectedValue) == Constants.LongNullValue || Session["hfBillBookNo"].ToString() != txtBillBookNo.Text) && txtBillBookNo.Text.Trim().Length > 0)
        {
            OrderEntryController OEC = new OrderEntryController();
            DataTable dtBillBookNo = OEC.SelectBillBookNo(Convert.ToInt32(Session["DistributorId"].ToString()), txtBillBookNo.Text, 0);
            if (dtBillBookNo.Rows.Count > 0)
            {
                flag = true;
            }
            DataTable dtBillBookNo2 = OEC.SelectBillBookNo(Convert.ToInt32(Session["DistributorId"].ToString()), txtBillBookNo.Text, 1);
            if (dtBillBookNo2.Rows.Count > 0)
            {
                flag = true;
            }
        }
        return flag;
    }

    private void EnableDisableController(bool CValue)
    {

        if (CValue == true)
        {
            if (decimal.Parse(txtGrossAmount.Text) > 0)
            {
                drpCustomer.Enabled = true;

            }
        }
        else
        {
            drpCustomer.Enabled = false;

        }
    }

    #region Click Operations

    protected void btnSave_Click(object sender, EventArgs e)
    {
        btnCalculate.Enabled = true;
        DataControl dc = new DataControl();
        DataTable Dtsku_Price = (DataTable)Session["Dtsku_Price"];
        PurchaseSKU = (DataTable)Session["PurchaseSKU"];
        DataRow[] foundRows = Dtsku_Price.Select("SKU_ID  = '" + ddlSKuCde.SelectedValue + "'");
        string[] selectedTex = ddlSKuCde.SelectedItem.Text.Split(':');
        if (btnSave.Text == "Add")
        {
            if (CheckDublicateSKU())
            {
                DataRow dr = PurchaseSKU.NewRow();
                dr["SKU_ID"] = ddlSKuCde.SelectedValue;
                dr["SKU_Code"] = selectedTex[1].ToString();
                dr["SKU_Name"] = selectedTex[0].ToString();
                dr["UNITS_IN_CASE"] = foundRows[0]["UNITS_IN_CASE"];
                dr["QUANTITY_UNIT"] = int.Parse(dc.chkNull_0(txtQuantity.Text));
                dr["QUANTITY_UNIT2"] = 0;
                dr["QUANTITY_CTN"] = int.Parse(dc.chkNull_0(txtCtn.Text));
                dr["QUANTITY_CTN2"] = 0;
                dr["FREE_SKU"] = int.Parse(dc.chkNull_0(txtFreeSku.Text));
                dr["QUANTITY"] = int.Parse(dc.chkNull_0(txtCtn.Text)) * int.Parse(dr["UNITS_IN_CASE"].ToString()) + int.Parse(dc.chkNull_0(txtQuantity.Text));
                dr["UNIT_PRICE"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString()); //dc.chkNull_0(txtUnitRate.Text);
                if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                {
                    dr["AMOUNT"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
                    dr["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    dr["TST_AMOUNT"] = 0;
                    dr["BATCH_NO"] = "T";
                }
                else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                {
                    dr["AMOUNT"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
                    dr["TST_AMOUNT"] = decimal.Parse(foundRows[0]["GST_RATE_TP"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
                    dr["GST_RATE"] = 0;
                    dr["BATCH_NO"] = "R";
                }
                else
                {
                    dr["AMOUNT"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
                    dr["TST_AMOUNT"] = 0;
                    dr["GST_RATE"] = 0;
                    dr["BATCH_NO"] = "E";
                }
                dr["STANDARD_DISCOUNT"] = 0;
                dr["GST_AMOUNT"] = 0;
                dr["Stock"] = txtStock.Text;

                if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 1)
                {
                    dr["EXTRA_DISCOUNT_PER"] = decimal.Parse(dc.chkNull_0(txtExtDiscount.Text));
                    dr["EXTRA_DISCOUNT"] = Convert.ToDecimal(dr["AMOUNT"]) * decimal.Parse(dc.chkNull_0(txtExtDiscount.Text)) / 100;
                }
                else if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 2)
                {
                    dr["EXTRA_DISCOUNT_PER"] = 0;
                    dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(txtExtDiscountValue.Text));
                }
                dr["NET_AMOUNT"] = Convert.ToDecimal(dr["AMOUNT"]) - Convert.ToDecimal(dr["EXTRA_DISCOUNT"]);
                dr["returnType"] = drpReturnType.SelectedValue;//for Sale Return
                dr["PRINCIPAL_ID"] = ddlPrincipal.SelectedValue;

                if (btnSaveOrder.Text == "Save Invoice" || btnSaveOrder.Text == "Update Invoice")
                {
                    int Stock = 0;
                    if (txtStock.Text != "0-0")
                    {
                        string[] strstock = txtStock.Text.Split('-');
                        Stock += Convert.ToInt32(strstock[0]) * Convert.ToInt32(dr["UNITS_IN_CASE"]);
                        Stock += Convert.ToInt32(strstock[1]);
                    }
                    if ((int.Parse(dr["QUANTITY_CTN"].ToString()) * int.Parse(dr["UNITS_IN_CASE"].ToString())) + int.Parse(dr["QUANTITY_UNIT"].ToString()) + int.Parse(dr["FREE_SKU"].ToString()) > Stock)
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('Available Stock is " + txtStock.Text + "')", true);
                        return;
                    }
                }
                PurchaseSKU.Rows.Add(dr);
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('SKU Already Exists')", true);
            }
        }
        else
        {
            DataRow dr = PurchaseSKU.Rows[RowId];
            dr["SKU_ID"] = ddlSKuCde.SelectedValue;
            dr["SKU_Code"] = selectedTex[1].ToString();
            dr["SKU_Name"] = selectedTex[0].ToString();
            dr["UNITS_IN_CASE"] = foundRows[0]["UNITS_IN_CASE"];
            dr["BATCH_NO"] = "";
            dr["QUANTITY_UNIT"] = int.Parse(dc.chkNull_0(txtQuantity.Text));
            dr["QUANTITY_UNIT2"] = dr["QUANTITY_UNIT2"];
            dr["QUANTITY_CTN"] = int.Parse(dc.chkNull_0(txtCtn.Text));
            dr["QUANTITY_CTN2"] = dr["QUANTITY_CTN2"];
            dr["FREE_SKU"] = int.Parse(dc.chkNull_0(txtFreeSku.Text));
            dr["UNIT_PRICE"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString());//dc.chkNull_0(txtUnitRate.Text);
            dr["QUANTITY"] = int.Parse(dc.chkNull_0(txtCtn.Text)) * int.Parse(dr["UNITS_IN_CASE"].ToString()) + int.Parse(dc.chkNull_0(txtQuantity.Text));
            dr["AMOUNT"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
            if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
            {
                dr["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                dr["AMOUNT"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
                dr["TST_AMOUNT"] = 0;
                dr["BATCH_NO"] = "T";
            }
            else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
            {
                dr["AMOUNT"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
                dr["TST_AMOUNT"] = decimal.Parse(foundRows[0]["GST_RATE_TP"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
                dr["GST_RATE"] = 0;
                dr["BATCH_NO"] = "R";
            }
            else
            {
                dr["AMOUNT"] = decimal.Parse(foundRows[0]["TRADE_PRICE"].ToString()) * (decimal.Parse(dr["QUANTITY_UNIT"].ToString()) + (decimal.Parse(dr["QUANTITY_CTN"].ToString()) * decimal.Parse(dr["UNITS_IN_CASE"].ToString())));
                dr["TST_AMOUNT"] = 0;
                dr["GST_RATE"] = 0;
                dr["BATCH_NO"] = "E";
            }
            if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 1)
            {
                dr["EXTRA_DISCOUNT_PER"] = decimal.Parse(dc.chkNull_0(txtExtDiscount.Text));
                dr["EXTRA_DISCOUNT"] = Convert.ToDecimal(dr["AMOUNT"]) * decimal.Parse(dc.chkNull_0(txtExtDiscount.Text)) / 100;
            }
            else if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 2)
            {
                dr["EXTRA_DISCOUNT_PER"] = 0;
                dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(txtExtDiscountValue.Text));
            }
            dr["NET_AMOUNT"] = Convert.ToDecimal(dr["AMOUNT"]) - Convert.ToDecimal(dr["EXTRA_DISCOUNT"]);
            dr["STANDARD_DISCOUNT"] = 0;
            dr["GST_AMOUNT"] = 0;
            dr["STANDARD_DISCOUNT_PER"] = 0;
            dr["returnType"] = drpReturnType.SelectedValue;//for Sale Return
            dr["PRINCIPAL_ID"] = ddlPrincipal.SelectedValue;
            if (btnSaveOrder.Text == "Save Invoice" || btnSaveOrder.Text == "Update Invoice")
            {
                int Stock = 0;
                if (txtStock.Text != "0-0")
                {
                    string[] strstock = txtStock.Text.Split('-');
                    Stock += Convert.ToInt32(strstock[0]) * Convert.ToInt32(dr["UNITS_IN_CASE"]);
                    Stock += Convert.ToInt32(strstock[1]);
                }
                if ((int.Parse(dr["QUANTITY_CTN"].ToString()) * int.Parse(dr["UNITS_IN_CASE"].ToString())) + int.Parse(dr["QUANTITY_UNIT"].ToString()) + int.Parse(dr["FREE_SKU"].ToString()) > Stock)
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('Available Stock is " + txtStock.Text + "')", true);
                    return;
                }
            }
        }
        Session.Add("PurchaseSKU", PurchaseSKU);
        EnableDisableController(false);
        //btnCalculate_Click(null, null);
        ClearAll();
        LoadGird();
        ScriptManager.GetCurrent(Page).SetFocus(ddlSKuCde);
    }
    protected void btnCalculate_Click(object sender, EventArgs e)
    {
        if (FindCustomer())
        {
            decimal ObjGrossSales = 0;
            decimal ObjStandardDiscount = 0;
            decimal ObjExtraDiscount = 0;
            decimal ObjTotalGST = 0;
            decimal ObjTotalSED = 0;
            decimal ObjTotalTST = 0;
            decimal ObjFreeGSTAmount = 0;
            decimal ObjFreeGrossAmt = 0;
            decimal ObjFreeTST = 0;
            // decimal ProductExtDist = 0;
            decimal ProductStdDist = 0;

            OrderEntryController or = new OrderEntryController();
            SKUGroupController GroupCtl = new SKUGroupController();
            DataTable dt = (DataTable)Session["PurchaseSKU"];
            foreach (DataRow dr in dt.Rows)
            {
                dr["CLAIM_EXTRA_AMOUNT"] = 0;
                dr["CLAIM_STANDARD_DISCOUNT"] = 0;
                dr["CLAIM_PER"] = 0;
            }
            PromotionCollections_Controller Arpc = (PromotionCollections_Controller)Session["Arpc"];
            CreateFreeSKU();
            LoadFreeGrid();

            if (dt.Rows.Count > 0 && drpCustomer.Items.Count > 0)
            {
                if (ChbDiscount.Checked == true)
                {
                    #region Discount
                    //ArrayList ARlistOffer = or.GetPromotionOffers(Arpc, mCustomerVolClassId, mCustomerTypeId, dt, false);

                    //for (int i = 0; i < ARlistOffer.Count; i++)
                    //{
                    //    PromoOffers_Collection pofferCol = (PromoOffers_Collection)ARlistOffer[i];

                    //    if (pofferCol.Is_Claimable)
                    //    {
                    //        if (pofferCol.SKU_ID > 0)
                    //        {
                    //            #region slab on SKU
                    //            foreach (DataRow dr in dt.Rows)
                    //            {
                    //                if (int.Parse(dr["SKU_ID"].ToString()) == pofferCol.SKU_ID)
                    //                {

                    //                    if (pofferCol.Is_And == true)
                    //                    {
                    //                        if (pofferCol.Offer_Value > 0)
                    //                        {
                    //                            //dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                            dr["CLAIM_EXTRA_AMOUNT"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_EXTRA_AMOUNT"].ToString())) + pofferCol.Offer_Value;

                    //                        }
                    //                        if (pofferCol.Discount > 0)
                    //                        {
                    //                            //dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100) * (decimal.Parse(dr["AMOUNT"].ToString()));
                    //                            dr["CLAIM_EXTRA_AMOUNT"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_EXTRA_AMOUNT"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100) * (decimal.Parse(dr["AMOUNT"].ToString()));
                    //                        }
                    //                    }
                    //                    else
                    //                    {
                    //                        if (pofferCol.Offer_Value > 0)
                    //                        {
                    //                            //dr["STANDARD_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                            //dr["CLAIM_AMOUNT"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_AMOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                        }
                    //                        if (pofferCol.Discount > 0)
                    //                        {
                    //                            //  dr["STANDARD_DISCOUNT_PER"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT_PER"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100);
                    //                            dr["CLAIM_PER"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_PER"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100);
                    //                        }
                    //                    }
                    //                    if (pofferCol.Free_SKU_ID > 0)
                    //                    {
                    //                        DataTable Dtsku_Price = (DataTable)Session["Dtsku_Price"];
                    //                        DataRow[] foundRows = Dtsku_Price.Select("SKU_ID  = '" + pofferCol.Free_SKU_ID.ToString() + "'");
                    //                        int RowId = FreeSKUExist(dtFreeSKU, pofferCol.Free_SKU_ID);
                    //                        if (foundRows.Length > 0)
                    //                        {
                    //                            if (RowId < 0)
                    //                            {
                    //                                DataRow dr1 = dtFreeSKU.NewRow();
                    //                                dr1["SKU_ID"] = foundRows[0]["SKU_ID"];
                    //                                dr1["SKU_Code"] = foundRows[0]["SKU_CODE"];
                    //                                dr1["SKU_Name"] = foundRows[0]["SKU_NAME"];
                    //                                dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString()));
                    //                                if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                    //                                {


                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = (decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) / 100) * decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                    dr1["TST_AMOUNT"] = 0;

                    //                                }
                    //                                else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                    //                                {

                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["TST_AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = 0;
                    //                                }
                    //                                else
                    //                                {

                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = 0;
                    //                                    dr1["TST_AMOUNT"] = 0;
                    //                                }

                    //                                dr1["PROMOTION_ID"] = pofferCol.Promotion_ID.ToString();
                    //                                dr1["BASKET_ID"] = pofferCol.Basket_ID.ToString();
                    //                                dr1["BASKET_DETAIL_ID"] = pofferCol.BasketDetail_ID.ToString();
                    //                                dr1["PROMOTION_OFFER_ID"] = pofferCol.PromoOffer_ID.ToString();
                    //                                ObjFreeGrossAmt += decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                ObjFreeGSTAmount += decimal.Parse(dc.chkNull_0(dr1["GST_AMOUNT"].ToString()));
                    //                                ObjFreeTST += decimal.Parse(dc.chkNull_0(dr1["TST_AMOUNT"].ToString()));
                    //                                dtFreeSKU.Rows.Add(dr1);
                    //                            }
                    //                            else
                    //                            {

                    //                                DataRow dr1 = dtFreeSKU.Rows[RowId];
                    //                                ObjFreeGrossAmt = ObjFreeGrossAmt - decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                ObjFreeGSTAmount = ObjFreeGSTAmount - decimal.Parse(dr1["GST_AMOUNT"].ToString());
                    //                                dr1["SKU_ID"] = foundRows[0]["SKU_ID"];
                    //                                dr1["SKU_Code"] = foundRows[0]["SKU_CODE"];
                    //                                dr1["SKU_Name"] = foundRows[0]["SKU_NAME"];
                    //                                dr1["Quantity"] = int.Parse(dr1["Quantity"].ToString()) + pofferCol.Quantity;
                    //                                dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString()));
                    //                                if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                    //                                {


                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = (decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) / 100) * decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                    dr1["TST_AMOUNT"] = 0;

                    //                                }
                    //                                else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                    //                                {
                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["TST_AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = 0;
                    //                                }
                    //                                else
                    //                                {
                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                }
                    //                                dr1["PROMOTION_ID"] = pofferCol.Promotion_ID.ToString();
                    //                                dr1["BASKET_ID"] = pofferCol.Basket_ID.ToString();
                    //                                dr1["BASKET_DETAIL_ID"] = pofferCol.BasketDetail_ID.ToString();
                    //                                dr1["PROMOTION_OFFER_ID"] = pofferCol.PromoOffer_ID.ToString();
                    //                                ObjFreeGrossAmt += decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                ObjFreeGSTAmount += decimal.Parse(dc.chkNull_0(dr1["GST_AMOUNT"].ToString()));
                    //                                ObjFreeTST += decimal.Parse(dc.chkNull_0(dr1["TST_AMOUNT"].ToString()));
                    //                            }

                    //                        }
                    //                    }
                    //                    Session.Add("dtFreeSKU", dtFreeSKU);
                    //                    LoadFreeGrid();
                    //                }
                    //            }
                    //            #endregion
                    //        }
                    //        else
                    //        {
                    //            #region slab on SKU Group

                    //            #region Free SKU
                    //            if (pofferCol.Free_SKU_ID > 0)
                    //            {
                    //                DataTable Dtsku_Price = (DataTable)Session["Dtsku_Price"];
                    //                DataRow[] foundRows = Dtsku_Price.Select("SKU_ID  = '" + pofferCol.Free_SKU_ID.ToString() + "'");
                    //                int RowId = FreeSKUExist(dtFreeSKU, pofferCol.Free_SKU_ID);
                    //                if (foundRows.Length > 0)
                    //                {
                    //                    if (RowId < 0)
                    //                    {
                    //                        DataRow dr1 = dtFreeSKU.NewRow();
                    //                        dr1["SKU_ID"] = foundRows[0]["SKU_ID"];
                    //                        dr1["SKU_Code"] = foundRows[0]["SKU_CODE"];
                    //                        dr1["SKU_Name"] = foundRows[0]["SKU_NAME"];
                    //                        dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(txtUnitRate.Text));
                    //                        dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                        dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString()));
                    //                        if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                    //                        {

                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = (decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) / 100) * decimal.Parse(dr1["AMOUNT"].ToString());
                    //                            dr1["TST_AMOUNT"] = 0;

                    //                        }
                    //                        else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                    //                        {
                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["TST_AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = 0;
                    //                        }
                    //                        else
                    //                        {
                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = 0;
                    //                            dr1["TST_AMOUNT"] = 0;
                    //                        }
                    //                        dr1["PROMOTION_ID"] = pofferCol.Promotion_ID.ToString();
                    //                        dr1["BASKET_ID"] = pofferCol.Basket_ID.ToString();
                    //                        dr1["BASKET_DETAIL_ID"] = pofferCol.BasketDetail_ID.ToString();
                    //                        dr1["PROMOTION_OFFER_ID"] = pofferCol.PromoOffer_ID.ToString();
                    //                        ObjFreeGrossAmt += decimal.Parse(dr1["AMOUNT"].ToString());
                    //                        ObjFreeGSTAmount += decimal.Parse(dc.chkNull_0(dr1["GST_AMOUNT"].ToString()));
                    //                        ObjFreeTST += decimal.Parse(dc.chkNull_0(dr1["TST_AMOUNT"].ToString()));
                    //                        dtFreeSKU.Rows.Add(dr1);
                    //                    }
                    //                    else
                    //                    {

                    //                        DataRow dr1 = dtFreeSKU.Rows[RowId];
                    //                        ObjFreeGrossAmt = ObjFreeGrossAmt - decimal.Parse(dr1["AMOUNT"].ToString());
                    //                        ObjFreeGSTAmount = ObjFreeGSTAmount - decimal.Parse(dr1["GST_AMOUNT"].ToString());
                    //                        dr1["SKU_ID"] = foundRows[0]["SKU_ID"];
                    //                        dr1["SKU_Code"] = foundRows[0]["SKU_CODE"];
                    //                        dr1["SKU_Name"] = foundRows[0]["SKU_NAME"];
                    //                        dr1["Quantity"] = int.Parse(dr1["Quantity"].ToString()) + pofferCol.Quantity;
                    //                        dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString()));
                    //                        if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                    //                        {

                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = (decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) / 100) * decimal.Parse(dr1["AMOUNT"].ToString());
                    //                            dr1["TST_AMOUNT"] = 0;

                    //                        }
                    //                        else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                    //                        {
                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["TST_AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = 0;
                    //                        }
                    //                        else
                    //                        {
                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = 0;
                    //                            dr1["TST_AMOUNT"] = 0;
                    //                        }
                    //                        dr1["PROMOTION_ID"] = pofferCol.Promotion_ID.ToString();
                    //                        dr1["BASKET_ID"] = pofferCol.Basket_ID.ToString();
                    //                        dr1["BASKET_DETAIL_ID"] = pofferCol.BasketDetail_ID.ToString();
                    //                        dr1["PROMOTION_OFFER_ID"] = pofferCol.PromoOffer_ID.ToString();
                    //                        ObjFreeGrossAmt += decimal.Parse(dr1["AMOUNT"].ToString());
                    //                        ObjFreeGSTAmount += decimal.Parse(dc.chkNull_0(dr1["GST_AMOUNT"].ToString()));
                    //                        ObjFreeTST += decimal.Parse(dc.chkNull_0(dr1["TST_AMOUNT"].ToString()));
                    //                    }
                    //                    Session.Add("dtFreeSKU", dtFreeSKU);
                    //                    LoadFreeGrid();

                    //                }
                    //            }
                    //            #endregion


                    //            foreach (DataRow dr in dt.Rows)
                    //            {


                    //                if (GroupCtl.ExistsInGroup(Constants.IntNullValue, pofferCol.Group_ID, int.Parse(dr["SKU_ID"].ToString())))
                    //                {

                    //                    if (pofferCol.Is_And == true)
                    //                    {
                    //                        if (pofferCol.Offer_Value > 0)
                    //                        {
                    //                            //dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                            dr["CLAIM_EXTRA_AMOUNT"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_EXTRA_AMOUNT"].ToString())) + pofferCol.Offer_Value;

                    //                        }
                    //                        if (pofferCol.Discount > 0)
                    //                        {
                    //                            //dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100) * (decimal.Parse(dr["AMOUNT"].ToString()));
                    //                            dr["CLAIM_EXTRA_AMOUNT"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_EXTRA_AMOUNT"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100) * (decimal.Parse(dr["AMOUNT"].ToString()));
                    //                        }
                    //                    }
                    //                    else
                    //                    {
                    //                        if (pofferCol.Offer_Value > 0)
                    //                        {
                    //                            //dr["STANDARD_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                            //dr["CLAIM_AMOUNT"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_AMOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                        }
                    //                        if (pofferCol.Discount > 0)
                    //                        {
                    //                            // dr["STANDARD_DISCOUNT_PER"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT_PER"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100);
                    //                            dr["CLAIM_PER"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_PER"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100);
                    //                        }
                    //                    }

                    //                }
                    //            }
                    //            #endregion
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (pofferCol.SKU_ID > 0)
                    //        {
                    //            #region slab on SKU
                    //            foreach (DataRow dr in dt.Rows)
                    //            {
                    //                if (int.Parse(dr["SKU_ID"].ToString()) == pofferCol.SKU_ID)
                    //                {

                    //                    if (pofferCol.Is_And == true)
                    //                    {
                    //                        if (pofferCol.Offer_Value > 0)
                    //                        {
                    //                            //dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())) + pofferCol.Offer_Value;

                    //                        }
                    //                        if (pofferCol.Discount > 0)
                    //                        {
                    //                            //dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100) * (decimal.Parse(dr["AMOUNT"].ToString()));
                    //                        }
                    //                    }
                    //                    else
                    //                    {
                    //                        if (pofferCol.Offer_Value > 0)
                    //                        {
                    //                            // dr["STANDARD_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                        }
                    //                        if (pofferCol.Discount > 0)
                    //                        {
                    //                            // dr["STANDARD_DISCOUNT_PER"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT_PER"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100);
                    //                        }
                    //                    }
                    //                    if (pofferCol.Free_SKU_ID > 0)
                    //                    {
                    //                        DataTable Dtsku_Price = (DataTable)Session["Dtsku_Price"];
                    //                        DataRow[] foundRows = Dtsku_Price.Select("SKU_ID  = '" + pofferCol.Free_SKU_ID.ToString() + "'");
                    //                        int RowId = FreeSKUExist(dtFreeSKU, pofferCol.Free_SKU_ID);
                    //                        if (foundRows.Length > 0)
                    //                        {
                    //                            if (RowId < 0)
                    //                            {
                    //                                DataRow dr1 = dtFreeSKU.NewRow();
                    //                                dr1["SKU_ID"] = foundRows[0]["SKU_ID"];
                    //                                dr1["SKU_Code"] = foundRows[0]["SKU_CODE"];
                    //                                dr1["SKU_Name"] = foundRows[0]["SKU_NAME"];
                    //                                dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString()));
                    //                                if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                    //                                {

                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = (decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) / 100) * decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                    dr1["TST_AMOUNT"] = 0;

                    //                                }
                    //                                else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                    //                                {
                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["TST_AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = 0;
                    //                                }
                    //                                else
                    //                                {
                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = 0;
                    //                                    dr1["TST_AMOUNT"] = 0;
                    //                                }

                    //                                dr1["PROMOTION_ID"] = pofferCol.Promotion_ID.ToString();
                    //                                dr1["BASKET_ID"] = pofferCol.Basket_ID.ToString();
                    //                                dr1["BASKET_DETAIL_ID"] = pofferCol.BasketDetail_ID.ToString();
                    //                                dr1["PROMOTION_OFFER_ID"] = pofferCol.PromoOffer_ID.ToString();
                    //                                ObjFreeGrossAmt += decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                ObjFreeGSTAmount += decimal.Parse(dc.chkNull_0(dr1["GST_AMOUNT"].ToString()));
                    //                                ObjFreeTST += decimal.Parse(dc.chkNull_0(dr1["TST_AMOUNT"].ToString()));
                    //                                dtFreeSKU.Rows.Add(dr1);
                    //                            }
                    //                            else
                    //                            {

                    //                                DataRow dr1 = dtFreeSKU.Rows[RowId];
                    //                                ObjFreeGrossAmt = ObjFreeGrossAmt - decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                ObjFreeGSTAmount = ObjFreeGSTAmount - decimal.Parse(dr1["GST_AMOUNT"].ToString());
                    //                                dr1["SKU_ID"] = foundRows[0]["SKU_ID"];
                    //                                dr1["SKU_Code"] = foundRows[0]["SKU_CODE"];
                    //                                dr1["SKU_Name"] = foundRows[0]["SKU_NAME"];
                    //                                dr1["Quantity"] = int.Parse(dr1["Quantity"].ToString()) + pofferCol.Quantity;
                    //                                dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString()));
                    //                                if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                    //                                {

                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = (decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) / 100) * decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                    dr1["TST_AMOUNT"] = 0;

                    //                                }
                    //                                else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                    //                                {
                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["TST_AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                    dr1["GST_AMOUNT"] = 0;
                    //                                }
                    //                                else
                    //                                {
                    //                                    dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                                    dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                                    dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                                }
                    //                                dr1["PROMOTION_ID"] = pofferCol.Promotion_ID.ToString();
                    //                                dr1["BASKET_ID"] = pofferCol.Basket_ID.ToString();
                    //                                dr1["BASKET_DETAIL_ID"] = pofferCol.BasketDetail_ID.ToString();
                    //                                dr1["PROMOTION_OFFER_ID"] = pofferCol.PromoOffer_ID.ToString();
                    //                                ObjFreeGrossAmt += decimal.Parse(dr1["AMOUNT"].ToString());
                    //                                ObjFreeGSTAmount += decimal.Parse(dc.chkNull_0(dr1["GST_AMOUNT"].ToString()));
                    //                                ObjFreeTST += decimal.Parse(dc.chkNull_0(dr1["TST_AMOUNT"].ToString()));
                    //                            }

                    //                        }
                    //                    }
                    //                    Session.Add("dtFreeSKU", dtFreeSKU);
                    //                    LoadFreeGrid();
                    //                }
                    //            }
                    //            #endregion
                    //        }
                    //        else
                    //        {
                    //            #region slab on SKU Group

                    //            #region Free SKU
                    //            if (pofferCol.Free_SKU_ID > 0)
                    //            {
                    //                DataTable Dtsku_Price = (DataTable)Session["Dtsku_Price"];
                    //                DataRow[] foundRows = Dtsku_Price.Select("SKU_ID  = '" + pofferCol.Free_SKU_ID.ToString() + "'");
                    //                int RowId = FreeSKUExist(dtFreeSKU, pofferCol.Free_SKU_ID);
                    //                if (foundRows.Length > 0)
                    //                {
                    //                    if (RowId < 0)
                    //                    {
                    //                        DataRow dr1 = dtFreeSKU.NewRow();
                    //                        dr1["SKU_ID"] = foundRows[0]["SKU_ID"];
                    //                        dr1["SKU_Code"] = foundRows[0]["SKU_CODE"];
                    //                        dr1["SKU_Name"] = foundRows[0]["SKU_NAME"];
                    //                        dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(txtUnitRate.Text));
                    //                        dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                        dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString()));
                    //                        if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                    //                        {

                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = (decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) / 100) * decimal.Parse(dr1["AMOUNT"].ToString());
                    //                            dr1["TST_AMOUNT"] = 0;

                    //                        }
                    //                        else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                    //                        {
                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["TST_AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = 0;
                    //                        }
                    //                        else
                    //                        {
                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = 0;
                    //                            dr1["TST_AMOUNT"] = 0;
                    //                        }
                    //                        dr1["PROMOTION_ID"] = pofferCol.Promotion_ID.ToString();
                    //                        dr1["BASKET_ID"] = pofferCol.Basket_ID.ToString();
                    //                        dr1["BASKET_DETAIL_ID"] = pofferCol.BasketDetail_ID.ToString();
                    //                        dr1["PROMOTION_OFFER_ID"] = pofferCol.PromoOffer_ID.ToString();
                    //                        ObjFreeGrossAmt += decimal.Parse(dr1["AMOUNT"].ToString());
                    //                        ObjFreeGSTAmount += decimal.Parse(dc.chkNull_0(dr1["GST_AMOUNT"].ToString()));
                    //                        ObjFreeTST += decimal.Parse(dc.chkNull_0(dr1["TST_AMOUNT"].ToString()));
                    //                        dtFreeSKU.Rows.Add(dr1);
                    //                    }
                    //                    else
                    //                    {

                    //                        DataRow dr1 = dtFreeSKU.Rows[RowId];
                    //                        ObjFreeGrossAmt = ObjFreeGrossAmt - decimal.Parse(dr1["AMOUNT"].ToString());
                    //                        ObjFreeGSTAmount = ObjFreeGSTAmount - decimal.Parse(dr1["GST_AMOUNT"].ToString());
                    //                        dr1["SKU_ID"] = foundRows[0]["SKU_ID"];
                    //                        dr1["SKU_Code"] = foundRows[0]["SKU_CODE"];
                    //                        dr1["SKU_Name"] = foundRows[0]["SKU_NAME"];
                    //                        dr1["Quantity"] = int.Parse(dr1["Quantity"].ToString()) + pofferCol.Quantity;
                    //                        dr1["UNIT_PRICE"] = Convert.ToDecimal(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString()));
                    //                        if (foundRows[0]["GST_ON"].ToString().Trim() == "T")
                    //                        {

                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = (decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) / 100) * decimal.Parse(dr1["AMOUNT"].ToString());
                    //                            dr1["TST_AMOUNT"] = 0;

                    //                        }
                    //                        else if (foundRows[0]["GST_ON"].ToString().Trim() == "R")
                    //                        {
                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["TST_AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["GST_RATE_TP"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = 0;
                    //                        }
                    //                        else
                    //                        {
                    //                            dr1["Quantity"] = pofferCol.Quantity.ToString();
                    //                            dr1["GST_RATE"] = foundRows[0]["GST_RATE_TP"];
                    //                            dr1["AMOUNT"] = decimal.Parse(dc.chkNull_0(foundRows[0]["TRADE_PRICE"].ToString())) * decimal.Parse(pofferCol.Quantity.ToString());
                    //                            dr1["GST_AMOUNT"] = 0;
                    //                            dr1["TST_AMOUNT"] = 0;
                    //                        }
                    //                        dr1["PROMOTION_ID"] = pofferCol.Promotion_ID.ToString();
                    //                        dr1["BASKET_ID"] = pofferCol.Basket_ID.ToString();
                    //                        dr1["BASKET_DETAIL_ID"] = pofferCol.BasketDetail_ID.ToString();
                    //                        dr1["PROMOTION_OFFER_ID"] = pofferCol.PromoOffer_ID.ToString();
                    //                        ObjFreeGrossAmt += decimal.Parse(dr1["AMOUNT"].ToString());
                    //                        ObjFreeGSTAmount += decimal.Parse(dc.chkNull_0(dr1["GST_AMOUNT"].ToString()));
                    //                        ObjFreeTST += decimal.Parse(dc.chkNull_0(dr1["TST_AMOUNT"].ToString()));
                    //                    }
                    //                    Session.Add("dtFreeSKU", dtFreeSKU);
                    //                    LoadFreeGrid();

                    //                }
                    //            }
                    //            #endregion


                    //            foreach (DataRow dr in dt.Rows)
                    //            {


                    //                if (GroupCtl.ExistsInGroup(Constants.IntNullValue, pofferCol.Group_ID, int.Parse(dr["SKU_ID"].ToString())))
                    //                {


                    //                    if (pofferCol.Is_And == true)
                    //                    {
                    //                        if (pofferCol.Offer_Value > 0)
                    //                        {
                    //                            //dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                        }
                    //                        if (pofferCol.Discount > 0)
                    //                        {
                    //                            //dr["EXTRA_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())) + ((decimal.Parse(pofferCol.Discount.ToString())) / 100) * (decimal.Parse(dr["AMOUNT"].ToString()) - decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString())));
                    //                        }
                    //                    }
                    //                    else
                    //                    {
                    //                        if (pofferCol.Offer_Value > 0)
                    //                        {
                    //                            //  dr["STANDARD_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString())) + pofferCol.Offer_Value;
                    //                        }
                    //                        if (pofferCol.Discount > 0)
                    //                        {
                    //                            //  dr["STANDARD_DISCOUNT_PER"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT_PER"].ToString())) + (decimal.Parse(pofferCol.Discount.ToString()) / 100);
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //            #endregion
                    //        }
                    //    }
                    //}

                    //#region Recalulate Products

                    //foreach (DataRow dr in dt.Rows)
                    //{

                    //    // dr["STANDARD_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT_PER"].ToString())) * (decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString())));


                    //    dr["CLAIM_STANDARD_DISCOUNT"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_PER"].ToString())) * (decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString())) - decimal.Parse(dc.chkNull_0(dr["CLAIM_EXTRA_AMOUNT"].ToString())));
                    //    dr["SED_AMOUNT"] = decimal.Parse(dc.chkNull_0(dr["CLAIM_STANDARD_DISCOUNT"].ToString())) + decimal.Parse(dc.chkNull_0(dr["CLAIM_EXTRA_AMOUNT"].ToString()));

                    //    ObjGrossSales += decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString()));
                    //    ObjStandardDiscount += decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString()));
                    //    if (dr["SKU_ID"].ToString() == ddlSKuCde.SelectedValue)
                    //    {
                    //        if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 1 && sender == null)
                    //        {

                    //            dr["EXTRA_DISCOUNT"] = (decimal.Parse(dr["AMOUNT"].ToString()) - decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString()))) * decimal.Parse(dc.chkNull_0(txtExtDiscount.Text)) / 100;
                    //        }
                    //        else
                    //        {
                    //            // do thing
                    //        }

                    //    }
                    //    ObjExtraDiscount += decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString()));
                    //    ObjTotalSED += decimal.Parse(dc.chkNull_0(dr["SED_AMOUNT"].ToString()));

                    //    //dr["EXTRA_DISCOUNT"] = ProductExtDist * (decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString())) - decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString())));

                    //    #region GST Calculation

                    //    decimal TempAmount = decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString()));
                    //    decimal TempDiscount = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString()));
                    //    decimal TempExtraDisAmt = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString()));
                    //    dr["GST_AMOUNT"] = (TempAmount - TempDiscount - TempExtraDisAmt) * decimal.Parse(dc.chkNull_0(dr["GST_RATE"].ToString())) / 100;
                    //    dr["NET_AMOUNT"] = (TempAmount - TempDiscount - TempExtraDisAmt) + decimal.Parse(dc.chkNull_0(dr["GST_AMOUNT"].ToString())) + decimal.Parse(dc.chkNull_0(dr["TST_AMOUNT"].ToString()));
                    //    ObjTotalGST += decimal.Parse(dc.chkNull_0(dr["GST_AMOUNT"].ToString()));
                    //    ObjTotalTST += decimal.Parse(dc.chkNull_0(dr["TST_AMOUNT"].ToString()));

                    //    #endregion
                    //}

                    ////Reverse Calculation for Extra Discount
                    //if (decimal.Parse(dc.chkNull_0(numtxtTotalExtraDiscnt.Text)) > 0)
                    //{
                    //    ProductExtDist = decimal.Parse(dc.chkNull_0(numtxtTotalExtraDiscnt.Text)) / (ObjGrossSales - ObjStandardDiscount);
                    //}


                    //#endregion
                    #endregion
                }
                else
                {
                    #region Reversecalulate Products

                    //decimal ObjGrossSales = 0;
                    decimal ObjTaxableAmount = 0;
                    decimal ObjClaimAblePer = 0;

                    foreach (DataRow dr in dt.Rows)
                    {
                        ObjGrossSales += decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString()));

                    }

                    if (decimal.Parse(dc.chkNull_0(numtxtUnClaimabledist.Text)) > 0)
                    {
                        ObjClaimAblePer = decimal.Parse(dc.chkNull_0(numtxtUnClaimabledist.Text)) / ObjGrossSales;
                    }

                    if (decimal.Parse(dc.chkNull_0(numTxtTotalStndrdDiscnt.Text)) > 0)
                    {
                        ProductStdDist = decimal.Parse(dc.chkNull_0(numTxtTotalStndrdDiscnt.Text)) / (ObjGrossSales - ObjTaxableAmount);
                    }

                    //if (decimal.Parse(dc.chkNull_0(numtxtTotalExtraDiscnt.Text)) > 0)
                    //{
                    //    ProductExtDist = decimal.Parse(dc.chkNull_0(numtxtTotalExtraDiscnt.Text)) / (ObjGrossSales - ObjTaxableAmount);
                    //}

                    ObjGrossSales = 0;

                    foreach (DataRow dr in dt.Rows)
                    {
                        decimal TempAmount = 0;
                        decimal TempDiscount = 0;
                        decimal TempExtraDisAmt = 0;

                        //dr["EXTRA_DISCOUNT"] = ProductExtDist * decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString()));
                        //  dr["STANDARD_DISCOUNT"] = ProductStdDist * decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString()));
                        dr["SED_AMOUNT"] = ObjClaimAblePer * decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString()));

                        #region GST Calculation

                        TempAmount = decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString()));
                        TempDiscount = decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString()));
                        TempExtraDisAmt = decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString()));

                        dr["GST_AMOUNT"] = (TempAmount - TempDiscount - TempExtraDisAmt) * decimal.Parse(dc.chkNull_0(dr["GST_RATE"].ToString())) / 100;
                        dr["NET_AMOUNT"] = (TempAmount - TempDiscount - TempExtraDisAmt) + decimal.Parse(dc.chkNull_0(dr["GST_AMOUNT"].ToString())) + decimal.Parse(dc.chkNull_0(dr["TST_AMOUNT"].ToString()));

                        ObjGrossSales += decimal.Parse(dc.chkNull_0(dr["AMOUNT"].ToString()));
                        ObjTotalGST += decimal.Parse(dc.chkNull_0(dr["GST_AMOUNT"].ToString()));
                        ObjTotalTST += decimal.Parse(dc.chkNull_0(dr["TST_AMOUNT"].ToString()));
                        ObjTotalSED += decimal.Parse(dc.chkNull_0(dr["SED_AMOUNT"].ToString()));

                        ObjStandardDiscount += decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString()));
                        //if (dr["SKU_ID"].ToString() == ddlSKuCde.SelectedValue)
                        //{
                        //    if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 1)
                        //    {
                        //        dr["EXTRA_DISCOUNT"] = (decimal.Parse(dr["AMOUNT"].ToString()) - decimal.Parse(dc.chkNull_0(dr["STANDARD_DISCOUNT"].ToString()))) * decimal.Parse(dc.chkNull_0(txtExtDiscount.Text)) / 100;
                        //    }
                        //    else if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 2)
                        //    {
                        //        // do nothing                           
                        //    }
                        //}
                        ObjExtraDiscount += decimal.Parse(dc.chkNull_0(dr["EXTRA_DISCOUNT"].ToString()));

                        #endregion
                    }
                    #endregion
                }

                #region Set Total Values

                txtGrossAmount.Text = Convert.ToString(Math.Round(ObjGrossSales, 2));
                //     numTxtTotalStndrdDiscnt.Text = Convert.ToString(Math.Round(ObjStandardDiscount, 2));
                numtxtTotalExtraDiscnt.Text = Convert.ToString(Math.Round(ObjExtraDiscount, 2));
                numTxtTotalGST.Text = Convert.ToString(Math.Round((ObjTotalGST + ObjFreeGSTAmount), 2));
                numTxtTotalTST.Text = Convert.ToString(Math.Round(ObjTotalTST + ObjFreeTST, 2));
                numTxtTotlAmnt.Text = Convert.ToString(Math.Round(ObjGrossSales - Convert.ToDecimal(dc.chkNull_0(numtxtTotalExtraDiscnt.Text)) - Convert.ToDecimal(dc.chkNull_0(numTxtTotalStndrdDiscnt.Text)) + ObjTotalGST + ObjFreeGSTAmount + ObjTotalTST + ObjFreeTST, 2));
                numtxtUnClaimabledist.Text = Convert.ToString(Math.Round(ObjTotalSED, 2));

                #endregion
                if (RblPayMode.SelectedIndex == 1)
                {
                    txtCashReceived.Enabled = true;
                    ScriptManager.GetCurrent(Page).SetFocus(txtCashReceived);
                }
                else
                {
                    txtCashReceived.Enabled = false;
                }
            }
            btnSaveOrder.Enabled = true;
            btnSaveInvoice.Enabled = true;
        }
    }
    protected void btnSaveOrder_Click(object sender, EventArgs e)
    {
        if (FindCustomer())
        {
            if (IsDayClosed())
            {
                UserController UserCtl = new UserController();

                UserCtl.InsertUserLogoutTime(Convert.ToInt32(Session["User_Log_ID"]), Convert.ToInt32(Session["UserID"]));
                Session.Clear();
                System.Web.Security.FormsAuthentication.SignOut();
                Response.Redirect("../Login.aspx");
            }
            else
            {
                string ManualID = null;
                if (txtBillBookNo.Text.Trim().Length > 0)
                {
                    ManualID = txtBillBookNo.Text.ToUpper();
                }
                DataControl DC = new DataControl();
                PurchaseSKU = (DataTable)Session["PurchaseSKU"];
                dtFreeSKU = (DataTable)Session["dtFreeSKU"];

                ///
                /// Remove unchecked free skus
                ///

                //foreach (GridViewRow gvr in GrdFreeSKU.Rows)
                //{
                //    CheckBox cbSelect = (CheckBox)gvr.FindControl("cbSelect");
                //    if (!cbSelect.Checked)
                //    {
                //        foreach (DataRow dr in dtFreeSKU.Rows)
                //        {
                //            if (dr["SKU_ID"].ToString() == gvr.Cells[1].Text)
                //            {
                //                int RowId = FreeSKUExist(dtFreeSKU, Convert.ToInt32(dr["SKU_ID"]));
                //                dtFreeSKU.Rows.RemoveAt(RowId);
                //                break;
                //            }
                //        }
                //    }
                //}

                OrderEntryController mOrderController = new OrderEntryController();

                if (btnSaveOrder.Text == "Save Order")
                {
                    if (!IsBillBookNoExist())
                    {
                        bool IsValidInsert = mOrderController.Add_Order(int.Parse(Session["DistributorId"].ToString()), ManualID, mTownId, (long)Session["AreaId"], 0, long.Parse(Session["CUSTOMER_ID"].ToString()), long.Parse(Session["CUSTOMER_ID"].ToString()), int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), int.Parse(RblPayMode.SelectedValue.ToString()),
                        decimal.Parse(DC.chkNull_0(txtGrossAmount.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalStndrdDiscnt.Text)), decimal.Parse(DC.chkNull_0(numtxtTotalExtraDiscnt.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalGST.Text)), decimal.Parse(DC.chkNull_0(numTxtTotlAmnt.Text)), 0, Constants.Order_Pending_Id, PurchaseSKU, dtFreeSKU, int.Parse(Session["UserId"].ToString()), DateTime.Parse(Session["OrderDate"].ToString()), decimal.Parse(DC.chkNull_0(numtxtUnClaimabledist.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalTST.Text)), txtRemarks.Text);
                        if (IsValidInsert)
                        {
                            LoadSKUDetail();
                            LoadPendingOrder();
                            ClearMasterALL();
                            Session.Remove("hfBillBookNo");
                            ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
                        }
                    }
                    else
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('This Bill Book No already exist,Kindly enter different Bill Book No');", true);
                    }
                }
                else if (btnSaveOrder.Text == "Update Order")
                {

                    bool IsValidInsert = mOrderController.Update_Order(long.Parse(drpDocumentNo.SelectedValue), int.Parse(Session["DistributorId"].ToString()), ManualID, mTownId, (long)Session["AreaId"], 0, long.Parse(Session["CUSTOMER_ID"].ToString()), long.Parse(Session["CUSTOMER_ID"].ToString()), int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), int.Parse(RblPayMode.SelectedValue.ToString()),
                    decimal.Parse(DC.chkNull_0(txtGrossAmount.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalStndrdDiscnt.Text)), decimal.Parse(DC.chkNull_0(numtxtTotalExtraDiscnt.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalGST.Text)), decimal.Parse(DC.chkNull_0(numTxtTotlAmnt.Text)), 0, Constants.Order_Pending_Id, PurchaseSKU, dtFreeSKU, int.Parse(Session["UserId"].ToString()), DateTime.Parse(Session["OrderDate"].ToString()), decimal.Parse(DC.chkNull_0(numtxtUnClaimabledist.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalTST.Text)), txtRemarks.Text);
                    if (IsValidInsert)
                    {
                        LoadSKUDetail();
                        LoadPendingOrder();
                        ClearMasterALL();
                        Session.Remove("hfBillBookNo");
                        btnSaveOrder.Text = "Save Order";
                        ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
                    }

                }


                else if (btnSaveOrder.Text == "Save Invoice")
                {
                    if (int.Parse(RblPayMode.SelectedValue.ToString()) == Constants.Credit_Order_Id || int.Parse(RblPayMode.SelectedValue.ToString()) == Constants.Advance_PaymentOrder_id)
                    {
                        #region Check Credit Limit
                        //  CustomerDataController CDC = new CustomerDataController();
                        // DataTable dt = CDC.SelectCustomerCreditBalance(long.Parse(Session["CUSTOMER_ID"].ToString()), Constants.IntNullValue, int.Parse(Session["DistributorId"].ToString()), int.Parse(RblPayMode.SelectedValue.ToString()));
                        //  decimal NetCashSale = decimal.Parse(numTxtTotlAmnt.Text) - decimal.Parse(DC.chkNull_0(txtCashReceived.Text));

                        //if (decimal.Parse(DC.chkNull_0(dt.Rows[0][0].ToString())) <= NetCashSale)
                        //{
                        //    ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('Customer Credit Limit/Advance Amount is " + DC.chkNull_0(dt.Rows[0][0].ToString()) + "');", true);
                        //    return;
                        //}

                        //  if (int.Parse(RblPayMode.SelectedValue.ToString()) == Constants.Credit_Order_Id)
                        //  {
                        //    LedgerController LController = new LedgerController();
                        //  DataTable dtCredit = LController.SelectCreditPendingInvoice(int.Parse(Session["DistributorId"].ToString()), Constants.IntNullValue, long.Parse(Session["CUSTOMER_ID"].ToString()), 0);
                        // if (dtCredit != null)
                        // {

                        //if (dtCredit.Rows.Count >= Convert.ToInt32(Session["INVOICE_COUNT"]))
                        //{
                        //    ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('Customer Credit Invoice Limit is " + Session["INVOICE_COUNT"].ToString() + "');", true);
                        //    return;
                        //}
                        // }
                        // }
                        #endregion
                    }

                    #region Check Stock Limit
                    //if (dtFreeSKU.Rows.Count <= 0)
                    //{
                    foreach (DataRow dr in PurchaseSKU.Rows)
                    {
                        PhaysicalStockController mController = new PhaysicalStockController();
                        DataTable dtstock = mController.SelectSKUClosingStock(int.Parse(Session["DistributorId"].ToString()), int.Parse(dr["SKU_ID"].ToString()), "", DateTime.Parse(Session["OrderDate"].ToString()));
                        if (dtstock.Rows.Count > 0)
                        {
                            int Total_Units = (int.Parse(dr["UNITS_IN_CASE"].ToString()) * int.Parse(dr["QUANTITY_CTN"].ToString())) + int.Parse(dr["QUANTITY_UNIT"].ToString()) + int.Parse(dr["FREE_SKU"].ToString());

                            if (int.Parse(dtstock.Rows[0][0].ToString()) < Total_Units)
                            {
                                ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + dr["SKU_Code"].ToString() + " Current Stock is " + dtstock.Rows[0][0].ToString() + "');", true);
                                return;
                            }
                            else if (Convert.ToDecimal(dtstock.Rows[0][1]) != Convert.ToDecimal(dr["UNIT_PRICE"]))
                            {
                                ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + dr["SKU_Code"].ToString() + " Order Rate is " + dr["UNIT_PRICE"].ToString() + " and Current Rate is " + dtstock.Rows[0][1].ToString() + "');", true);
                                return;
                            }
                        }
                        else
                        {
                            ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert(' + " + dr["SKU_Code"].ToString() + "No Stock Found');", true);
                            return;
                        }
                    }
                    //}
                    // else
                    {
                        //foreach (DataRow dr in PurchaseSKU.Rows)
                        //{
                        //    PhaysicalStockController mController = new PhaysicalStockController();
                        //    DataTable dtstock = mController.SelectSKUClosingStock(int.Parse(Session["DistributorId"].ToString()), int.Parse(dr["SKU_ID"].ToString()), "", DateTime.Parse(Session["OrderDate"].ToString()));
                        //    if (dtstock.Rows.Count > 0)
                        //    {
                        //        foreach (DataRow dr2 in dtFreeSKU.Rows)
                        //        {
                        //            if (dr2["SKU_ID"].ToString() == dr["SKU_ID"].ToString())
                        //            {
                        //                int Total_Units = Convert.ToInt32(dr2["Quantity"]) + (int.Parse(dr["UNITS_IN_CASE"].ToString()) * int.Parse(dr["QUANTITY_CTN"].ToString())) + int.Parse(dr["QUANTITY_UNIT"].ToString());
                        //                if (int.Parse(dtstock.Rows[0][0].ToString()) < Total_Units)
                        //                {
                        //                    ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + dr["SKU_Code"].ToString() + " Current Stock is " + dtstock.Rows[0][0].ToString() + "');", true);
                        //                    return;
                        //                }
                        //                else if (Convert.ToDecimal(dtstock.Rows[0][1]) != Convert.ToDecimal(dr["UNIT_PRICE"]))
                        //                {
                        //                    ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + dr["SKU_Code"].ToString() + " Order Rate is " + dr["UNIT_PRICE"].ToString() + " and Current Rate is " + dtstock.Rows[0][1].ToString() + "');", true);
                        //                    return;
                        //                }
                        //                break;
                        //            }
                        //        }
                        //    }
                        //    else
                        //    {
                        //        ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert(' + " + dr["SKU_Code"].ToString() + "No Stock Found');", true);
                        //        return;
                        //    }
                        //}
                        //foreach (DataRow dr in dtFreeSKU.Rows)
                        //{
                        //    PhaysicalStockController mController = new PhaysicalStockController();
                        //    DataTable dtstock = mController.SelectSKUClosingStock(int.Parse(Session["DistributorId"].ToString()), int.Parse(dr["SKU_ID"].ToString()), "", DateTime.Parse(Session["OrderDate"].ToString()));
                        //    if (dtstock.Rows.Count > 0)
                        //    {
                        //        if (!IsFreeSKUinOrder(dr["SKU_ID"].ToString()))
                        //        {
                        //            int Total_Units = Convert.ToInt32(dr["Quantity"]);
                        //            if (int.Parse(dtstock.Rows[0][0].ToString()) < Total_Units)
                        //            {
                        //                ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + dr["SKU_Code"].ToString() + " Current Stock is " + dtstock.Rows[0][0].ToString() + "');", true);
                        //                return;
                        //            }
                        //            else if (Convert.ToDecimal(dtstock.Rows[0][1]) != Convert.ToDecimal(dr["UNIT_PRICE"]))
                        //            {
                        //                ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + dr["SKU_Code"].ToString() + " Order Rate is " + dr["UNIT_PRICE"].ToString() + " and Current Rate is " + dtstock.Rows[0][1].ToString() + "');", true);
                        //                return;
                        //            }
                        //            break;
                        //        }
                        //    }
                        //    else
                        //    {
                        //        ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert(' + " + dr["SKU_Code"].ToString() + "No Stock Found');", true);
                        //        return;
                        //    }
                        //}
                    }
                    #endregion

                    bool IsValidInsert = mOrderController.Add_Invoice(int.Parse(Session["DistributorId"].ToString()), ManualID, mTownId, long.Parse(Session["AreaId"].ToString()), 0, long.Parse(Session["CUSTOMER_ID"].ToString()), long.Parse(Session["CUSTOMER_ID"].ToString()), int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), long.Parse(drpDocumentNo.SelectedValue.ToString()),
                    decimal.Parse(DC.chkNull_0(txtGrossAmount.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalStndrdDiscnt.Text)), decimal.Parse(DC.chkNull_0(numtxtTotalExtraDiscnt.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalGST.Text)), decimal.Parse(DC.chkNull_0(numTxtTotlAmnt.Text)), 0, int.Parse(RblPayMode.SelectedValue.ToString()),
                    PurchaseSKU, dtFreeSKU, int.Parse(Session["UserId"].ToString()), decimal.Parse(DC.chkNull_0(txtCashReceived.Text)), DateTime.Parse(Session["OrderDate"].ToString()), decimal.Parse(DC.chkNull_0(numTxtTotalTST.Text)), decimal.Parse(DC.chkNull_0(numtxtUnClaimabledist.Text)), txtRemarks.Text);
                    if (IsValidInsert)
                    {
                        LoadCustomerData();
                        LoadPendingOrder();
                        ClearMasterALL();
                        Session.Remove("hfBillBookNo");
                        LoadSKUDetail();
                        ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
                    }
                }
                else if (btnSaveOrder.Text == "Update Invoice")
                {

                    if (int.Parse(RblPayMode.SelectedValue.ToString()) == Constants.Credit_Order_Id || int.Parse(RblPayMode.SelectedValue.ToString()) == Constants.Advance_PaymentOrder_id)
                    {
                        #region Check Credit Limit
                        //CustomerDataController CDC = new CustomerDataController();
                        //DataTable dt = CDC.SelectCustomerCreditBalance(long.Parse(Session["CUSTOMER_ID"].ToString()), Constants.IntNullValue, int.Parse(Session["DistributorId"].ToString()), int.Parse(RblPayMode.SelectedValue.ToString()));
                        //decimal NetCashSale = decimal.Parse(numTxtTotlAmnt.Text) - decimal.Parse(DC.chkNull_0(txtCashReceived.Text));

                        ////if (decimal.Parse(DC.chkNull_0(dt.Rows[0][0].ToString())) <= NetCashSale)
                        ////{
                        ////    ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('Customer Credit Limit/Advance Amount is " + DC.chkNull_0(dt.Rows[0][0].ToString()) + "');", true);
                        ////    return;
                        ////}

                        //if (int.Parse(RblPayMode.SelectedValue.ToString()) == Constants.Credit_Order_Id)
                        //{
                        //    LedgerController LController = new LedgerController();
                        //    DataTable dtCredit = LController.SelectCreditPendingInvoice(int.Parse(Session["DistributorId"].ToString()), Constants.IntNullValue, long.Parse(Session["CUSTOMER_ID"].ToString()), 0);

                        //    //if (dtCredit.Rows.Count >= Convert.ToInt32(Session["INVOICE_COUNT"]))
                        //    //{
                        //    //    ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('Customer Credit Invoice Limit is " + Session["INVOICE_COUNT"].ToString() + "');", true);
                        //    //    return;
                        //    //}
                        //}
                        #endregion
                    }

                    bool IsValidInsert = mOrderController.Update_Invoice(Convert.ToInt64(drpDocumentNo.SelectedValue), int.Parse(Session["DistributorId"].ToString()), mTownId, long.Parse(Session["AreaId"].ToString()), 0, long.Parse(Session["CUSTOMER_ID"].ToString()), long.Parse(Session["CUSTOMER_ID"].ToString()), int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), long.Parse(drpDocumentNo.SelectedValue.ToString()),
                    decimal.Parse(DC.chkNull_0(txtGrossAmount.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalStndrdDiscnt.Text)), decimal.Parse(DC.chkNull_0(numtxtTotalExtraDiscnt.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalGST.Text)), decimal.Parse(DC.chkNull_0(numTxtTotlAmnt.Text)), 0, int.Parse(RblPayMode.SelectedValue.ToString()),
                    PurchaseSKU, dtFreeSKU, int.Parse(Session["UserId"].ToString()), decimal.Parse(DC.chkNull_0(txtCashReceived.Text)), DateTime.Parse(Session["OrderDate"].ToString()), decimal.Parse(DC.chkNull_0(numTxtTotalTST.Text)), decimal.Parse(DC.chkNull_0(numtxtUnClaimabledist.Text)), txtRemarks.Text);
                    if (IsValidInsert)
                    {
                        btnSaveOrder.Text = "Save Invoice";
                        LoadCustomerData();
                        LoadPendingOrder();
                        ClearMasterALL();
                        Session.Remove("hfBillBookNo");
                        LoadSKUDetail();
                        ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
                    }
                }
                else if (btnSaveOrder.Text == "Sale Return")
                {
                    bool IsValidInsert = mOrderController.Add_SaleReturn(int.Parse(Session["DistributorId"].ToString()), mTownId, (long)Session["AreaId"], 0, long.Parse(Session["CUSTOMER_ID"].ToString()), long.Parse(Session["CUSTOMER_ID"].ToString()), int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), long.Parse(drpDocumentNo.SelectedValue.ToString()),
                    decimal.Parse(DC.chkNull_0(txtGrossAmount.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalStndrdDiscnt.Text)), decimal.Parse(DC.chkNull_0(numtxtTotalExtraDiscnt.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalGST.Text)), decimal.Parse(DC.chkNull_0(numTxtTotlAmnt.Text)), 0, int.Parse(RblPayMode.SelectedValue.ToString()), PurchaseSKU, dtFreeSKU, int.Parse(Session["UserId"].ToString()), DateTime.Parse(Session["CurrentWorkDate"].ToString()), decimal.Parse(DC.chkNull_0(numTxtTotalTST.Text)), decimal.Parse(DC.chkNull_0(numtxtUnClaimabledist.Text)), Convert.ToInt32(Session["UserID"]), txtRemarks.Text);
                    if (IsValidInsert)
                    {
                        LoadPendingOrder();
                        ClearMasterALL();
                        ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
                    }
                }
                else if (btnSaveOrder.Text == "Update Return")
                {
                    bool IsValidInsert = mOrderController.Update_SaleReturn(int.Parse(Session["DistributorId"].ToString()), mTownId, (long)Session["AreaId"], long.Parse(Session["CUSTOMER_ID"].ToString()), long.Parse(Session["CUSTOMER_ID"].ToString()), int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), long.Parse(drpDocumentNo.SelectedValue.ToString()),
                    decimal.Parse(DC.chkNull_0(txtGrossAmount.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalStndrdDiscnt.Text)), decimal.Parse(DC.chkNull_0(numtxtTotalExtraDiscnt.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalGST.Text)), decimal.Parse(DC.chkNull_0(numTxtTotlAmnt.Text)), 0, int.Parse(RblPayMode.SelectedValue.ToString()), PurchaseSKU, dtFreeSKU, int.Parse(Session["UserId"].ToString()), DateTime.Parse(Session["CurrentWorkDate"].ToString()), decimal.Parse(DC.chkNull_0(numTxtTotalTST.Text)), decimal.Parse(DC.chkNull_0(numtxtUnClaimabledist.Text)), Convert.ToInt32(Session["UserID"]), txtRemarks.Text);
                    if (IsValidInsert)
                    {
                        btnSaveOrder.Text = "Sale Return";
                        LoadPendingOrder();
                        ClearMasterALL();
                        ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
                    }
                }
            }
        }
    }
    protected void btnCancel_Click(object sender, EventArgs e)
    {
        Session.Remove("CUSTOMER_ID");
        Session.Remove("INVOICE_COUNT");
        Session.Remove("AreaId");
        Session.Remove("OrderBookerId");
        Session.Remove("DeliveryManId");
        Response.Redirect("~/Forms/frmOrderEntryStep1.aspx?Status=" + false + "&LevelType=3&LevelID=" + Request.QueryString["LevelID"].ToString());
    }

    #endregion

    #region Index/Change

    protected void drpDocumentNo_SelectedIndexChanged(object sender, EventArgs e)
    {
        int OrderNo = (int)Session["OrderNo"];
        if (long.Parse(drpDocumentNo.SelectedValue.ToString()) != Constants.LongNullValue)
        {
            txtBillBookNo.Text = string.Empty;
            DataTable dt = (DataTable)Session["dtOrder"];
            DataRow[] foundRows = dt.Select("SALE_INVOICE_ID  = '" + drpDocumentNo.SelectedValue + "'");
            if (foundRows.Length > 0)
            {
                FindCustomer();
                txtGrossAmount.Text = String.Format("{0:0.00}", foundRows[0]["TOTAL_AMOUNT"]);
                numTxtTotalStndrdDiscnt.Text = String.Format("{0:0.00}", foundRows[0]["EXTRA_DISCOUNT_AMOUNT"]);
                numtxtTotalExtraDiscnt.Text = String.Format("{0:0.00}", foundRows[0]["STANDARD_DISCOUNT_AMOUNT"]);
                numTxtTotalGST.Text = String.Format("{0:0.00}", foundRows[0]["GST_AMOUNT"]);
                numTxtTotlAmnt.Text = String.Format("{0:0.00}", foundRows[0]["TOTAL_NET_AMOUNT"]);
                numTxtTotalTST.Text = String.Format("{0:0.00}", foundRows[0]["TST_AMOUNT"]);
                drpCustomer.SelectedValue = foundRows[0]["CUSTOMER_ID"].ToString();
                numtxtUnClaimabledist.Text = String.Format("{0:0.00}", foundRows[0]["SED_AMOUNT"]);
                txtRemarks.Text = foundRows[0]["REMARKS"].ToString();

                if (OrderNo == -2)
                {
                    ExistenOrderDetail(long.Parse(drpDocumentNo.SelectedValue.ToString()), 2);
                    btnSaveOrder.Text = "Update Return";
                }
                else
                {
                    ExistenOrderDetail(long.Parse(drpDocumentNo.SelectedValue.ToString()), 0);
                    btnSaveOrder.Text = "Update Order";
                    RblPayMode.SelectedValue = foundRows[0]["ORDER_TYPE_ID"].ToString();
                    tdInvButton.Visible = true;
                }
                ClearAll();
                ScriptManager.GetCurrent(Page).SetFocus(ddlSKuCde);
                btnCalculate.Enabled = true;
            }
        }
        else
        {
            CreateFreeSKU();
            CreateTable();
            LoadGird();
            LoadFreeGrid();
            ClearMasterALL();
            ClearAll();
            if (OrderNo == -2)
            {
                btnSaveOrder.Text = "Sale Return";
            }
            else
            {
                btnSaveOrder.Text = "Save Order";
            }
            Session.Remove("hfBillBookNo");
        }
    }

    protected void ddlSKuCde_SelectedIndexChanged(object sender, EventArgs e)
    {
        try
        {
            txtStock.Text = "0-0";

            DataTable dtskuPrice = (DataTable)Session["Dtsku_Price"];

            DataRow[] foundRows = dtskuPrice.Select("SKU_ID  = '" + ddlSKuCde.SelectedValue + "'");
            txtStock.Text = (foundRows[0]["Stock"].ToString());
            txtUnitRate.Text = (foundRows[0]["TRADE_PRICE"].ToString());

            if (txtStock.Text == "")
            {
                txtStock.Text = "0-0";
            }


        }
        catch
        {
            txtStock.Text = "0-0";

            txtUnitRate.Text = "";
        }

    }

    protected void rbExtraDiscountType_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 1)
        {
            txtExtDiscount.Enabled = true;
            txtExtDiscountValue.Enabled = false;
            txtExtDiscount.Text = "";
            txtExtDiscountValue.Text = "";
        }
        else if (int.Parse(rbExtraDiscountType.SelectedValue.ToString()) == 2)
        {
            txtExtDiscount.Enabled = false;
            txtExtDiscountValue.Enabled = true;
            txtExtDiscount.Text = "";
            txtExtDiscountValue.Text = "";
        }
    }

    protected void ddlPrincipal_SelectedIndexChanged(object sender, EventArgs e)
    {
        this.LoadSKUDetail();
        this.LoadSKULastPromotionDiscount();
        ScriptManager.GetCurrent(Page).SetFocus(ddlPrincipal);
    }

    #endregion

    private bool IsFreeSKUinOrder(string SkUID)
    {
        bool flag = false;
        PurchaseSKU = (DataTable)Session["PurchaseSKU"];
        foreach (DataRow dr in PurchaseSKU.Rows)
        {
            if (dr["SKU_ID"].ToString() == SkUID)
            {
                flag = true;
                break;
            }
        }
        return flag;
    }

    private void ClearAll()
    {

        txtUnitRate.Text = "";
        txtExtDiscount.Text = "";
        txtExtDiscountValue.Text = "";
        txtFreeSku.Text = "";
        txtQuantity.Text = "";
        txtCtn.Text = "";
        txtStock.Text = "0-0";
        btnSave.Text = "Add";
        ddlSKuCde.Enabled = true;

        ddlSKuCde.Focus();
    }

    private void ClearMasterALL()
    {
        EnableDisableController(true);
        Session.Remove("PurchaseSKU");
        Session.Remove("dtFreeSKU");
        CreateTable();
        CreateFreeSKU();

        LoadGird();
        LoadFreeGrid();
        txtGrossAmount.Text = "";
        numtxtTotalExtraDiscnt.Text = "";
        numTxtTotalStndrdDiscnt.Text = "";
        numTxtTotalGST.Text = "";
        numTxtTotlAmnt.Text = "";
        numTxtTotalSED.Text = "";
        numTxtTotalTST.Text = "";
        numTxtTotalSED.Text = "";
        txtRemarks.Text = string.Empty;
        numtxtUnClaimabledist.Text = "";
        RowId = 0;
        mCustomerTypeId = 0;
        mCustomerVolClassId = 0;
        drpDocumentNo.SelectedIndex = 0;
        txtCashReceived.Text = "";
        txtBillBookNo.Text = string.Empty;
        btnSaveInvoice.Enabled = false;
        tdInvButton.Visible = false;

    }

    private bool IsDayClosed()
    {
        bool flag = false;
        DistributorController DistrCtl = new DistributorController();
        DataTable dtDayClose = DistrCtl.MaxDayClose(Convert.ToInt32(Session["DistributorId"]), 3);
        if (Convert.ToDateTime(Session["CurrentWorkDate"]) <= Convert.ToDateTime(dtDayClose.Rows[0]["DayClose"]))
        {
            flag = false;
        }
        else
        {
            flag = true;
        }

        return flag;
    }

    private bool FindCustomer()
    {
        DataTable dtCustomer = (DataTable)Session["dtCustomer"];
        DataRow[] foundRows = dtCustomer.Select("CUSTOMER_ID  = '" + drpCustomer.SelectedValue + "'");
        if (foundRows.Length > 0)
        {
            Session.Add("CUSTOMER_ID", long.Parse(foundRows[0]["CUSTOMER_ID"].ToString()));
            Session.Add("INVOICE_COUNT", int.Parse(foundRows[0]["INVOICE_COUNT"].ToString()));
            mCustomerTypeId = int.Parse(foundRows[0]["CHANNEL_TYPE_ID"].ToString());
            mCustomerVolClassId = int.Parse(foundRows[0]["VOLUME_CLASS_ID"].ToString());
            mTownId = int.Parse(foundRows[0]["TOWN_ID"].ToString());
            return true;
        }
        return false;
    }

    private void LoadSKULastPromotionDiscount()
    {
        try
        {
            SKUPriceDetailController PController = new SKUPriceDetailController();
            DataTable Last_Promotion = PController.SelectSkuLastPromotion(int.Parse(ddlPrincipal.SelectedValue), int.Parse(Session["DistributorId"].ToString()), int.Parse(drpCustomer.SelectedValue), int.Parse(Session["UserId"].ToString()), int.Parse(ddlSKuCde.SelectedValue), Constants.IntNullValue, DateTime.Parse(Session["CurrentWorkDate"].ToString()));
            if (Last_Promotion != null)
            {
                if (Last_Promotion.Rows.Count > 0)
                {
                    lblDiscountValue.Text = "";
                    lblDiscountTime.Text = "";
                    lblfreeskutime.Text = "";
                    lblfreeskuValue.Text = "";
                    lblQty.Text = "";
                    lblQtyTime.Text = "";

                    lblinvoice.Text = "";
                    lblInvoiceDate.Text = "";

                    lblDiscountValue.Text = Math.Round(decimal.Parse(dc.chkNull_0(Last_Promotion.Rows[0]["DISCOUNT"].ToString())), 2).ToString();
                    lblfreeskuValue.Text = Math.Round(decimal.Parse(dc.chkNull_0(Last_Promotion.Rows[0]["FREE_SKU"].ToString())), 0).ToString();
                    lblQty.Text = Math.Round(decimal.Parse(dc.chkNull_0(Last_Promotion.Rows[0]["QUANTITY"].ToString())), 0).ToString();
                    lblinvoice.Text = Last_Promotion.Rows[0]["SALE_INVOICE_ID"].ToString();


                    lblDiscountTime.Text = Convert.ToDateTime(Last_Promotion.Rows[0]["TIME_STAMP"].ToString()).ToString("dd-MMM-yyyy");
                    lblfreeskutime.Text = Convert.ToDateTime(Last_Promotion.Rows[0]["TIME_STAMP"].ToString()).ToString("dd-MMM-yyyy");
                    lblQtyTime.Text = Convert.ToDateTime(Last_Promotion.Rows[0]["TIME_STAMP"].ToString()).ToString("dd-MMM-yyyy");
                    lblInvoiceDate.Text = Convert.ToDateTime(Last_Promotion.Rows[0]["TIME_STAMP"].ToString()).ToString("dd-MMM-yyyy");


                }
                else
                {
                    lblDiscountValue.Text = "";
                    lblDiscountTime.Text = "";
                    lblfreeskutime.Text = "";
                    lblfreeskuValue.Text = "";
                    lblQty.Text = "";
                    lblQtyTime.Text = "";
                    lblinvoice.Text = "";
                    lblInvoiceDate.Text = "";
                }

            }
        }
        catch (Exception ex)
        {
            ExceptionPublisher.PublishException(ex);
        }

    }

    protected void drpCustomer_SelectedIndexChanged(object sender, EventArgs e)
    {
        LoadSKULastPromotionDiscount();
        ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
    }

    protected void ddlSKuCde_SelectedIndexChanged1(object sender, EventArgs e)
    {
        LoadSKULastPromotionDiscount();
        ScriptManager.GetCurrent(Page).SetFocus(ddlSKuCde);
    }

    protected void btnSaveInvoice_Click(object sender, EventArgs e)
    {
        if (FindCustomer())
        {
            if (IsDayClosed())
            {
                UserController UserCtl = new UserController();

                UserCtl.InsertUserLogoutTime(Convert.ToInt32(Session["User_Log_ID"]), Convert.ToInt32(Session["UserID"]));
                Session.Clear();
                System.Web.Security.FormsAuthentication.SignOut();
                Response.Redirect("../Login.aspx");
            }
            else
            {
                string ManualID = null;
                if (txtBillBookNo.Text.Trim().Length > 0)
                {
                    ManualID = txtBillBookNo.Text.ToUpper();
                }
                DataControl DC = new DataControl();
                PurchaseSKU = (DataTable)Session["PurchaseSKU"];
                dtFreeSKU = (DataTable)Session["dtFreeSKU"];

                OrderEntryController mOrderController = new OrderEntryController();
                #region Check Stock Limit
                foreach (DataRow dr in PurchaseSKU.Rows)
                {
                    PhaysicalStockController mController = new PhaysicalStockController();
                    DataTable dtstock = mController.SelectSKUClosingStock(int.Parse(Session["DistributorId"].ToString()), int.Parse(dr["SKU_ID"].ToString()), "", DateTime.Parse(Session["OrderDate"].ToString()));
                    if (dtstock.Rows.Count > 0)
                    {
                        int Total_Units = (int.Parse(dr["UNITS_IN_CASE"].ToString()) * int.Parse(dr["QUANTITY_CTN"].ToString())) + int.Parse(dr["QUANTITY_UNIT"].ToString()) + int.Parse(dr["FREE_SKU"].ToString());

                        if (int.Parse(dtstock.Rows[0][0].ToString()) < Total_Units)
                        {
                            ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + dr["SKU_Code"].ToString() + " Current Stock is " + dtstock.Rows[0][0].ToString() + "');", true);
                            return;
                        }
                        else if (Convert.ToDecimal(dtstock.Rows[0][1]) != Convert.ToDecimal(dr["UNIT_PRICE"]))
                        {
                            ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert('" + dr["SKU_Code"].ToString() + " Order Rate is " + dr["UNIT_PRICE"].ToString() + " and Current Rate is " + dtstock.Rows[0][1].ToString() + "');", true);
                            return;
                        }
                    }
                    else
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "msg", "alert(' + " + dr["SKU_Code"].ToString() + "No Stock Found');", true);
                        return;
                    }
                }
                #endregion

                bool IsValidInsert = mOrderController.Add_Invoice(int.Parse(Session["DistributorId"].ToString()), ManualID, mTownId, long.Parse(Session["AreaId"].ToString()), 0, long.Parse(Session["CUSTOMER_ID"].ToString()), long.Parse(Session["CUSTOMER_ID"].ToString()), int.Parse(Session["OrderBookerId"].ToString()), int.Parse(Session["DeliveryManId"].ToString()), long.Parse(drpDocumentNo.SelectedValue.ToString()),
                decimal.Parse(DC.chkNull_0(txtGrossAmount.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalStndrdDiscnt.Text)), decimal.Parse(DC.chkNull_0(numtxtTotalExtraDiscnt.Text)), decimal.Parse(DC.chkNull_0(numTxtTotalGST.Text)), decimal.Parse(DC.chkNull_0(numTxtTotlAmnt.Text)), 0, int.Parse(RblPayMode.SelectedValue.ToString()),
                PurchaseSKU, dtFreeSKU, int.Parse(Session["UserId"].ToString()), decimal.Parse(DC.chkNull_0(txtCashReceived.Text)), DateTime.Parse(Session["OrderDate"].ToString()), decimal.Parse(DC.chkNull_0(numTxtTotalTST.Text)), decimal.Parse(DC.chkNull_0(numtxtUnClaimabledist.Text)), txtRemarks.Text);
                if (IsValidInsert)
                {
                    LoadCustomerData();
                    LoadPendingOrder();
                    ClearMasterALL();
                    Session.Remove("hfBillBookNo");
                    LoadSKUDetail();
                    drpDocumentNo_SelectedIndexChanged(null, null);
                    ScriptManager.GetCurrent(Page).SetFocus(drpCustomer);
                }
            }
        }
    }
}