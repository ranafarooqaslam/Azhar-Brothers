using System;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using SAMSBusinessLayer.Classes;
using SAMSCommon.Classes;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Web;
using System.Drawing;
using System.Linq;

/// <summary>
/// Form For Stock Reconciliation Report
/// </summary>
public partial class Forms_RptStockReconcilation : System.Web.UI.Page
{
    /// <summary>
    /// Page_Load Function
    /// </summary>
    /// <param name="sender">object</param>
    /// <param name="e">EventArgs</param>
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            this.LoadDistributor();
            this.LoadPrincipal();
            SAMSCommon.Classes.Configuration.SystemCurrentDateTime = (DateTime)this.Session["CurrentWorkDate"];
            txtStartDate.Text = SAMSCommon.Classes.Configuration.SystemCurrentDateTime.ToString("dd-MMM-yyyy");
            txtEndDate.Text = SAMSCommon.Classes.Configuration.SystemCurrentDateTime.ToString("dd-MMM-yyyy");
        }
    }

    /// <summary>
    /// Loads Principals To Principal Combo
    /// </summary>
    private void LoadPrincipal()
    {
        SKUPriceDetailController PController = new SKUPriceDetailController();
        DataTable m_dt = PController.SelectDataPrice(Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, int.Parse(this.Session["UserId"].ToString()), Constants.IntNullValue, 0, DateTime.Parse(this.Session["CurrentWorkDate"].ToString()));        
        clsWebFormUtil.FillListBox(this.cblPrincipal, m_dt, 0, 1);

        foreach (ListItem li in cblPrincipal.Items)
        {
            li.Selected = true;
        }

        cblPrincipal.Attributes.Add("onclick", string.Format("javascript:CheckBoxListSelect({0})", cblPrincipal.ClientID));
    }

    /// <summary>
    /// Loads Locations To Location Combo
    /// </summary>
    private void LoadDistributor()
    {
        DistributorController DController = new DistributorController();
        DataTable dt = DController.SelectDistributorInfo(Constants.IntNullValue, int.Parse(this.Session["UserId"].ToString()), int.Parse(this.Session["CompanyId"].ToString()));
       // drpDistributor.Items.Add(new ListItem("All", Constants.IntNullValue.ToString()));       
        clsWebFormUtil.FillDropDownList(this.drpDistributor, dt, 0, 2);
    }

    /// <summary>
    /// Shows Stock Reconciliation in PDF
    /// </summary>
    /// <param name="sender">object</param>
    /// <param name="e">EventArgs</param>
    protected void btnViewPDF_Click(object sender, EventArgs e)
    {
        string PrincipalIDs = null;
        string PrincipalNames = "";
        foreach (ListItem li in cblPrincipal.Items)
        {
            if (li.Selected)
            {
                PrincipalIDs += li.Value + ",";
                PrincipalNames += li.Text + ", ";
            }
        }

        PrincipalIDs = PrincipalIDs.Substring(0, PrincipalIDs.Length - 1);
        PrincipalNames = PrincipalNames.Substring(0, PrincipalNames.Length - 1);
        DocumentPrintController mController = new DocumentPrintController();
        RptInventoryController RptInventoryCtl = new RptInventoryController();
        SAMSBusinessLayer.Reports.StockReConsilation CrpReport = new SAMSBusinessLayer.Reports.StockReConsilation();
        DataTable dt = mController.SelectReportTitle(int.Parse(drpDistributor.SelectedValue.ToString()));
        DataSet ds = RptInventoryCtl.SelectPrincipalStockReconcilation(int.Parse(drpDistributor.SelectedValue.ToString()), PrincipalIDs, DateTime.Parse(txtStartDate.Text), DateTime.Parse(txtEndDate.Text), int.Parse(this.Session["UserId"].ToString()),ddlType.SelectedIndex,Convert.ToInt32(rblRate.SelectedValue));
        CrpReport.SetDataSource(ds);
        CrpReport.Refresh();


        CrpReport.SetParameterValue("division", drpDistributor.SelectedItem.Text);
        CrpReport.SetParameterValue("Principal", PrincipalNames);
        CrpReport.SetParameterValue("fromdate", txtStartDate.Text);
        CrpReport.SetParameterValue("todate", txtEndDate.Text );
        CrpReport.SetParameterValue("CompanyName", dt.Rows[0]["COMPANY_NAME"].ToString());
        CrpReport.SetParameterValue("Price", rblRate.SelectedItem.Text);
        CrpReport.SetParameterValue("ZeroElimination", cbZero.Checked);
        CrpReport.SetParameterValue("ReportType", "Stock Reconciliation Report ( " + ddlType.SelectedItem.Text + " )");

        this.Session.Add("CrpReport", CrpReport);
        this.Session.Add("ReportType", 0);  
        string url = "'Default.aspx'";
        string script = "<script language='JavaScript' type='text/javascript'> window.open(" + url + ",\"Link\",\"toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,width=800,height=600,left=10,top=10\");</script>";
        Type cstype = this.GetType();
        ClientScriptManager cs = Page.ClientScript;
        cs.RegisterStartupScript(cstype, "OpenWindow", script);            
    }

    /// <summary>
    /// Shows Stock Reconciliation in Excel
    /// </summary>
    /// <param name="sender">object</param>
    /// <param name="e">EventArgs</param>
    protected void btnViewExcel_Click(object sender, EventArgs e)
    {
        string PrincipalIDs = null;
        string PrincipalNames = "";
        foreach (ListItem li in cblPrincipal.Items)
        {
            if (li.Selected)
            {
                PrincipalIDs += li.Value + ",";
                PrincipalNames += li.Text + ", ";
            }
        }
        PrincipalIDs = PrincipalIDs.Substring(0, PrincipalIDs.Length - 1);
        PrincipalNames = PrincipalNames.Substring(0, PrincipalNames.Length - 1);
        DocumentPrintController mController = new DocumentPrintController();
        RptInventoryController RptInventoryCtl = new RptInventoryController();
        DataTable dt = mController.SelectReportTitle(int.Parse(drpDistributor.SelectedValue.ToString()));
        DataSet ds = RptInventoryCtl.SelectPrincipalStockReconcilation(int.Parse(drpDistributor.SelectedValue.ToString()), PrincipalIDs, DateTime.Parse(txtStartDate.Text), DateTime.Parse(txtEndDate.Text), int.Parse(this.Session["UserId"].ToString()), ddlType.SelectedIndex, Convert.ToInt32(rblRate.SelectedValue));

        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        using (ExcelPackage p = new ExcelPackage())
        {
            ExcelWorksheet ws = p.Workbook.Worksheets.Add("StockReconcilliationReport");

            var headerCells = ws.Cells[1, 1, 2, ExcelPackage.MaxColumns];
            var headerFont = headerCells.Style.Font;
            headerFont.Bold = true;

            string[] strColumnNames = { "SKU Code", "SKU Name", "Opening", "+Purchas", "-Purchase Return", "-Sales", "+Sales  Return", "+ Transfer In", "- Transfer Out","Free StockIn","Free StockOu", "Shortage", "+Excess", "Damage","Rate Difference","Closing Stock", "Clsoing Amount","Stock availfor Days" };

            ExportToExcel(ws, p, dt.Rows[0]["COMPANY_NAME"].ToString(), "Stock Reconciliation Report ( " + ddlType.SelectedItem.Text + " )", "Location: " + this.drpDistributor.SelectedItem.Text.ToString(), "Principal: " + PrincipalNames,Convert.ToDateTime(this.txtStartDate.Text), Convert.ToDateTime(this.txtEndDate.Text), strColumnNames, ds.Tables["StockRegister"]);

            Byte[] fileBytes = p.GetAsByteArray();
            Response.Clear();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment;filename=StockReconcilliationReport.xlsx");
            Response.Charset = "";
            Response.ContentType = "application/vnd.ms-excel";
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.BinaryWrite(fileBytes);
            Response.End();
        }
    }
    public static void ExportToExcel(ExcelWorksheet ws, ExcelPackage p, string strCompanyName, string strHeader, string strLocation, string strPrincipal, DateTime dtFrom, DateTime dtTo, string[] ColumnNames, DataTable dbStockRegister)
    {
        int rowIndex = 1;

        if (dbStockRegister.Rows.Count > 0)
        {
            ws.Cells[rowIndex, 1].Value = strCompanyName;

            rowIndex++;
            ws.Cells[rowIndex, 1].Value = strHeader;

            rowIndex++;
            ws.Cells[rowIndex, 1].Value = strLocation;

            rowIndex++;
            ws.Cells[rowIndex, 1].Value = strPrincipal;

            rowIndex++;

            ws.Cells[rowIndex, 1].Value = "From Date:";
            ws.Cells[rowIndex, 3].Value = dtFrom;
            ws.Cells[rowIndex, 3].Style.Numberformat.Format = "dd-MMM-yyyy";

            ws.Cells[rowIndex, 6].Value = "To Date:";
            ws.Cells[rowIndex, 8].Value = dtTo;
            ws.Cells[rowIndex, 8].Style.Numberformat.Format = "dd-MMM-yyyy";

            rowIndex++;
            int ColumnIndex = 0;
            foreach (var col in ColumnNames)
            {
                ColumnIndex++;
                ws.Cells[rowIndex, ColumnIndex].Value = col;
                ws.Cells[rowIndex, ColumnIndex].Style.Font.Bold = true;
                ws.Cells[rowIndex, ColumnIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }
            
            decimal openingStockValue = 0;
            decimal purchasedValue = 0;
            decimal purchaseReturnValue = 0;
            decimal saledValue = 0;
            decimal saleReturnValue = 0;
            decimal transferInValue = 0;
            decimal transferOutValue = 0;
            decimal freeStockIn = 0;
            decimal freeStockOut = 0;
            decimal shortValue = 0;
            decimal excessValue = 0;
            decimal damageStockValue = 0;
            decimal sumOfRateDiff = 0;
            decimal sumOfClosingStock = 0;
            decimal sumOfClosingAmount = 0;

            var groupedData = dbStockRegister.AsEnumerable()
            .GroupBy(row => new
            {
                HieName = row["SKU_HIE_NAME"].ToString(),
                HieCode = row["SKU_HIE_CODE"]
            })
            .OrderBy(g => g.Key.HieCode);

            decimal closingStock = 0;
            decimal closingAmount = 0;
            decimal stockPerDay = 0;
            foreach (var group in groupedData)
            {
                rowIndex++;
                ws.Cells[rowIndex, 1].Value = group.Key.HieName;
                ws.Cells[rowIndex, 1, rowIndex, ColumnNames.Length].Merge = true;
                ws.Cells[rowIndex, 1].Style.Font.Bold = true;
                ws.Cells[rowIndex, 1].Style.Fill.PatternType =
                    OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[rowIndex, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                ws.Cells[rowIndex, 1].Style.Border.BorderAround(
                    ExcelBorderStyle.Thin, Color.Black);

                closingStock = 0;
                closingAmount = 0;
                stockPerDay = 0;
                // Items under this group
                foreach (DataRow DataTableRow in group)
                {
                    rowIndex++;
                    closingStock = 0;
                    closingStock = Convert.ToDecimal(DataTableRow["OPENING_STOCK"])
                        + Convert.ToDecimal(DataTableRow["PURCHASED"])
                    + Convert.ToDecimal(DataTableRow["SALE_RETURN"])
                    + Convert.ToDecimal(DataTableRow["EXCESS"])
                    + Convert.ToDecimal(DataTableRow["TRANSFER_IN"])
                    + Convert.ToDecimal(DataTableRow["FREESTOCK_IN"])
                    - Convert.ToDecimal(DataTableRow["SALED"])
                    - Convert.ToDecimal(DataTableRow["PURCHASE_RETURN"])
                    - Convert.ToDecimal(DataTableRow["DAMAGED_STOCK"])
                    - Convert.ToDecimal(DataTableRow["TRANSFER_OUT"])
                    - Convert.ToDecimal(DataTableRow["FREESTOCK_OUT"])
                    - Convert.ToDecimal(DataTableRow["SHORT"]);

                    closingAmount = Convert.ToDecimal(DataTableRow["OPENING_STOCK_VALUE"])
                    + Convert.ToDecimal(DataTableRow["PURCHASED_VALUE"])
                    + Convert.ToDecimal(DataTableRow["SALE_RETURN_VALUE"])
                    + Convert.ToDecimal(DataTableRow["TRANSFER_IN_VALUE"])
                    + Convert.ToDecimal(DataTableRow["EXCESS_VALUE"])
                    - Convert.ToDecimal(DataTableRow["TRANSFER_OUT_VALUE"])
                    - Convert.ToDecimal(DataTableRow["SHORT_VALUE"])
                    - Convert.ToDecimal(DataTableRow["SALED_VALUE"])
                    - Convert.ToDecimal(DataTableRow["PURCHASE_RETURN_VALUE"])
                    - Convert.ToDecimal(DataTableRow["FREESTOCK_OUT_VALUE"])
                    - Convert.ToDecimal(DataTableRow["DAMAGED_STOCK_VALUE"])
                    + Convert.ToDecimal(DataTableRow["RATEDIEFFER"]);

                    if(Convert.ToDecimal(DataTableRow["SALED"]) == 0)
                    {
                        stockPerDay = closingStock * Convert.ToDecimal(DataTableRow["SKU_ID"]);
                    }
                    else
                    {
                        stockPerDay = closingStock * Convert.ToDecimal(DataTableRow["SKU_ID"]) / Convert.ToDecimal(DataTableRow["SALED"]);
                    }

                    ws.Cells[rowIndex, 1].Value = DataTableRow["SKU_CODE"];
                    ws.Cells[rowIndex, 2].Value = DataTableRow["SKU_NAME"];
                    ws.Cells[rowIndex, 3].Value = DataTableRow["OPENING_STOCK"];
                    ws.Cells[rowIndex, 4].Value = DataTableRow["PURCHASED"];
                    ws.Cells[rowIndex, 5].Value = DataTableRow["PURCHASE_RETURN"];
                    ws.Cells[rowIndex, 6].Value = DataTableRow["SALED"];
                    ws.Cells[rowIndex, 7].Value = DataTableRow["SALE_RETURN"];
                    ws.Cells[rowIndex, 8].Value = DataTableRow["TRANSFER_IN"];
                    ws.Cells[rowIndex, 9].Value = DataTableRow["TRANSFER_OUT"];
                    ws.Cells[rowIndex, 10].Value = DataTableRow["FREESTOCK_IN"];
                    ws.Cells[rowIndex, 11].Value = DataTableRow["FREESTOCK_OUT"];
                    ws.Cells[rowIndex, 12].Value = DataTableRow["SHORT"];
                    ws.Cells[rowIndex, 13].Value = DataTableRow["EXCESS"];
                    ws.Cells[rowIndex, 14].Value = DataTableRow["DAMAGED_STOCK"];
                    ws.Cells[rowIndex, 15].Value = DataTableRow["RATEDIEFFER"];
                    ws.Cells[rowIndex, 16].Value = closingStock;
                    ws.Cells[rowIndex, 17].Value = closingAmount;
                    ws.Cells[rowIndex, 18].Value = stockPerDay;

                    ws.Cells[rowIndex, 1, rowIndex, ColumnNames.Length]
                        .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                    openingStockValue += Convert.ToDecimal(DataTableRow["OPENING_STOCK_VALUE"]);
                    purchasedValue += Convert.ToDecimal(DataTableRow["PURCHASED_VALUE"]);
                    purchaseReturnValue += Convert.ToDecimal(DataTableRow["PURCHASE_RETURN_VALUE"]);
                    saledValue += Convert.ToDecimal(DataTableRow["SALED_VALUE"]);
                    saleReturnValue += Convert.ToDecimal(DataTableRow["SALE_RETURN_VALUE"]);
                    transferInValue += Convert.ToDecimal(DataTableRow["TRANSFER_IN_VALUE"]);
                    transferOutValue+= Convert.ToDecimal(DataTableRow["TRANSFER_OUT_VALUE"]);
                    freeStockIn += Convert.ToDecimal(DataTableRow["FREESTOCK_IN_VALUE"]);
                    freeStockOut += Convert.ToDecimal(DataTableRow["FREESTOCK_OUT_VALUE"]);
                    shortValue += Convert.ToDecimal(DataTableRow["SHORT_VALUE"]);
                    excessValue += Convert.ToDecimal(DataTableRow["EXCESS_VALUE"]);
                    damageStockValue += Convert.ToDecimal(DataTableRow["DAMAGED_STOCK_VALUE"]);
                    sumOfRateDiff += Convert.ToDecimal(DataTableRow["RATEDIEFFER"]);
                    sumOfClosingStock += closingStock;
                    sumOfClosingAmount += closingAmount;
                }
                rowIndex++;
                ws.Cells[rowIndex, 3].Value = openingStockValue;
                ws.Cells[rowIndex, 3].Style.Font.Bold = true;
                ws.Cells[rowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 4].Value = purchasedValue;
                ws.Cells[rowIndex, 4].Style.Font.Bold = true;
                ws.Cells[rowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 5].Value = purchaseReturnValue;
                ws.Cells[rowIndex, 5].Style.Font.Bold = true;
                ws.Cells[rowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 6].Value = saledValue;
                ws.Cells[rowIndex, 6].Style.Font.Bold = true;
                ws.Cells[rowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 7].Value = saleReturnValue;
                ws.Cells[rowIndex, 7].Style.Font.Bold = true;
                ws.Cells[rowIndex, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 8].Value = transferInValue;
                ws.Cells[rowIndex, 8].Style.Font.Bold = true;
                ws.Cells[rowIndex, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 9].Value = transferOutValue;
                ws.Cells[rowIndex, 9].Style.Font.Bold = true;
                ws.Cells[rowIndex, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 10].Value = freeStockIn;
                ws.Cells[rowIndex, 10].Style.Font.Bold = true;
                ws.Cells[rowIndex, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 11].Value = freeStockOut;
                ws.Cells[rowIndex, 11].Style.Font.Bold = true;
                ws.Cells[rowIndex, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 12].Value = shortValue;
                ws.Cells[rowIndex, 12].Style.Font.Bold = true;
                ws.Cells[rowIndex, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 13].Value = excessValue;
                ws.Cells[rowIndex, 13].Style.Font.Bold = true;
                ws.Cells[rowIndex, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 14].Value = damageStockValue;
                ws.Cells[rowIndex, 14].Style.Font.Bold = true;
                ws.Cells[rowIndex, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 15].Value = sumOfRateDiff;
                ws.Cells[rowIndex, 15].Style.Font.Bold = true;
                ws.Cells[rowIndex, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 16].Value = sumOfClosingStock;
                ws.Cells[rowIndex, 16].Style.Font.Bold = true;
                ws.Cells[rowIndex, 16].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[rowIndex, 17].Value = sumOfClosingAmount;
                ws.Cells[rowIndex, 17].Style.Font.Bold = true;
                ws.Cells[rowIndex, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }

            
            ws.Cells[rowIndex, 3].Value = openingStockValue;
            ws.Cells[rowIndex, 3].Style.Font.Bold = true;
            ws.Cells[rowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 4].Value = purchasedValue;
            ws.Cells[rowIndex, 4].Style.Font.Bold = true;
            ws.Cells[rowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 5].Value = purchaseReturnValue;
            ws.Cells[rowIndex, 5].Style.Font.Bold = true;
            ws.Cells[rowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 6].Value = saledValue;
            ws.Cells[rowIndex, 6].Style.Font.Bold = true;
            ws.Cells[rowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 7].Value = saleReturnValue;
            ws.Cells[rowIndex, 7].Style.Font.Bold = true;
            ws.Cells[rowIndex, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 8].Value = transferInValue;
            ws.Cells[rowIndex, 8].Style.Font.Bold = true;
            ws.Cells[rowIndex, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 9].Value = transferOutValue;
            ws.Cells[rowIndex, 9].Style.Font.Bold = true;
            ws.Cells[rowIndex, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 10].Value = freeStockIn;
            ws.Cells[rowIndex, 10].Style.Font.Bold = true;
            ws.Cells[rowIndex, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 11].Value = freeStockOut;
            ws.Cells[rowIndex, 11].Style.Font.Bold = true;
            ws.Cells[rowIndex, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 12].Value = shortValue;
            ws.Cells[rowIndex, 12].Style.Font.Bold = true;
            ws.Cells[rowIndex, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 13].Value = excessValue;
            ws.Cells[rowIndex, 13].Style.Font.Bold = true;
            ws.Cells[rowIndex, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 14].Value = damageStockValue;
            ws.Cells[rowIndex, 14].Style.Font.Bold = true;
            ws.Cells[rowIndex, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 15].Value = sumOfRateDiff;
            ws.Cells[rowIndex, 15].Style.Font.Bold = true;
            ws.Cells[rowIndex, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 16].Value = sumOfClosingStock;
            ws.Cells[rowIndex, 16].Style.Font.Bold = true;
            ws.Cells[rowIndex, 16].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 17].Value = sumOfClosingAmount;
            ws.Cells[rowIndex, 17].Style.Font.Bold = true;
            ws.Cells[rowIndex, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        }
        rowIndex++;

        ws.Column(2).AutoFit();
        ws.Column(3).AutoFit();
    }
}
