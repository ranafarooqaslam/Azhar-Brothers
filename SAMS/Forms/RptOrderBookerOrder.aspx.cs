using System;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using SAMSBusinessLayer.Classes;
using SAMSCommon.Classes;
using SAMSBusinessLayer.Reports;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Web;
using System.Linq;
using System.Drawing;

/// <summary>
/// Form For Order Booker Reports
/// </summary>
public partial class Forms_RptOrderBookerOrder : System.Web.UI.Page
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
            LoadPrincipal();
            LoadLocation();
            LoadOrderBooker();
            LoadRoute();
            SAMSCommon.Classes.Configuration.SystemCurrentDateTime = (DateTime)this.Session["CurrentWorkDate"];
            this.txtFromDate.Text = SAMSCommon.Classes.Configuration.SystemCurrentDateTime.ToString("dd-MMM-yyyy");
            this.txtToDate.Text = SAMSCommon.Classes.Configuration.SystemCurrentDateTime.ToString("dd-MMM-yyyy");
        }
    }

    private void LoadRoute()
    {
        if (DrpLocation.Items.Count > 0)
        {
            DrpRoute.Items.Clear();
            DistributorAreaController mController = new DistributorAreaController();
            DataTable dt = mController.SelectDist_Area(Constants.LongNullValue, Constants.DateNullValue, Constants.DateNullValue, int.Parse(DrpLocation.SelectedValue.ToString()), Constants.IntNullValue, null, null);
            DrpRoute.Items.Add(new ListItem("All", Constants.IntNullValue.ToString()));
            clsWebFormUtil.FillDropDownList(DrpRoute, dt, 0, 6);
        }
        else
        {
            DrpRoute.Items.Clear();
        }
    }
    protected void LoadPrincipal()
    {
        try
        {
            SKUPriceDetailController PController = new SKUPriceDetailController();
            DataTable m_dt = PController.SelectDataPrice(Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, Constants.IntNullValue, int.Parse(this.Session["UserId"].ToString()), Constants.IntNullValue, 0, DateTime.Parse(this.Session["CurrentWorkDate"].ToString()));
            this.drpPrincipal.Items.Add(new ListItem("All", Constants.IntNullValue.ToString()));
            clsWebFormUtil.FillDropDownList(this.drpPrincipal, m_dt, "Company_Id", "Company_Name");
        }
        catch (Exception ex)
        {
            ex.ToString();
        }
    }

    /// <summary>
    /// Loads OrderBookers To OrderBooker Combo
    /// </summary>
    private void LoadOrderBooker()
    {
        DrpOrderBooker.Items.Clear();

        if (DrpLocation.Items.Count > 0)
        {
            
                SaleForceController mDController = new SaleForceController();
                DataTable m_dt = mDController.SelectSaleForceAssignedArea(Constants.SALES_FORCE_ORDERBOOKER, int.Parse(DrpLocation.SelectedValue.ToString()), Constants.IntNullValue, int.Parse(this.Session["CompanyId"].ToString()));
                DrpOrderBooker.Items.Add(new ListItem("All", Constants.IntNullValue.ToString()));
                clsWebFormUtil.FillDropDownList(this.DrpOrderBooker, m_dt, 0, 3);
          
        }
        
    }

    /// <summary>
    /// Loads Locations To Location Combo
    /// </summary>
    protected void LoadLocation()
    {
        try
        {
            DistributorController DController = new DistributorController();
            DataTable dt = DController.SelectDistributorInfo(Constants.IntNullValue, int.Parse(this.Session["UserId"].ToString()), int.Parse(this.Session["CompanyId"].ToString()));
            this.DrpLocation.DataSource = dt;
            this.DrpLocation.DataTextField = "DISTRIBUTOR_NAME";
            this.DrpLocation.DataValueField = "DISTRIBUTOR_ID";
            this.DrpLocation.DataBind();
        }
        catch (Exception ex)
        {
            ex.ToString();
        }
    }

    /// <summary>
    /// Shows Order Booker Reports in PDF
    /// </summary>
    /// <param name="sender">object</param>
    /// <param name="e">EventArgs</param>
    protected void btnViewPDF_Click(object sender, EventArgs e)
    {
        PrintReport(0);
    }

    /// <summary>
    /// Shows Order Booker Reports in Excel
    /// </summary>
    /// <param name="sender">object</param>
    /// <param name="e">EventArgs</param>
    protected void btnViewExcel_Click(object sender, EventArgs e)
    {
        SAMSBusinessLayer.Classes.DocumentPrintController DPrint = new SAMSBusinessLayer.Classes.DocumentPrintController();
        RptSaleController RptSaleCtl = new RptSaleController();
        DataTable dt = DPrint.SelectReportTitle(int.Parse(DrpLocation.SelectedValue.ToString()));

        string FromDate = null;
        string ToDate = null;

        DataSet ds = null;

        DateTime parsed_date_fromdate = DateTime.Parse(this.txtFromDate.Text);
        DateTime parsed_date_todate = DateTime.Parse(this.txtToDate.Text);
        FromDate = parsed_date_fromdate.ToShortDateString();
        ToDate = parsed_date_todate.ToShortDateString();

        ds = RptSaleCtl.OrderBookerOrderTakingSheet(int.Parse(drpPrincipal.SelectedValue.ToString()), int.Parse(DrpLocation.SelectedValue.ToString()), Convert.ToDateTime(FromDate + " 00:00:00"), Convert.ToDateTime(ToDate + " 00:00:00"), int.Parse(DrpOrderBooker.SelectedValue.ToString()), int.Parse(this.Session["UserId"].ToString()), int.Parse(DrpRoute.SelectedValue));
        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        using (ExcelPackage p = new ExcelPackage())
        {
            ExcelWorksheet ws = p.Workbook.Worksheets.Add("OBOrderTakingReport");

            var headerCells = ws.Cells[1, 1, 2, ExcelPackage.MaxColumns];
            var headerFont = headerCells.Style.Font;
            headerFont.Bold = true;          

            string[] strColumnNames = { "Sr No", "Customer", "Address", "Order No", "Order Gross Amount", "Order Disc. Amount", "Order Net Amount", "Remarks" };

            ExportToExcel(ws, p, dt.Rows[0]["COMPANY_NAME"].ToString(), "Orderbooker Order Taking Report", "Location: " + this.DrpLocation.SelectedItem.Text.ToString(), "Principal: " + drpPrincipal.SelectedItem.Text, "Order Booker: " + DrpOrderBooker.SelectedItem.Text,"Area: " + DrpRoute.SelectedItem.Text, Convert.ToDateTime(this.txtFromDate.Text), Convert.ToDateTime(this.txtToDate.Text), strColumnNames, ds.Tables["RptOrderBookerSheet"]);

            Byte[] fileBytes = p.GetAsByteArray();
            Response.Clear();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment;filename=OBOrderTakingReport.xlsx");
            Response.Charset = "";
            Response.ContentType = "application/vnd.ms-excel";
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.BinaryWrite(fileBytes);
            Response.End();
        }
    }

    public static void ExportToExcel(ExcelWorksheet ws, ExcelPackage p,string  strCompanyName, string strHeader, string strLocation,string strPrincipal,string strOrderbooker, string strArea, DateTime dtFrom, DateTime dtTo, string[] ColumnNames, DataTable RptOrderBookerSheet)
    {
        int rowIndex = 1;

        if (RptOrderBookerSheet.Rows.Count > 0)
        {
            ws.Cells[rowIndex, 1].Value = strCompanyName;

            rowIndex++;
            ws.Cells[rowIndex, 1].Value = strHeader;

            rowIndex++;
            ws.Cells[rowIndex, 1].Value = strLocation;

            rowIndex++;
            ws.Cells[rowIndex, 1].Value = strPrincipal;

            rowIndex++;
            ws.Cells[rowIndex, 1].Value = strOrderbooker;

            ws.Cells[rowIndex, 7].Value = "From Date:";
            ws.Cells[rowIndex, 8].Value = dtFrom;
            ws.Cells[rowIndex, 8].Style.Numberformat.Format = "dd-MMM-yyyy";

            rowIndex++;
            ws.Cells[rowIndex, 1].Value = strArea;
            ws.Cells[rowIndex, 7].Value = "To Date:";
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
            decimal GrossAmount = 0;
            decimal Discount = 0;
            decimal NetAmount = 0;

            decimal TotalGrossAmount = 0;
            decimal TotalDiscount = 0;
            decimal TotalNetAmount = 0;
            
            int SNo = 0;
            foreach (DataRow DataTableRow in RptOrderBookerSheet.Rows)
            {
                SNo++;
                rowIndex++;
                foreach (DataColumn DataTableColumn in RptOrderBookerSheet.Columns)
                {
                    var cell0 = ws.Cells[rowIndex, 1];
                    cell0.Value = SNo;
                    cell0.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    cell0.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    if (DataTableColumn.ColumnName == "CUSTOMER_CODE")
                    {
                        var cell = ws.Cells[rowIndex, 2];
                        cell.Value = DataTableRow[DataTableColumn.ColumnName] + "-" + DataTableRow["CUSTOMER_NAME"];
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    }
                    else if (DataTableColumn.ColumnName == "sku_name")
                    {
                        var cell = ws.Cells[rowIndex, 3];
                        cell.Value = DataTableRow[DataTableColumn.ColumnName];
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    }
                    else if (DataTableColumn.ColumnName == "PACKSIZE")
                    {
                        var cell = ws.Cells[rowIndex, 4];
                        cell.Value = DataTableRow[DataTableColumn.ColumnName];
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    }
                    else if (DataTableColumn.ColumnName == "QUANTITY_UNIT")
                    {
                        var cell = ws.Cells[rowIndex, 5];
                        cell.Value = DataTableRow[DataTableColumn.ColumnName];
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    }
                    else if (DataTableColumn.ColumnName == "AMOUNT")
                    {
                        var cell = ws.Cells[rowIndex, 6];
                        cell.Value = DataTableRow[DataTableColumn.ColumnName];
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    }
                    var cell7 = ws.Cells[rowIndex, 7];
                    cell7.Value = Convert.ToDecimal(DataTableRow["QUANTITY_UNIT"]) - Convert.ToDecimal(DataTableRow["AMOUNT"]);
                    cell7.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                    try
                    {
                        GrossAmount = Convert.ToDecimal(DataTableRow["QUANTITY_UNIT"]);
                    }
                    catch (Exception)
                    {
                        GrossAmount = 0;
                    }
                    try
                    {
                        Discount = Convert.ToDecimal(DataTableRow["AMOUNT"]);
                    }
                    catch (Exception)
                    {
                        Discount = 0;
                    }
                    try
                    {
                        NetAmount = Convert.ToDecimal(DataTableRow["QUANTITY_UNIT"]) - Convert.ToDecimal(DataTableRow["AMOUNT"]);
                    }
                    catch (Exception)
                    {
                        NetAmount = 0;
                    }
                }
                    TotalGrossAmount += GrossAmount;
                    TotalDiscount += Discount;
                TotalNetAmount += NetAmount;
            }

            rowIndex++;
            ws.Cells[rowIndex, 4].Value = "Total";
            ws.Cells[rowIndex, 4].Style.Font.Bold = true;
            ws.Cells[rowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 5].Value = TotalGrossAmount;
            ws.Cells[rowIndex, 5].Style.Font.Bold = true;
            ws.Cells[rowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 6].Value = TotalDiscount;
            ws.Cells[rowIndex, 6].Style.Font.Bold = true;
            ws.Cells[rowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells[rowIndex, 7].Value = TotalNetAmount;
            ws.Cells[rowIndex, 7].Style.Font.Bold = true;
            ws.Cells[rowIndex, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        }
        rowIndex++;

        ws.Column(2).AutoFit();
        ws.Column(3).AutoFit();
    }

    /// <summary>
    /// Loads OrderBookers
    /// </summary>
    protected void DrpLocation_SelectedIndexChanged(object sender, EventArgs e)
    {
        this.LoadOrderBooker();
        LoadRoute();
    }
    
    protected void PrintReport(int Type)
    {
        try
        {
            SAMSBusinessLayer.Classes.DocumentPrintController DPrint = new SAMSBusinessLayer.Classes.DocumentPrintController();
            RptSaleController RptSaleCtl = new RptSaleController();
            DataTable dt = DPrint.SelectReportTitle(int.Parse(DrpLocation.SelectedValue.ToString()));

            string FromDate = null;
            string ToDate = null;

            DataSet ds = null;

            DateTime parsed_date_fromdate = DateTime.Parse(this.txtFromDate.Text);
            DateTime parsed_date_todate = DateTime.Parse(this.txtToDate.Text);
            FromDate = parsed_date_fromdate.ToShortDateString();
            ToDate = parsed_date_todate.ToShortDateString();

            ds = RptSaleCtl.OrderBookerOrderTakingSheet(int.Parse(drpPrincipal.SelectedValue.ToString()), int.Parse(DrpLocation.SelectedValue.ToString()), Convert.ToDateTime(FromDate + " 00:00:00"), Convert.ToDateTime(ToDate + " 00:00:00"), int.Parse(DrpOrderBooker.SelectedValue.ToString()), int.Parse(this.Session["UserId"].ToString()), int.Parse(DrpRoute.SelectedValue));
            SAMSBusinessLayer.Reports.CrpOrderBookerOrderTakingSheet CrpReport = new SAMSBusinessLayer.Reports.CrpOrderBookerOrderTakingSheet();
            CrpReport.SetDataSource(ds);
            CrpReport.Refresh();
            CrpReport.SetParameterValue("Principal", this.drpPrincipal.SelectedItem.Text.ToString());
            CrpReport.SetParameterValue("Location", this.DrpLocation.SelectedItem.Text.ToString());
            CrpReport.SetParameterValue("OrderBooker", this.DrpOrderBooker.SelectedItem.Text.ToString());
            CrpReport.SetParameterValue("From_Date", this.txtFromDate.Text);
            CrpReport.SetParameterValue("To_Date", this.txtToDate.Text);
            CrpReport.SetParameterValue("CompanyName", dt.Rows[0]["COMPANY_NAME"].ToString());
            CrpReport.SetParameterValue("Area", this.DrpRoute.SelectedItem.Text.ToString());

            this.Session.Add("CrpReport", CrpReport);
            this.Session.Add("ReportType", Type);
            string url = "'Default.aspx'";
            string script = "<script language='JavaScript' type='text/javascript'> window.open(" + url + ",\"Link\",\"toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,width=800,height=600,left=10,top=10\");</script>";
            Type cstype = this.GetType();
            ClientScriptManager cs = Page.ClientScript;
            cs.RegisterStartupScript(cstype, "OpenWindow", script);

        }
        catch (Exception ex)
        {
            ex.ToString();
        }
    }
}