using System;
using System.Data;
using SAMSCommon.Classes;
using SAMSDataAccessLayer.Classes;
using SAMSDatabaseLayer.Classes;

namespace SAMSBusinessLayer.Classes
{
    /// <summary>
    /// Class For Fetching Data Of Inventory Reports
    /// </summary>
    public class RptInventoryController
    {
        #region Constructor

        /// <summary>
        /// Constructor for RptInventoryController
        /// </summary>
        public RptInventoryController()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		#endregion

        #region Public Methods

        /// <summary>
        /// Gets Data For Stock Reconciliation Report
        /// </summary>
        /// <param name="p_Distributor_ID">Location</param>
        /// <param name="p_Principal_Id">Principal</param>
        /// <param name="p_FromDate">DateFrom</param>
        /// <param name="p_To_Date">DateTo</param>
        /// <param name="p_USER_ID">User</param>
        /// <returns>DataSet</returns>
        public DataSet SelectPrincipalStockReconcilation(int p_Distributor_ID, string p_Principal_Id, DateTime p_FromDate, DateTime p_To_Date, int p_USER_ID, int p_UOM_ID, int p_PRICE_TYPE)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                uspRptSelectStockRegister ObjPrint = new uspRptSelectStockRegister();
                SAMSBusinessLayer.Reports.DsReport ds = new SAMSBusinessLayer.Reports.DsReport();
                ObjPrint.Connection = mConnection;
                ObjPrint.distributor_id = p_Distributor_ID;
                ObjPrint.Company_Id = p_Principal_Id;
                ObjPrint.DateFrom = p_FromDate;
                ObjPrint.dateto = p_To_Date;
                ObjPrint.USER_ID = p_USER_ID;
                ObjPrint.UOM_ID = p_UOM_ID;
                ObjPrint.PRICE_TYPE = p_PRICE_TYPE;
                DataTable dt = ObjPrint.ExecuteTable();
                foreach (DataRow dr in dt.Rows)
                {
                    ds.Tables["StockRegister"].ImportRow(dr);
                }

                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        /// <summary>
        /// Gets Data For Date Wise Stock Report
        /// </summary>
        /// <param name="p_Distributor_ID">Location</param>
        /// <param name="p_Principal_Id">Principal</param>
        /// <param name="p_FromDate">DateFrom</param>
        /// <param name="p_To_Date">DateTo</param>
        /// <param name="p_TypeId">Type</param>
        /// <returns>DataSet</returns>
        public DataSet SelectPurchaseTransferStock(int p_Distributor_ID, int p_Principal_Id, DateTime p_FromDate, DateTime p_To_Date, int p_TypeId, int p_RATE_TYPE)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                UspDailyPurchaseTransfer ObjPrint = new UspDailyPurchaseTransfer();
                SAMSBusinessLayer.Reports.DsReport ds = new SAMSBusinessLayer.Reports.DsReport();
                ObjPrint.Connection = mConnection;
                ObjPrint.DISTRIBUTOR_ID = p_Distributor_ID;
                ObjPrint.TYPEID = p_TypeId;
                ObjPrint.PRINCIPAL_ID = p_Principal_Id;
                ObjPrint.FROM_DATE = p_FromDate;
                ObjPrint.TO_DATE = p_To_Date;
                ObjPrint.RATE_TYPE = p_RATE_TYPE;
                DataTable dt = ObjPrint.ExecuteTable();
                foreach (DataRow dr in dt.Rows)
                {
                    ds.Tables["DailyPurchaseTransferReport"].ImportRow(dr);
                }

                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        #region Transfer In/Out Report

        /// <summary>
        /// Gets Data For Transfer In/Out Report (In Value)
        /// </summary>
        /// <param name="p_Principal_ID">Principal</param>
        /// <param name="p_Distributor_ID">Location</param>
        /// <param name="p_FromTime">DateFrom</param>
        /// <param name="p_ToDate">DateTo</param>
        /// <param name="p_TransferType">Type</param>
        /// <param name="p_type">ReportTyp</param>
        /// <returns>DataSet</returns>
        public DataSet TransferInOutValue(int p_Principal_ID, int p_Distributor_ID, DateTime p_FromTime, DateTime p_ToDate, string p_TransferType, int p_type)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                SAMSBusinessLayer.Reports.DsReport ds = new SAMSBusinessLayer.Reports.DsReport();

                UspTransferInOutValue mTransferIn = new UspTransferInOutValue();
                mTransferIn.Connection = mConnection;

                mTransferIn.PRINCIPAL_ID = p_Principal_ID;
                mTransferIn.DISTRIBUTOR_ID = p_Distributor_ID;
                mTransferIn.FromDate = p_FromTime;
                mTransferIn.ToDate = p_ToDate;
                mTransferIn.TransferType = p_TransferType;
                mTransferIn.ReportType = p_type;
                DataTable DT = mTransferIn.ExecuteTable();

                foreach (DataRow dr in DT.Rows)
                {
                    ds.Tables["RptTransferInOutValueWise"].ImportRow(dr);
                }
                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        /// <summary>
        /// Gets Data For Transfer In/Out Report (In Quantity And Carton)
        /// </summary>
        /// <param name="p_Principal_ID">Principal</param>
        /// <param name="p_Distributor_ID">Location</param>
        /// <param name="p_FromTime">DateFrom</param>
        /// <param name="p_ToDate">DateTo</param>
        /// <param name="p_TransferType">Type</param>
        /// <param name="p_type">ReportType</param>
        /// <returns>DataSet</returns>
        public DataSet TransferIn(int p_Principal_ID, int p_Distributor_ID, DateTime p_FromTime, DateTime p_ToDate, string p_TransferType, int p_type)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                SAMSBusinessLayer.Reports.DsReport ds = new SAMSBusinessLayer.Reports.DsReport();

                UspTransferInrpt mTransferIn = new UspTransferInrpt();
                mTransferIn.Connection = mConnection;

                mTransferIn.PRINCIPAL_ID = p_Principal_ID;
                mTransferIn.DISTRIBUTOR_ID = p_Distributor_ID;
                mTransferIn.FromDate = p_FromTime;
                mTransferIn.ToDate = p_ToDate;
                mTransferIn.TransferType = p_TransferType;
                mTransferIn.ReportType = p_type;
                DataTable DT = mTransferIn.ExecuteTable();

                foreach (DataRow dr in DT.Rows)
                {
                    ds.Tables["TransferIn"].ImportRow(dr);
                }
                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        #endregion

        #region  Physical Stock Report

        /// <summary>
        /// Gets Data For  Physical Stock Report (SKU Wise)
        /// </summary>
        /// <param name="p_Principal_ID">Principal</param>
        /// <param name="p_Distributor_ID">Location</param>
        /// <param name="p_Date">Date</param>
        /// <returns>DataSet</returns>
        public DataSet PhysicalStockTaking(int p_Principal_ID, int p_Distributor_ID, DateTime p_Date)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                SAMSBusinessLayer.Reports.DsReport ds = new SAMSBusinessLayer.Reports.DsReport();

                RptPhysicalStockTaking mStockTaking = new RptPhysicalStockTaking();

                mStockTaking.Connection = mConnection;
                mStockTaking.PRINCIPAL_ID = p_Principal_ID;
                mStockTaking.DISTRIBUTOR_ID = p_Distributor_ID;
                mStockTaking.Date = p_Date;

                DataTable DT = mStockTaking.ExecuteTable();
                foreach (DataRow dr in DT.Rows)
                {
                    ds.Tables["PhysicalStockTaking"].ImportRow(dr);
                }
                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        /// <summary>
        /// Gets Data For  Physical Stock Report (Value Wise)
        /// </summary>
        /// <param name="p_Principal_ID">Principal</param>
        /// <param name="p_Distributor_ID">Location</param>
        /// <param name="p_Date">Date</param>
        /// <param name="p_UserId">User</param>
        /// <returns>DataSet</returns>
        public DataSet PhysicalStockTakingValueWise(int p_Principal_ID, int p_Distributor_ID, DateTime p_Date, int p_UserId)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                SAMSBusinessLayer.Reports.DsReport ds = new SAMSBusinessLayer.Reports.DsReport();

                RptPhysicalStockTakingSummary mStockTaking = new RptPhysicalStockTakingSummary();

                mStockTaking.Connection = mConnection;
                mStockTaking.PRINCIPAL_ID = p_Principal_ID;
                mStockTaking.DISTRIBUTOR_ID = p_Distributor_ID;
                mStockTaking.Date = p_Date;
                mStockTaking.USER_ID = p_UserId;

                DataTable DT = mStockTaking.ExecuteTable();
                foreach (DataRow dr in DT.Rows)
                {
                    ds.Tables["RptPhysicalStockValue"].ImportRow(dr);
                }
                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        #endregion

        /// <summary>
        /// Gets Data For Purchase Document Report
        /// </summary>
        /// <param name="p_Distributor_ID">Location</param>
        /// <param name="p_Principal_Id">Principal</param>
        /// <param name="p_FromDate">DateFrom</param>
        /// <param name="p_To_Date">DateTo</param>
        /// <param name="p_TypeId">Type</param>
        /// <returns>DataSet</returns>
        public DataSet SelectPurchaseDocument(int p_Distributor_ID, int p_Principal_Id, DateTime p_FromDate, DateTime p_To_Date, int p_TypeId)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                GetPurchasedocument ObjPrint = new GetPurchasedocument();
                SAMSBusinessLayer.Reports.DsReport ds = new SAMSBusinessLayer.Reports.DsReport();
                ObjPrint.Connection = mConnection;
                ObjPrint.DISTRIBUTOR_ID = p_Distributor_ID;
                ObjPrint.PRINCIPAL_ID = p_Principal_Id;
                ObjPrint.FROM_DATE = p_FromDate;
                ObjPrint.TO_DATE = p_To_Date;
                ObjPrint.TYPE_ID = p_TypeId;
                DataTable dt = ObjPrint.ExecuteTable();
                foreach (DataRow dr in dt.Rows)
                {
                    ds.Tables["RptPurchaseDocument"].ImportRow(dr);
                }

                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        public DataSet selectExpiryDate(int p_Distributor_ID, int p_Principal_Id, DateTime p_FromDate, int p_TypeId)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                GetExipryDate ObjPrint = new GetExipryDate();
                SAMSBusinessLayer.Reports.DsReport ds = new SAMSBusinessLayer.Reports.DsReport();
                ObjPrint.Connection = mConnection;
                ObjPrint.DISTRIBUTOR_ID = p_Distributor_ID;
                ObjPrint.PRINCIPAL_ID = p_Principal_Id;
                ObjPrint.FROM_DATE = p_FromDate;
                ObjPrint.TO_DATE = Constants.DateNullValue;
                ObjPrint.TYPE_ID = p_TypeId;
                DataTable dt = ObjPrint.ExecuteTable();
                foreach (DataRow dr in dt.Rows)
                {
                    ds.Tables["RptPurchaseDocument"].ImportRow(dr);
                }

                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }


        #region Added By Hazrat Ali

        /// <summary>
        /// Gets Data For Stock Valuation Report
        /// </summary>
        /// <param name="p_StockDate">Date</param>
        /// <param name="p_Distributor_ID">Location</param>
        /// <param name="p_Principal_ID">Principal</param>
        /// <param name="p_USER_ID">User</param>
        /// <param name="p_ReportType">ReportType</param>
        /// <returns>DataSet</returns>
        public DataSet SelectStockValuation(DateTime p_StockDate, int p_Distributor_ID, int p_Principal_ID, int p_USER_ID, int p_ReportType)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                uspGetStockValuation ObjPrint = new uspGetStockValuation();
                SAMSBusinessLayer.Reports.DsReport2 ds = new SAMSBusinessLayer.Reports.DsReport2();
                ObjPrint.Connection = mConnection;
                ObjPrint.STOCK_DATE = p_StockDate;
                ObjPrint.DISTRIBUTOR_ID = p_Distributor_ID;
                ObjPrint.PRINCIPAL_ID = p_Principal_ID;
                ObjPrint.USER_ID = p_USER_ID;
                ObjPrint.TYPE = p_ReportType;
                DataTable dt = ObjPrint.ExecuteTable();
                if (p_ReportType == 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        ds.Tables["rptStockValuationDetail"].ImportRow(dr);
                    }
                }
                else
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        ds.Tables["rptStockValuationSummary"].ImportRow(dr);
                    }
                }

                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        public DataSet GetCategoryWisePurchase(int p_Principal_ID, int p_Distributor_ID, DateTime p_StartDate, DateTime p_EndDate, int p_CatagoryID)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                SAMSBusinessLayer.Reports.dsSalesPurchaseRegister ds = new SAMSBusinessLayer.Reports.dsSalesPurchaseRegister();

                uspGetCategoryWisePurchase mPurchase = new uspGetCategoryWisePurchase();
                mPurchase.Connection = mConnection;
                mPurchase.Principal_ID = p_Principal_ID;
                mPurchase.Distributor_ID = p_Distributor_ID;
                mPurchase.FromDate = p_StartDate;
                mPurchase.ToDate = p_EndDate;
                mPurchase.CATEGORY_ID = p_CatagoryID;

                DataTable DT = mPurchase.ExecuteTable();
                foreach (DataRow dr in DT.Rows)
                {
                    ds.Tables["uspGetCategoryWisePurchase"].ImportRow(dr);
                }
                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        public DataSet GetExpiryDamageData(int p_Distributor_ID, int p_Principal_Id, DateTime p_FromDate, DateTime p_To_Date, int p_TypeId)
        {
            IDbConnection mConnection = null;
            try
            {
                mConnection = ProviderFactory.GetConnection(Configuration.ConnectionString, EnumProviders.SQLClient);
                mConnection.Open();
                uspGetExpiryDamageData ObjPrint = new uspGetExpiryDamageData();
                SAMSBusinessLayer.Reports.dsSalesPurchaseRegister ds = new SAMSBusinessLayer.Reports.dsSalesPurchaseRegister();
                ObjPrint.Connection = mConnection;
                ObjPrint.DISTRIBUTOR_ID = p_Distributor_ID;
                ObjPrint.TYPEID = p_TypeId;
                ObjPrint.PRINCIPAL_ID = p_Principal_Id;
                ObjPrint.FROM_DATE = p_FromDate;
                ObjPrint.TO_DATE = p_To_Date;
                DataTable dt = ObjPrint.ExecuteTable();
                foreach (DataRow dr in dt.Rows)
                {
                    ds.Tables["dtExpiryDamage"].ImportRow(dr);
                }

                return ds;
            }
            catch (Exception exp)
            {
                ExceptionPublisher.PublishException(exp);
                return null;
            }
            finally
            {
                if (mConnection != null && mConnection.State == ConnectionState.Open)
                {
                    mConnection.Close();
                }
            }
        }

        #endregion

        #endregion
    }
}