using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using AccessMatrix.Core;
using System.Configuration;

namespace AccessMatrix
{
    /// <summary>
    /// 
    /// </summary>
    public static class AccessMatrixEngine
    {
        /// <summary>
        /// Get All Data From Configuration Table [TableName = @OCFL]
        /// </summary>
        /// <returns></returns>
        public static Recordset GetAllConfiguration()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT * FROM \"@OCFL\" ORDER BY \"U_AccessMatrixKey\"";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            return oRecordSet;
        }

        /// <summary>
        /// Get Data Source From Configration Table [TableName = @OCFL]
        /// If eventType is not provided, this function will assume as "Tab" event.
        /// </summary>
        /// <param name="formId"></param>
        /// <param name="itemId"></param>
        /// <param name="colId"></param>
        /// <param name="eventType"></param>
        /// <returns></returns>
        public static Recordset GetConfiguration(string formId, string itemId, string colId, int eventType)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            string documentType = String.Empty;
            try
            {
                colId = (colId == "" ? "0" : colId); //If colUID is "0", the event return as "".
                if (itemId == "38") // Item
                {
                    documentType = "1";
                }
                else if (itemId == "39") // Service
                {
                    documentType = "2";
                }
                else
                {
                    documentType = "3";
                }
                if(formId.Contains("UDO"))
                {
                    formId = System.Text.RegularExpressions.Regex.Replace(formId.Substring(formId.LastIndexOf('_') + 1), @"[\d-]", string.Empty);
                }
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT * FROM \"@OCFL\" WHERE \"U_FormId\" = '" + formId + "' AND \"U_ItemId\" = '" + itemId + "' AND \"U_ColumnId\" = '" + colId + "' AND \"U_EventType\" = " + eventType + " AND \"U_DocType\" = " + documentType;
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        public static Recordset GetConfigurationForLineLoop(string formId, string itemId, string colId, int eventType)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                colId = (colId == "" ? "0" : colId); //If colUID is "0", the event return as "".
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT * FROM \"@OCFL\" WHERE \"U_FormId\" = '" + formId + "' AND \"U_ItemId\" = '" + itemId + "' AND \"U_ColumnId\" = '" + colId + "' AND \"U_EventType\" = " + eventType;
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            return oRecordSet;
        }

        /// <summary>
        /// Get Business Partner Group Name by Group Code.
        /// </summary>
        /// <returns>Business Partner Group Name.</returns>
        public static string GetBPGroupNameByGroupCode(string strBPGroupCode)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"GroupName\" FROM OCRG T0 WHERE T0.\"GroupCode\" = " + strBPGroupCode;
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet.Fields.Item(0).Value.ToString();
        }

        public static Recordset GetBusinessPartnersByUser()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = String.Empty;
                if(ConfigurationManager.AppSettings["DbType"].ToString().ToUpper() == "SQL")
                    strQuery = "SELECT T0.\"CardCode\", T0.\"CardName\" FROM OCRD T0 INNER JOIN \"@UBP1\" T1 ON T1.\"U_GroupCode\" = T0.\"GroupCode\" INNER JOIN \"@OUBP\" T2 ON T2.\"Code\" = T1.\"Code\" WHERE T2.\"U_User\" = (SELECT \"USER_CODE\" FROM OUSR WHERE \"INTERNAL_K\" = " + AddOnUtilities.oCompany.UserSignature + ")";
                else
                    strQuery = "SELECT T0.\"CardCode\", T0.\"CardName\" FROM OCRD T0 INNER JOIN \"@UBP1\" T1 ON T1.\"U_GroupCode\" = T0.\"GroupCode\" INNER JOIN \"@OUBP\" T2 ON T2.\"Code\" = T1.\"Code\" WHERE T2.\"U_User\" = (SELECT \"USER_CODE\" FROM OUSR WHERE \"INTERNAL_K\" = " + AddOnUtilities.oCompany.UserSignature + ")";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        public static Recordset GetBusinessPartnerNamesByUser()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = String.Empty;
                if (ConfigurationManager.AppSettings["DbType"].ToString().ToUpper() == "SQL")
                    strQuery = "SELECT T0.\"CardName\", T0.\"CardCode\" FROM OCRD T0 INNER JOIN \"@UBP1\" T1 ON T1.\"U_GroupCode\" = T0.\"GroupCode\" INNER JOIN \"@OUBP\" T2 ON T2.\"Code\" = T1.\"Code\" WHERE T2.\"U_User\" = (SELECT \"USER_CODE\" FROM OUSR WHERE \"INTERNAL_K\" = " + AddOnUtilities.oCompany.UserSignature + ")";
                else
                    strQuery = "SELECT T0.\"CardName\", T0.\"CardCode\" FROM OCRD T0 INNER JOIN \"@UBP1\" T1 ON T1.\"U_GroupCode\" = T0.\"GroupCode\" INNER JOIN \"@OUBP\" T2 ON T2.\"Code\" = T1.\"Code\" WHERE T2.\"U_User\" = (SELECT \"USER_CODE\" FROM OUSR WHERE \"INTERNAL_K\" = " + AddOnUtilities.oCompany.UserSignature + ")";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        public static Recordset GetAllBusinessPartnerGroups()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = String.Empty;
                if (ConfigurationManager.AppSettings["DbType"].ToString().ToUpper() == "SQL")
                    strQuery = "SELECT T0.\"GroupCode\", T0.\"GroupName\" FROM OCRG T0";
                else
                    strQuery = "SELECT T0.\"GroupCode\", T0.\"GroupName\" FROM OCRG T0";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        public static Recordset GetProjectsByUser()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T1.\"U_Project\" FROM \"@OUPR\" T0 INNER JOIN \"@UPR1\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_User\" = (SELECT \"USER_CODE\" FROM OUSR WHERE \"INTERNAL_K\" = " + AddOnUtilities.oCompany.UserSignature + ")";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        public static Recordset GetAccountsByUser()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T1.\"U_AccountCode\" FROM \"@OUAC\" T0 JOIN \"@UAC1\" T1 ON T0.\"Code\" = T1.\"Code\" JOIN \"OACT\" T2 ON T1.\"U_AccountCode\" = T2.\"AcctCode\" WHERE T0.\"U_User\" = (SELECT \"USER_CODE\" FROM OUSR WHERE \"INTERNAL_K\" = " + AddOnUtilities.oCompany.UserSignature + ")";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        public static string GetAccountNameByCode(string strAccountCode)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"AcctName\" FROM OACT T0 WHERE T0.\"AcctCode\" = '" + strAccountCode + "'";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet.Fields.Item(0).Value.ToString();
        }

        public static Recordset GetAllAccounts()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"AcctCode\", T0.\"AcctName\" FROM OACT T0";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        public static string GetNameByCode(string strCode, string strAccessMatrixKey)
        {
            string strName = "";
            switch (strAccessMatrixKey)
            {
                case "5":
                    strName = GetProjectNameByProjectCode(strCode);
                    break;
                case "6":
                    strName = GetDimension1NameByCode(strCode);
                    break;
                case "7":
                    strName = GetDimension2NameByCode(strCode);
                    break;
                case "8":
                    strName = GetDimension3NameByCode(strCode);
                    break;
                case "9":
                    strName = GetDimension4NameByCode(strCode);
                    break;
                case "10":
                    strName = GetDimension5NameByCode(strCode);
                    break;
                case "12":
                    strName = GetBPGroupNameByGroupCode(strCode);
                    break;
                case "13":
                    strName = GetAccountNameByCode(strCode);
                    break;
            }
            return strName;
        }

        /// <summary>
        /// Retrieve all users from SAP.
        /// </summary>
        /// <returns>RecordSet filled with SAP Users.</returns>
        public static Recordset GetUsers()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"USER_CODE\" FROM OUSR T0";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Retrieve all projects from SAP.
        /// </summary>
        /// <returns>RecordSet filled with Projects.</returns>
        public static Recordset GetProjects()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"PrjCode\", T0.\"PrjName\" FROM OPRJ T0";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Retrieve Project Name by Project Code.
        /// </summary>
        /// <returns>Project Name</returns>
        public static string GetProjectNameByProjectCode(string strProjectCode)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"PrjName\" FROM OPRJ T0 WHERE T0.\"PrjCode\" = '" + strProjectCode + "'";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet.Fields.Item(0).Value.ToString();
        }

        /// <summary>
        /// Retrieve all branches from SAP.
        /// </summary>
        /// <returns>RecordSet filled with Branches.</returns>
        public static Recordset GetBranches()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"BPLName\" FROM OBPL T0";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Retrieve all dimension 1 from SAP.
        /// </summary>
        /// <returns>RecordSet filled with Dimension 1.</returns>
        public static Recordset GetDimension1()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrCode\", T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 1";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Retrieve dimension 1 Name by dimension 1 Code.
        /// </summary>
        /// <returns>Dimension 1 Name.</returns>
        public static string GetDimension1NameByCode(string strDimension1Code)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 1 AND T0.\"OcrCode\" = '" + strDimension1Code + "'";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet.Fields.Item(0).Value.ToString();
        }

        /// <summary>
        /// Retrieve all dimension 2 from SAP.
        /// </summary>
        /// <returns>RecordSet filled with Dimension 2.</returns>
        public static Recordset GetDimension2()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrCode\", T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 2";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Retrieve dimension 2 Name by dimension 1 Code.
        /// </summary>
        /// <returns>Dimension 2 Name.</returns>
        public static string GetDimension2NameByCode(string strDimension2Code)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 2 AND T0.\"OcrCode\" = '" + strDimension2Code + "'";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet.Fields.Item(0).Value.ToString();
        }

        /// <summary>
        /// Retrieve all dimension 3 from SAP.
        /// </summary>
        /// <returns>RecordSet filled with Dimension 3.</returns>
        public static Recordset GetDimension3()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrCode\", T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 3";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Retrieve dimension 3 Name by dimension 3 Code.
        /// </summary>
        /// <returns>Dimension 3 Name.</returns>
        public static string GetDimension3NameByCode(string strDimension3Code)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 3 AND T0.\"OcrCode\" = '" + strDimension3Code + "'";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet.Fields.Item(0).Value.ToString();
        }

        /// <summary>
        /// Retrieve all dimension 4 from SAP.
        /// </summary>
        /// <returns>RecordSet filled with Dimension 4.</returns>
        public static Recordset GetDimension4()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrCode\", T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 4";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Retrieve dimension 4 Name by dimension 4 Code.
        /// </summary>
        /// <returns>Dimension 4 Name.</returns>
        public static string GetDimension4NameByCode(string strDimension4Code)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 4 AND T0.\"OcrCode\" = '" + strDimension4Code + "'";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet.Fields.Item(0).Value.ToString();
        }

        /// <summary>
        /// Retrieve all dimension 5 from SAP.
        /// </summary>
        /// <returns>RecordSet filled with Dimension 5.</returns>
        public static Recordset GetDimension5()
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrCode\", T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 5";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Retrieve dimension 5 Name by dimension 5 Code.
        /// </summary>
        /// <returns>Dimension 5 Name.</returns>
        public static string GetDimension5NameByCode(string strDimension5Code)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "SELECT T0.\"OcrName\" FROM OOCR T0 WHERE T0.\"DimCode\" = 5 AND T0.\"OcrCode\" = '" + strDimension5Code + "'";
                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet.Fields.Item(0).Value.ToString();
        }

        /// <summary>
        /// Get Available A Type of Dimensions List Defined by Project.
        /// </summary>
        /// <param name="dimensionType"></param>
        /// <returns></returns>
        public static Recordset GetDimensionsByBranch(string dimensionType, string branch = "", bool isBranchId = false)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string strQuery = "";
                if(isBranchId)
                {
                    switch (dimensionType)
                    {
                        case "Dimension1":
                            strQuery = "SELECT T1.\"U_Dimension1\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM1\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = (SELECT \"BPLName\" FROM OBPL WHERE \"BPLId\" = " + branch + ")";
                            break;
                        case "Dimension2":
                            strQuery = "SELECT T1.\"U_Dimension2\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM2\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = (SELECT \"BPLName\" FROM OBPL WHERE \"BPLId\" = " + branch + ")";
                            break;
                        case "Dimension3":
                            strQuery = "SELECT T1.\"U_Dimension3\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM3\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = (SELECT \"BPLName\" FROM OBPL WHERE \"BPLId\" = " + branch + ")";
                            break;
                        case "Dimension4":
                            strQuery = "SELECT T1.\"U_Dimension4\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM4\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = (SELECT \"BPLName\" FROM OBPL WHERE \"BPLId\" = " + branch + ")";
                            break;
                        case "Dimension5":
                            strQuery = "SELECT T1.\"U_Dimension5\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM5\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = (SELECT \"BPLName\" FROM OBPL WHERE \"BPLId\" = " + branch + ")";
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    switch (dimensionType)
                    {
                        case "Dimension1":
                            strQuery = "SELECT T1.\"U_Dimension1\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM1\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = '" + branch + "'";
                            break;
                        case "Dimension2":
                            strQuery = "SELECT T1.\"U_Dimension2\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM2\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = '" + branch + "'";
                            break;
                        case "Dimension3":
                            strQuery = "SELECT T1.\"U_Dimension3\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM3\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = '" + branch + "'";
                            break;
                        case "Dimension4":
                            strQuery = "SELECT T1.\"U_Dimension4\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM4\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = '" + branch + "'";
                            break;
                        case "Dimension5":
                            strQuery = "SELECT T1.\"U_Dimension5\" FROM \"@OBDM\" T0 INNER JOIN \"@BDM5\" T1 ON T1.\"Code\" = T0.\"Code\" WHERE T0.\"U_Branch\" = '" + branch + "'";
                            break;
                        default:
                            break;
                    }
                }

                oRecordSet.DoQuery(strQuery);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oRecordSet;
        }

        /// <summary>
        /// Check if User's input value is valid or not.
        /// </summary>
        /// <param name="rsConfigData"></param>
        /// <returns>If input is invalid, return value starts with "Invalid". If input is valid, return "".</returns>
        public static string CheckIfValidOrNot(SAPbobsCOM.Recordset rsConfigData)
        {
            try
            {
                string controlName = rsConfigData.Fields.Item("U_ControlName").Value.ToString();
                string controlType = rsConfigData.Fields.Item("U_ControlType").Value.ToString();
                string accessMatrixKey = rsConfigData.Fields.Item("U_AccessMatrixKey").Value.ToString();
                string itemId = rsConfigData.Fields.Item("U_ItemId").Value.ToString();
                string colId = rsConfigData.Fields.Item("U_ColumnId").Value.ToString();
                string input = GetInputValue(itemId, colId, controlType);
                string result = String.Empty;

                SAPbobsCOM.Recordset rs = null;

                //Retrieve valid data according to the control name.
                if (controlName == "Business Partner")
                {
                    rs = GetBusinessPartnersByUser();
                }
                else if (controlName == "Business Partner Name")
                {
                    rs = GetBusinessPartnerNamesByUser();
                }
                else if (controlName == "All Business Partner Group")
                {
                    rs = GetAllBusinessPartnerGroups();
                }
                else if (controlName.StartsWith("Dimension")) //Dimension by Branch
                {
                    string branch = "";
                    string relatedItemId = rsConfigData.Fields.Item("U_RelatedItem").Value.ToString();
                    string relatedColId = rsConfigData.Fields.Item("U_RelatedColumn").Value.ToString();
                    bool isBranchId = false;

                    //Use Dynamic because Control Type can be identified only during runtime. It depends on the document types.
                    //If Control type is ComboBox, branch value from UI is ID.
                    //If Control type is EditText, branch value from UI is Value.
                    dynamic BranchControl = AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(relatedItemId).Specific;
                    if(BranchControl is SAPbouiCOM.ComboBox)
                    {
                        branch = ((SAPbouiCOM.ComboBox)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(relatedItemId).Specific).Value;
                        isBranchId = true;
                    }
                    else //if(BranchControl is SAPbouiCOM.EditText)
                    {
                        branch = ((SAPbouiCOM.EditText)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(relatedItemId).Specific).Value;
                        isBranchId = false;
                    }
                    
                    rs = GetDimensionsByBranch(controlName, branch, isBranchId);
                }
                else if (controlName == "Project")
                {
                    rs = GetProjectsByUser();
                }
                else if (controlName == "All Users")
                {
                    rs = GetUsers();
                }
                else if (controlName == "All Projects")
                {
                    rs = GetProjects();
                }
                else if (controlName.Contains("Dimension"))
                {
                    string dimensionType = controlName.ElementAt(controlName.Length - 1).ToString();
                    switch (dimensionType)
                    {
                        case "1":
                            rs = GetDimension1();
                            break;
                        case "2":
                            rs = GetDimension2();
                            break;
                        case "3":
                            rs = GetDimension3();
                            break;
                        case "4":
                            rs = GetDimension4();
                            break;
                        case "5":
                            rs = GetDimension5();
                            break;
                    }
                }
                else if (controlName == "All Branches")
                {
                    rs = GetBranches();
                }
                else if (controlName == "G/L Account")
                {
                    rs = GetAccountsByUser();
                }
                else if (controlName == "All Accounts")
                {
                    rs = GetAllAccounts();
                }

                //Retrieve user's input value.
                //input = GetInputValue(itemId, colId, controlType);

                // Below switch case statement is used to identify which kind of data is valid/invalid. So that error can be showed as "Invalid Input Type (Project, User, Dimension, etc.)
                // Check if user's input is in valid values.
                switch (accessMatrixKey)
                {
                    //------------------------------------------------------------------------ Checking for Projects ------------------------------------------------------------------------//
                    case "1": // Projects by User
                    case "5": // All Projects
                        if (input != "")
                        {
                            while (!rs.EoF)
                            {
                                string retrievedData = rs.Fields.Item(0).Value.ToString();
                                rs.MoveNext();
                                if (input == retrievedData)
                                {
                                    return "";
                                }
                            }
                            return "Invalid Project.";
                        }
                        break;
                    //------------------------------------------------------------------------ End of Checking for Projects ------------------------------------------------------------------------//

                    //------------------------------------------------------------------------ Checking for Dimensions ------------------------------------------------------------------------//
                    case "2": // Dimensions by Branch
                    case "6": // All Dimension 1
                    case "7": // All Dimension 2
                    case "8": // All Dimension 3
                    case "9": // All Dimension 4
                    case "10": // All Dimension 5
                        if (input != "") // To check if project column is blank or not first.
                        {
                            while (!rs.EoF)
                            {
                                string retrievedData = rs.Fields.Item(0).Value.ToString();
                                rs.MoveNext();
                                if (input == retrievedData)
                                {
                                    return "";
                                }
                            }
                            return "Invalid Dimension.";
                        }
                        break;
                    //------------------------------------------------------------------------ End of Checking for Dimensions ------------------------------------------------------------------------//

                    //------------------------------------------------------------------------ Checking for Business Partners ------------------------------------------------------------------------//
                    case "3": // Business Partners by User
                    case "12": // All Business Partners
                        if (input != "")
                        {
                            while (!rs.EoF)
                            {
                                string retrievedData = rs.Fields.Item(0).Value.ToString();
                                rs.MoveNext();
                                if (input == retrievedData)
                                {
                                    return "";
                                }
                            }
                            return "Invalid Business Partner.";
                        }
                        break;
                    //------------------------------------------------------------------------ End of Checking for Business Partners ------------------------------------------------------------------------//

                    //------------------------------------------------------------------------ Start of Checking for G/L Account ------------------------------------------------------------------------//
                    case "13": // Chart of Accounts By User
                        if (input != "")
                        {
                            while (!rs.EoF)
                            {
                                string retrievedData = rs.Fields.Item(0).Value.ToString();
                                rs.MoveNext();
                                if (input == retrievedData)
                                {
                                    return "";
                                }
                            }
                            return "Invalid Account Code.";
                        }
                        break;
                    //------------------------------------------------------------------------ End of Checking for G/L Account ------------------------------------------------------------------------//

                    //------------------------------------------------------------------------ Checking for Users ------------------------------------------------------------------------//
                    case "4": // All Users
                        if (input != "") // To check if project column is blank or not first.
                        {
                            while (!rs.EoF)
                            {
                                string retrievedData = rs.Fields.Item(0).Value.ToString();
                                rs.MoveNext();
                                if (input == retrievedData)
                                {
                                    return "";
                                }
                            }
                            return "Invalid User.";
                        }
                        break;
                    //------------------------------------------------------------------------ End of Checking for Users ------------------------------------------------------------------------//
                    case "11":
                        if (input != "") // To check if project column is blank or not first.
                        {
                            while (!rs.EoF)
                            {
                                string retrievedData = rs.Fields.Item(0).Value.ToString();
                                rs.MoveNext();
                                if (input == retrievedData)
                                {
                                    return "";
                                }
                            }
                            return "Invalid Branch.";
                        }
                        break;
                    default:
                        break;
                }
                return "";
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
                return ex.Message + " " + ex.StackTrace;
            }
        }

        public static string GetInputValue(string itemId, string colId, string controlType, string relatedCol = "")
        {
            string input = String.Empty;

            try
            {
                int rowIndex;
                SAPbouiCOM.Matrix oMatrix = null;
                switch (controlType)
                {
                    case "1": //Textbox
                        input = ((SAPbouiCOM.EditText)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(itemId).Specific).Value;
                        break;
                    case "2": //Checkbox (Only for BP Selection Window)
                        string colIdForDescription = (relatedCol == "" ? itemId : relatedCol);
                        oMatrix = (SAPbouiCOM.Matrix)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(itemId).Specific;
                        rowIndex = oMatrix.GetCellFocus().rowIndex;
                        input = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(colIdForDescription).Cells.Item(rowIndex).Specific).Value;
                        break;
                    case "3": //Dropdown List
                        break;
                    case "4": //Matrix
                        oMatrix = (SAPbouiCOM.Matrix)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(itemId).Specific;
                        rowIndex = oMatrix.GetCellFocus().rowIndex;
                        input = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(colId).Cells.Item(rowIndex).Specific).Value;
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return input;
        }

        public static void SetValueToMatrix(string itemId, string colId, string value)
        {
            var currentEventFilters = AddOnUtilities.oApplication.GetFilter();
            AddOnUtilities.oApplication.SetFilter(null);
            
            // do your actions here
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(itemId).Specific;
            int rowIndex = oMatrix.GetCellFocus().rowIndex;
            ((SAPbouiCOM.EditText)oMatrix.Columns.Item(colId).Cells.Item(rowIndex).Specific).Value = value;

            AddOnUtilities.oApplication.SetFilter(currentEventFilters);
        }


        public static string GetMatrixData(SAPbouiCOM.Matrix oMatrix, string itemId, string colId, int rowIndex)
        {
            string input = String.Empty;
            try
            {
                input = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(colId).Cells.Item(rowIndex).Specific).Value;
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return input;
        }

        public static string CheckIfValidOrNotAllLine(SAPbobsCOM.Recordset rsConfigData, string documentType)
        {
            try
            {
                string controlName = rsConfigData.Fields.Item("U_ControlName").Value.ToString();
                string controlType = rsConfigData.Fields.Item("U_ControlType").Value.ToString();
                string accessMatrixKey = rsConfigData.Fields.Item("U_AccessMatrixKey").Value.ToString();
                string itemId = rsConfigData.Fields.Item("U_ItemId").Value.ToString();
                string colId = rsConfigData.Fields.Item("U_ColumnId").Value.ToString();
                string relatedItemId = rsConfigData.Fields.Item("U_RelatedItem").Value.ToString();
                string relatedColId = rsConfigData.Fields.Item("U_RelatedColumn").Value.ToString();
                string relatedItemId2 = rsConfigData.Fields.Item("U_RelatedItem2").Value.ToString();
                string relatedColId2 = rsConfigData.Fields.Item("U_RelatedColumn2").Value.ToString();
                string result = "Invalid " + controlName + " at line : ";
                bool isValid = true, isValidAll = true;

                SAPbobsCOM.Recordset rs = null;

                if (controlName == "Business Partners")
                {
                    rs = GetBusinessPartnersByUser();
                }
                else if (controlName.Contains("Dimension"))
                {
                    string branch = "";
                    bool isBranchId = false;

                    dynamic BranchControl = AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(relatedItemId2).Specific;

                    if (BranchControl is SAPbouiCOM.ComboBox)
                    {
                        branch = ((SAPbouiCOM.ComboBox)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(relatedItemId2).Specific).Value;
                        isBranchId = true;
                    }
                    else //if(BranchControl is SAPbouiCOM.EditText)
                    {
                        branch = ((SAPbouiCOM.EditText)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(relatedItemId2).Specific).Value;
                        isBranchId = false;
                    }

                    rs = GetDimensionsByBranch(controlName, branch, isBranchId);
                }
                else if (controlName == "Project")
                {
                    rs = GetProjectsByUser();
                }
                else if (controlName == "G/L Account")
                {
                    rs = GetAccountsByUser();
                }

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(relatedItemId).Specific;
                int RowCount = oMatrix.RowCount;

                for (int CurrentRow = 1; CurrentRow < RowCount; CurrentRow++)
                {
                    string input = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(relatedColId).Cells.Item(CurrentRow).Specific)).Value;
                    rs.MoveFirst();
                    if (input != "")
                    {
                        while (!rs.EoF)
                        {
                            string retrievedData = rs.Fields.Item(0).Value.ToString();
                            rs.MoveNext();
                            if (input == retrievedData)
                            {
                                isValid = true;
                                break;
                            }
                            else
                            {
                                isValid = false;
                            }
                        }
                        if (!isValid && rs.EoF)
                        {
                            isValidAll = false;
                            result += CurrentRow.ToString() + ", ";
                        }
                    }
                }
                if (!isValidAll)
                {
                    result = result.Substring(0, result.Length - 2) + ".";
                    return result;
                }
                return "";

            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
                return ex.Message + " " + ex.StackTrace;
            }
        }
    }
}
