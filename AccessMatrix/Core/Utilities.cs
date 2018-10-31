using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Configuration;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;
using AccessMatrix.Enum;
using AccessMatrix.Properties;
using System.Resources;
using System.Globalization;
using System.Collections;

namespace AccessMatrix.Core
{
    public static class AddOnUtilities
    {
        public static int IRetCode;
        public static int IErrCode;
        public static String sErrMsg;

        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbouiCOM.Application oApplication = null;

        //public static string s_connectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            try
            {
                // A.Check License
                // B. Connection : DI/UI/SSO/ etc.
                Connect();

                // C. Add-On requirement - UDT UDF UDO
                // Below function is just for testing!
                //ClearAllUDV();

                // TODO : add bool to avoid execute the function below
                Metadata.CreateTableForCustomizeForms();


                // D. (To-Do) Register UDO
                // E. (To-Do) UI: Create Menus
                // F. (To-Do) Add User Queries
                CreateCustomQueries(ConfigurationManager.AppSettings["QueryCategory"]);

                // G. Create UDV to filter controls
                CreateUDV();

                //Test testForm = new AccessMatrix.Test();
                //testForm.ShowDialog();
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public static bool CreateCustomQueries(string qryCategory)
        {
            Recordset oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            UserQueries oUserQueries = (SAPbobsCOM.UserQueries)AddOnUtilities.oCompany.GetBusinessObject(BoObjectTypes.oUserQueries);
            QueryCategories oQueryCategories = (SAPbobsCOM.QueryCategories)AddOnUtilities.oCompany.GetBusinessObject(BoObjectTypes.oQueryCategories);
            int lRetCode = 0, queryCategoryId = 0;

            try
            {
                string strQuery = "SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\" = '" + qryCategory + "'";
                oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRecordSet.DoQuery(strQuery);
                if(oRecordSet.RecordCount < 1)
                {
                    string tmp = oRecordSet.Fields.Item(0).Value.ToString();

                    bool isQueryCategoryExist = Convert.ToBoolean(oRecordSet.Fields.Item(0).Value);
                    if (!isQueryCategoryExist) //Create only if query category is not created yet.
                    {
                        oQueryCategories.Name = qryCategory;
                        oQueryCategories.Permissions = "YYYYYYYYYYYYYYY";
                        lRetCode = oQueryCategories.Add();

                        if (lRetCode != 0)
                        {
                            string errMsg = AddOnUtilities.oCompany.GetLastErrorDescription();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oQueryCategories);
                    }
                }

                List<string[]> lstQueryList = AddOnUtilities.ReadQueries();
                queryCategoryId = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                foreach (string[] query in lstQueryList)
                {
                    string queryDescription = query[0];
                    strQuery = "SELECT \"QName\" FROM OUQR WHERE \"QCategory\" = (SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\" = '" + qryCategory + "') AND \"QName\" = '" + queryDescription + "'";
                    oRecordSet.DoQuery(strQuery);
                    if (oRecordSet.RecordCount == 0)
                    {
                        oUserQueries.QueryDescription = queryDescription;
                        oUserQueries.Query = query[1];
                        oUserQueries.QueryType = UserQueryTypeEnum.uqtWizard;
                        oUserQueries.QueryCategory = queryCategoryId;
                        lRetCode = oUserQueries.Add();

                        if (lRetCode != 0)
                        {
                            string errMsg = AddOnUtilities.oCompany.GetLastErrorDescription();
                            AddOnUtilities.MsgBoxWrapper(errMsg, MsgBoxType.B1StatusBar, BoMessageTime.bmt_Short, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oQueryCategories);
            }

            return lRetCode == 0;
        }

        //static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        //{
        //    switch (EventType)
        //    {
        //        case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
        //            //Exit Add-On
        //            System.Windows.Forms.Application.Exit();
        //            break;
        //        case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
        //            break;
        //        case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
        //            break;
        //        case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
        //            break;
        //        case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
        //            break;
        //        default:
        //            break;
        //    }
        //}

        internal static bool UDOExist(string udoName)
        {
            UserObjectsMD UDO = (UserObjectsMD)AddOnUtilities.oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
            try
            {
                bool result = UDO.GetByKey(udoName);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UDO);
                UDO = null;
                return result;
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UDO);
            }
        }

        public static void Connect()
        {
            switch (ConfigurationManager.AppSettings["ConnectionType"])
            {
                case "DI":
                    ConnectViaDI();
                    break;
                case "UI":
                    ConnectViaUI();
                    break;
                case "SSO":
                    ConnectViaUISingleSignOn();
                    break;
                case "MULTIADDON":
                    //To-Do
                    break;
                default:
                    break;
            }
        }

        public static void ConnectViaUISingleSignOn()
        {
            try
            {
                // 1. Connect via UI
                ConnectViaUI();
                // 2. Create a DI Company and get connect cookie
                oCompany = new SAPbobsCOM.Company();
                String sCookie = oCompany.GetContextCookie();
                // 3. Get Connection Info from UI Application for DI Company
                String connContext = oApplication.Company.GetConnectionContext(sCookie);
                // 4. Set the connection info for DI Company
                oCompany.SetSboLoginContext(connContext);
                IRetCode = oCompany.Connect();

                // 5. Error Handling
                if (IRetCode != 0)
                {
                    MsgBoxWrapper(oCompany.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
        }

        public static void ConnectViaUI()
        {
            try
            {
                SboGuiApi uiApi = new SboGuiApi();
                // 1. Get the Connection String
                String connectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                // 2. Connect to UI API Object
                uiApi.Connect(connectionString);
                // License for your Add On (your IP)
                // Add-On Identifier
                // 3. Get the B1 Application Object
                oApplication = uiApi.GetApplication();
                // 4. (To-Do) Delegate other sub component (e.g. AppEvent, etc.)
                GenericEventHandler.RegisterEventHandler();
                // 5. (To-Do) Setup Event filters
            }
            catch (Exception ex)
            {
                MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
        }

        public static void ConnectViaDI()
        {
            String serverName, databaseName, sapUserName, sapPassword, licenseServer;
            //if (Resources.AppStage.ToString().ToLower() == "deployment")
            //{
            //    // lcm assignment - Parameters | lcm_deployment - shared parameters
            //    serverName = oApplication.Company.GetExtensionProperty(connectionString, SAPbouiCOM.BoExtensionLCMStageType.lcm_assignment, "servername");
            //    databaseName = oApplication.Company.GetExtensionProperty(connectionString, SAPbouiCOM.BoExtensionLCMStageType.lcm_assignment, "databasename");
            //    sapUserName = oApplication.Company.GetExtensionProperty(connectionString, SAPbouiCOM.BoExtensionLCMStageType.lcm_assignment, "username");
            //    sapPassword = oApplication.Company.GetExtensionProperty(connectionString, SAPbouiCOM.BoExtensionLCMStageType.lcm_assignment, "password");
            //    licenseServer = oApplication.Company.GetExtensionProperty(connectionString, SAPbouiCOM.BoExtensionLCMStageType.lcm_assignment, "licenseserver");
            //}
            //else
            //{
            //    serverName = AccessMatrix.Properties.Settings.Default.ServerName.ToString();
            //    databaseName = AccessMatrix.Properties.Settings.Default.DatabaseName.ToString();
            //    sapUserName = AccessMatrix.Properties.Settings.Default.SAPUserName.ToString();
            //    sapPassword = AccessMatrix.Properties.Settings.Default.SAPPassword.ToString();
            //    licenseServer = AccessMatrix.Properties.Settings.Default.LicenseServer.ToString();
            //}
            serverName = ConfigurationManager.AppSettings["ServerName"];
            databaseName = ConfigurationManager.AppSettings["DatabaseName"];
            sapUserName = ConfigurationManager.AppSettings["SAPUserName"];
            sapPassword = ConfigurationManager.AppSettings["SAPPassword"];
            licenseServer = ConfigurationManager.AppSettings["LicenseServer"];
            
            BoDataServerTypes dbType = BoDataServerTypes.dst_HANADB;

            if (ConfigurationManager.AppSettings["DbType"] == "SQL")
            {
                dbType = ConfigurationManager.AppSettings["DbType"] == "2012" ? BoDataServerTypes.dst_MSSQL2012 : BoDataServerTypes.dst_MSSQL2014;
            }
            else
            {
                dbType = BoDataServerTypes.dst_HANADB;
            }
            ConnectViaDI(serverName, dbType, databaseName, sapUserName, sapPassword, licenseServer, BoSuppLangs.ln_English);
        }

        public static void ConnectViaDI(String dbServer,
                                BoDataServerTypes dbServerType,
                                String companyDb, String user,
                                String password,
                                String licenseServer,
                                BoSuppLangs language = BoSuppLangs.ln_English)
        {
            try
            {
                // 1. Initiate the company instance for the first time
                if (oCompany == null)
                    oCompany = new SAPbobsCOM.Company();
                // 2. Set Company Property for login
                oCompany.Server = dbServer;
                oCompany.DbServerType = dbServerType;
                oCompany.CompanyDB = companyDb;
                oCompany.UserName = user;
                oCompany.Password = password;
                oCompany.LicenseServer = licenseServer;
                oCompany.language = language;
                oCompany.UseTrusted = false;

                IRetCode = oCompany.Connect();

                // 4. Error Handling
                if (IRetCode != 0)
                {
                    MsgBoxWrapper(oCompany.GetLastErrorDescription());
                }

                // 5. (Optional) XML Export Configuration
                oCompany.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;
            }
            catch (Exception ex)
            {
                MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
        }

        public static void DIErrorHandler(String op)
        {
            String msg = String.Format("{0} succeeded.", op);

            if (IRetCode != 0)
            {
                //Error
                oCompany.GetLastError(out IErrCode, out sErrMsg);
                msg = String.Format("{0} operation failed. Error Code: {1}. Error Message {2}", op, IErrCode, sErrMsg);
            }
            else
            {
                //Succeed. Do Nothing
            }
            //MsgBoxWrapper(msg);
        }

        public static void MsgBoxWrapper(String msg, MsgBoxType msgBoxType = MsgBoxType.B1MsgBox, BoMessageTime msgTime = BoMessageTime.bmt_Short, bool isError = true)
        {
            if (msgBoxType == MsgBoxType.B1MsgBox)
            {
                oApplication.MessageBox(msg);
            }
            else
            {
                oApplication.SetStatusBarMessage(msg, msgTime, isError);
            }
        }

        public static List<string[]> ReadQueries()
        {
            List<string[]> lstQueryList = new List<string[]>();

            ResourceManager MyResourceClass = new ResourceManager(typeof(Resources /* Reference to your resources class -- may be named differently in your case */));
            ResourceSet resourceSet = MyResourceClass.GetResourceSet(CultureInfo.CurrentUICulture, true, true);
            foreach (DictionaryEntry entry in resourceSet)
            {
                string resourceKey = entry.Key.ToString();
                object resource = entry.Value;

                if(resourceKey.Contains("QRY_"))
                {
                    string[] strQuery = { resourceKey, resource.ToString() };
                    lstQueryList.Add(strQuery);
                }
            }
            return lstQueryList;
        }

        // TODO : refactored this function
        public static List<string[]> ReadQueries(String filePath)
        {
            List<string[]> lstQueryList = new List<string[]>();

            if (!File.Exists(filePath))
            {
                MsgBoxWrapper("File Not Found.", MsgBoxType.WinMsgBox, BoMessageTime.bmt_Short, true);
                return null;
            }
            using (var reader = new StreamReader(filePath))
            {
                string line = reader.ReadLine();
                char[] delimiters = new char[] { ';' };


                lstQueryList.Add(line.Split(delimiters, StringSplitOptions.None));
                while (!reader.EndOfStream)
                {
                    line = reader.ReadLine();
                    lstQueryList.Add(line.Split(delimiters, StringSplitOptions.None));
                }
                reader.Close();
            }

            return lstQueryList;
        }

        public static void CreateUDV()
        {
            string formId, itemId, colId, queryName;
            int queryId, controlType, lRetCode = 0;
            Recordset oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordSet = AccessMatrixEngine.GetAllConfiguration();
             while (!oRecordSet.EoF)
            {
                
                formId = oRecordSet.Fields.Item("U_FormId").Value.ToString();
                itemId = oRecordSet.Fields.Item("U_ItemId").Value.ToString();
                colId = oRecordSet.Fields.Item("U_ColumnId").Value.ToString();
                queryName = oRecordSet.Fields.Item("U_QueryName").Value.ToString();
                controlType = Convert.ToInt32(oRecordSet.Fields.Item("U_ControlType").Value.ToString());
                
                if (formId == "UDO_FT_UserProject")
                {
                    string zz = "asd";
                }

                if (queryName == "")
                {
                    oRecordSet.MoveNext();
                    continue;
                }

                queryId = GetQueryId(queryName);

                ClearUDVIfExist(formId, itemId, (controlType == 1 ? "-1" : colId));//If control is textbox, SAP return the colUID value as "1". But UDV only works as "-1".

                FormattedSearches oFormattedSearches = (FormattedSearches)oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);
                
                oFormattedSearches.FormID = formId;
                oFormattedSearches.ItemID = itemId;
                oFormattedSearches.ColumnID = (controlType == 1 ? "-1" : colId); //If control is textbox, SAP return the colUID value as "1". But UDV only works as "-1".
                oFormattedSearches.Action = BoFormattedSearchActionEnum.bofsaQuery;
                oFormattedSearches.QueryID = queryId;

                lRetCode = oFormattedSearches.Add();

                if (lRetCode != 0)
                {
                    string errMsg = AddOnUtilities.oCompany.GetLastErrorDescription();
                    AddOnUtilities.MsgBoxWrapper(errMsg, MsgBoxType.B1StatusBar, BoMessageTime.bmt_Short, true);
                }
                oRecordSet.MoveNext();
            }
        }

        public static void ClearUDVIfExist(string formId, string itemId, string colId)
        {
            string strQuery = "select \"IndexID\", \"ItemID\", \"ColID\" from CSHS WHERE \"FormID\" = '" + formId + "' AND \"ItemID\" = '" + itemId + "' AND \"ColID\" = '" + colId + "'";
            Recordset oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            FormattedSearches oFormattedSearches = (FormattedSearches)oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);

            try
            {
                oRecordSet.DoQuery(strQuery);

                if (oRecordSet.RecordCount > 0)
                {
                    int formattedSearchKey = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                    oFormattedSearches.GetByKey(formattedSearchKey);
                    oFormattedSearches.Remove();
                }
            }
            catch (Exception ex)
            {
                MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormattedSearches);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static void ClearAllUDV()
        {
            string strQuery = "select \"IndexID\", \"ItemID\", \"ColID\" from CSHS";
            Recordset oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            FormattedSearches oFormattedSearches = (FormattedSearches)oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);

            try
            {
                oRecordSet.DoQuery(strQuery);

                while (!oRecordSet.EoF)
                {
                    int formattedSearchKey = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
                    oFormattedSearches.GetByKey(formattedSearchKey);
                    oFormattedSearches.Remove();
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormattedSearches);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static int GetQueryId(string QueryName)
        {
            string strQuery = "SELECT \"IntrnalKey\" FROM OUQR WHERE \"QName\" = '" + QueryName + "'";
            Recordset rs = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(strQuery);

            if (rs.RecordCount > 0)
            {
                return Convert.ToInt32(rs.Fields.Item(0).Value.ToString());
            }
            return -1;
        }
    }
}
