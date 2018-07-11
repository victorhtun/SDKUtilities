using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Configuration;
using SAPbobsCOM;
using SAPbouiCOM;

namespace SDKUtilities
{
    public static class AddOnUtilities
    {
        public static int IRetCode;
        public static int IErrCode;
        public static String sErrMsg;

        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbouiCOM.Application oApplication = null;

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
                // D. (To-Do) Register UDO
                // E. (To-Do) UI: Create Menus

                //Menu MyMenu = new Menu();
                //MyMenu.AddMenuItems();

                //oApplication.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                
                //oApplication.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                //oApplication.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }


        public static void Connect()
        {
            switch (SDKUtilities.Properties.Settings.Default.ConnectionType.ToString().ToUpper())
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

                DIErrorHandler("Connection via SSO.");
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
                String connectionString = Environment.GetCommandLineArgs().GetValue(1).ToString();
                // 2. Connect to UI API Object
                uiApi.Connect(connectionString);
                // License for your Add On (your IP)
                // Add-On Identifier
                // 3. Get the B1 Application Object
                oApplication = uiApi.GetApplication();
                // 4. (To-Do) Delegate other sub component (e.g. AppEvent, etc.)
                //oApp4AppEvent = oApplication;

                //oApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(EventHandlerApp.AppEventHandler);
                // 5. (To-Do) Setup Event filters

                MsgBoxWrapper("Connect Via UI API.");
            }
            catch (Exception ex)
            {
                MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
        }

        public static void ConnectViaDI()
        {
            // to do - change the parameters value needs to come from appsettings
            ConnectViaDI("CEL0035", BoDataServerTypes.dst_MSSQL2014, "SBODemoGB", "manager", "1234", "CEL0035:30000", BoSuppLangs.ln_English);
        }

        private static void ConnectViaDI(String dbServer,
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

                // 3. Connect
                IRetCode = oCompany.Connect();
                // 4. Error Handling
                DIErrorHandler("Connection to DI API.");
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

        public static void MsgBoxWrapper(String msg, MsgBoxType msgBoxType = MsgBoxType.B1MsgBox)
        {
            if (msgBoxType == MsgBoxType.B1MsgBox)
            {
                oApplication.MessageBox(msg);
            }
            else
            {
                oApplication.SetStatusBarMessage(msg);
            }
        }

        public enum MsgBoxType
        {
            WinMsgBox = 0,
            B1MsgBox = 1,
            B1StatusBar = 2
        };
    }
}
