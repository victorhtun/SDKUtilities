using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAPbobsCOM;
using AccessMatrix.Core;

namespace AccessMatrix
{
    public partial class Test : Form
    {
        public Test()
        {
            InitializeComponent();
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            AddOnUtilities.Connect();

        }

        private void btnCreateUDT_Click(object sender, EventArgs e)
        {
            CreateTableForCustomizeForms();

            //List<string> childTables = new List<string>();
            //childTables.Add("UDM1");
            //childTables.Add("UDM2");
            //childTables.Add("@UDM3");
            //childTables.Add("@UDM4");
            //childTables.Add("@UDM5");

            //CreateUDO("UserDimension", "UserDimension", BoUDOObjType.boud_MasterData, "OUDM", childTables, BoYesNoEnum.tYES, "UserDimension", 3328, 1);
        }

        public void CreateTableForCustomizeForms()
        {
            Dictionary<string, string> dictValidValues = new Dictionary<string, string>();

            //Header
            Metadata.CreateUDT("OUDM", "UserDimension", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            Metadata.CreateUDF("@OUDM", "User", "User Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25);

            //Details
            Metadata.CreateUDT("UDM1", "Dimension 1 Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@UDM1", "Dimension1", "Dimension 1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDT("UDM2", "Dimension 2 Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@UDM2", "Dimension2", "Dimension 2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDT("UDM3", "Dimension 3 Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@UDM3", "Dimension3", "Dimension 3", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDT("UDM4", "Dimension 4 Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@UDM4", "Dimension4", "Dimension 4", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDT("UDM5", "Dimension 5 Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@UDM5", "Dimension5", "Dimension 5", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);

            //Configuration Table
            Metadata.CreateUDT("OCFL", "Configuration Form List", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
            Metadata.CreateUDF("@OCFL", "FormId", "Form Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200);
            Metadata.CreateUDF("@OCFL", "FormName", "Form Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
            Metadata.CreateUDF("@OCFL", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 2);
            Metadata.CreateUDF("@OCFL", "AccessMatrixKey", "Access Matrix Key", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 2);
            Metadata.CreateUDF("@OCFL", "ControlName", "Control Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            Metadata.CreateUDF("@OCFL", "ItemId", "Item Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
            Metadata.CreateUDF("@OCFL", "ColumnId", "Column Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
            Metadata.CreateUDF("@OCFL", "IsSingle", "Is Single", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 2);
            Metadata.CreateUDF("@OCFL", "ControlType", "Control Type", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 2);
            Metadata.CreateUDF("@OCFL", "EventType", "Event Type", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 2);
            Metadata.CreateUDF("@OCFL", "IsCustomized", "Required Customization", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1);
        }


        private void CreateUDO(string udoName, string udoCode, BoUDOObjType udoType, string headerTableName, List<String> childTables, BoYesNoEnum isMenuItem, string menuCaption, int fatherMenuId, int menuPosition, BoYesNoEnum canCancel = BoYesNoEnum.tYES, BoYesNoEnum canClose = BoYesNoEnum.tNO, BoYesNoEnum canCreateDefaultForm = BoYesNoEnum.tYES, BoYesNoEnum canDelete = BoYesNoEnum.tYES, BoYesNoEnum canFind = BoYesNoEnum.tYES, BoYesNoEnum canLog = BoYesNoEnum.tYES, BoYesNoEnum canYearTransfer = BoYesNoEnum.tNO)
        {
            try
            {
                if (!(AddOnUtilities.UDOExist(udoName)))
                {
                    UserObjectsMD UDO = (UserObjectsMD)AddOnUtilities.oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

                    //Set Services
                    // TODO: use optional params for properties below
                    UDO.CanCancel = canCancel;
                    UDO.CanClose = canClose;
                    UDO.CanCreateDefaultForm = canCreateDefaultForm;
                    UDO.CanDelete = canDelete;
                    UDO.CanFind = canFind;
                    UDO.CanLog = canLog;
                    UDO.CanYearTransfer = canYearTransfer;

                    UDO.TableName = headerTableName;
                    UDO.FormColumns.SonNumber = childTables.Count;
                    UDO.EnableEnhancedForm = BoYesNoEnum.tYES; // To show Header Line Style
                    UDO.Code = udoCode;
                    UDO.Name = udoName;
                    UDO.ObjectType = udoType;

                    //Display columns
                    UDO.FormColumns.FormColumnAlias = "Code";
                    UDO.FormColumns.FormColumnDescription = "Code";
                    UDO.FormColumns.Editable = BoYesNoEnum.tYES;
                    UDO.FormColumns.Add();

                    UDO.FormColumns.FormColumnAlias = "U_User";
                    UDO.FormColumns.FormColumnDescription = "User";
                    UDO.FormColumns.Editable = BoYesNoEnum.tYES;
                    UDO.FormColumns.Add();

                    //UDO.FormColumns.FormColumnAlias = "U_Dimension1";
                    //UDO.FormColumns.FormColumnDescription = "Dimension 1";
                    //UDO.FormColumns.Add();

                    //UDO.FormColumns.FormColumnAlias = "U_Dimension2";
                    //UDO.FormColumns.FormColumnDescription = "Dimension 2";
                    //UDO.FormColumns.Add();

                    //Columns to be used in find mode
                    UDO.FindColumns.ColumnAlias = "Code";
                    UDO.FindColumns.ColumnDescription = "Code";
                    UDO.FindColumns.Add();

                    UDO.FindColumns.ColumnAlias = "U_User";
                    UDO.FindColumns.ColumnDescription = "User";
                    UDO.FindColumns.Add();

                    // Set UDO to have a menu 
                    UDO.MenuItem = isMenuItem;
                    UDO.MenuCaption = menuCaption;
                    // Set father and gnment of menu item. 
                    UDO.FatherMenuID = fatherMenuId;
                    UDO.Position = 1;
                    // Set UDO menu UID 
                    UDO.MenuUID = "UserDimension";

                    //Link with child tables.
                    foreach (var childTable in childTables)
                    {
                        UDO.ChildTables.TableName = childTable;
                        UDO.ChildTables.Add();
                    }

                    int RetCode = UDO.Add();
                    if (RetCode != 0)
                    {
                        string errMsg = AddOnUtilities.oCompany.GetLastErrorDescription();
                        MessageBox.Show("Failed to add UDO");
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(UDO);
                }
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
                SAPbouiCOM.EventForm e = null;
                
            }
        }

        private void btnCreateQueries_Click(object sender, EventArgs e)
        {
            UserQueries oUserQueries = (SAPbobsCOM.UserQueries)AddOnUtilities.oCompany.GetBusinessObject(BoObjectTypes.oUserQueries);
            QueryCategories oQueryCategories = (SAPbobsCOM.QueryCategories)AddOnUtilities.oCompany.GetBusinessObject(BoObjectTypes.oQueryCategories);
            int lRetCode = 0, queryCategoryId = 0;

            string strQuery = "select \"CategoryId\" from \"OQCN\" where \"CatName\" = 'Customization'";
            SAPbobsCOM.Recordset oRecordSet = ((SAPbobsCOM.Recordset)(AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            oRecordSet.DoQuery(strQuery);

            bool isQueryCategoryExist = Convert.ToBoolean(oRecordSet.Fields.Item(0).Value);
            if(!isQueryCategoryExist)
            {
                oQueryCategories.Name = "Customization";
                oQueryCategories.Permissions = "YYYYYYYYYYYYYYY";
                lRetCode = oQueryCategories.Add();

                if(lRetCode != 0)
                {
                    string errMsg = AddOnUtilities.oCompany.GetLastErrorDescription();
                    MessageBox.Show("Failed to create Query Category.");
                    MessageBox.Show(errMsg);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oQueryCategories);
            }

            List<string[]> lstQueryList = AddOnUtilities.ReadQueries("C:\\Users\\LynnHtet\\Desktop\\UserQueryList.txt");
            queryCategoryId = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
            foreach (string[] query in lstQueryList)
            {
                
                oUserQueries.QueryDescription = query[0];
                oUserQueries.Query = query[1];
                oUserQueries.QueryType = UserQueryTypeEnum.uqtWizard;
                oUserQueries.QueryCategory = queryCategoryId;
                lRetCode = oUserQueries.Add();

                if (lRetCode != 0)
                {
                    string errMsg = AddOnUtilities.oCompany.GetLastErrorDescription();
                    MessageBox.Show("Failed to create Query Category.");
                    MessageBox.Show(errMsg);
                }
            }
        }

        private void btnAddUDVs_Click(object sender, EventArgs e)
        {
            AddOnUtilities.CreateUDV();
        }

        private void btnConnectUI_Click(object sender, EventArgs e)
        {
            AddOnUtilities.ConnectViaUI();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddOnUtilities.ConnectViaDI();
        }
    }
}
