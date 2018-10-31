using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessMatrix.Core
{
    public static class Metadata
    {
        public static void CreateTableForCustomizeForms()
        {
            Dictionary<string, string> dictValidValues = new Dictionary<string, string>();

            //--------------- Branch Dimension Table. ---------------//
            //Header
            Metadata.CreateUDT("OBDM", "BranchDimension", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            Metadata.CreateUDF("@OBDM", "Branch", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25);

            //Details
            Metadata.CreateUDT("BDM1", "Dimension 1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@BDM1", "Dimension1", "Dimension 1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDF("@BDM1", "Dimension1Name", "Dimension 1 Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30);

            Metadata.CreateUDT("BDM2", "Dimension 2", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@BDM2", "Dimension2", "Dimension 2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDF("@BDM2", "Dimension2Name", "Dimension 2 Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30);

            Metadata.CreateUDT("BDM3", "Dimension 3", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@BDM3", "Dimension3", "Dimension 3", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDF("@BDM3", "Dimension3Name", "Dimension 3 Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30);

            Metadata.CreateUDT("BDM4", "Dimension 4", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@BDM4", "Dimension4", "Dimension 4", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDF("@BDM4", "Dimension4Name", "Dimension 4 Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30);

            Metadata.CreateUDT("BDM5", "Dimension 5", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@BDM5", "Dimension5", "Dimension 5", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8);
            Metadata.CreateUDF("@BDM5", "Dimension5Name", "Dimension 5 Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30);
            //--------------- End of Branch Dimension Table. ---------------//

            //--------------- Business Partners by User Table. ---------------//
            //Header
            Metadata.CreateUDT("OUBP", "UserBusinessPartners", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            Metadata.CreateUDF("@OUBP", "User", "User Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25);

            //Details
            Metadata.CreateUDT("UBP1", "UserBusinessPartnersDetails", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@UBP1", "GroupCode", "BP Group Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15);
            Metadata.CreateUDF("@UBP1", "GroupName", "BP Group Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20);
            //--------------- End of Business Partners by User Table. ---------------//

            //--------------- Projects by User Table. ---------------//
            //Header
            Metadata.CreateUDT("OUPR", "UserProject", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            Metadata.CreateUDF("@OUPR", "User", "UserCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25);

            //Details
            Metadata.CreateUDT("UPR1", "UserProjectDetails", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@UPR1", "Project", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20);
            Metadata.CreateUDF("@UPR1", "ProjectName", "ProjectName", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            //--------------- End of Projects by User Table. ---------------//


            //--------------- Chart of Accounts by User Table. ---------------//
            //Header
            Metadata.CreateUDT("OUAC", "UserChartofAccounts", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            Metadata.CreateUDF("@OUAC", "User", "UserCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25);

            //Details
            Metadata.CreateUDT("UAC1", "UserChartofAccounts", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            Metadata.CreateUDF("@UAC1", "AccountCode", "G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20);
            Metadata.CreateUDF("@UAC1", "AccountName", "G/L Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            //--------------- End of Chart of Accounts by User Table. ---------------//


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
            Metadata.CreateUDF("@OCFL", "QueryName", "Query Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250);
            Metadata.CreateUDF("@OCFL", "IsCustomized", "Required Customization", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1);
            Metadata.CreateUDF("@OCFL", "RelatedItem", "Related Item Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
            Metadata.CreateUDF("@OCFL", "RelatedColumn", "Related Column Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
            Metadata.CreateUDF("@OCFL", "RelatedItem2", "Related Item Id2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
            Metadata.CreateUDF("@OCFL", "RelatedColumn2", "Related Column Id2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50);
            Metadata.CreateUDF("@OCFL", "RelatedQuery", "RelatedQuery", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250);
        }

        public static void CreateUDT(String udtName, String udtDesc, SAPbobsCOM.BoUTBTableType udtType)
        {
            if (!CheckTableExists(udtName))
            {
                //SDK -> UserTableMD Object -> Fields Required
                SAPbobsCOM.IUserTablesMD oUDTMD = null;
                try
                {

                    // 1. Get Company Object
                    oUDTMD = AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                    // 2. Set Table Properties
                    oUDTMD.TableName = udtName;
                    oUDTMD.TableDescription = udtDesc;
                    oUDTMD.TableType = udtType;
                    // 3. Add

                    AddOnUtilities.IRetCode = oUDTMD.Add();
                    // 4. Error Handling
                    AddOnUtilities.DIErrorHandler(String.Format("UDT {0} Created.", udtName));
                }
                catch (Exception ex)
                {
                    AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace, Enum.MsgBoxType.B1StatusBar, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                finally
                {
                    //Important - release COM Object
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDTMD);
                    GC.Collect();
                    oUDTMD = null;
                }
            }
        }

        public static void CreateUDF(String udt, //can put business object as well
            String udfName, String udfDesc, SAPbobsCOM.BoFieldTypes udfType, SAPbobsCOM.BoFldSubTypes udfSubType, int udfEditSize, Dictionary<string, string> udfValidValues = null, SAPbobsCOM.BoYesNoEnum udfMandatory = SAPbobsCOM.BoYesNoEnum.tNO)
        {
            if (!CheckFieldExists(udt, udfName))
            {
                SAPbobsCOM.UserFieldsMD oUDFMD = null;

                try
                {
                    oUDFMD = AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                    oUDFMD.TableName = udt;
                    oUDFMD.Name = udfName;
                    oUDFMD.Description = udfDesc;
                    oUDFMD.Type = udfType;
                    oUDFMD.SubType = udfSubType;
                    oUDFMD.EditSize = udfEditSize;
                    oUDFMD.Mandatory = udfMandatory;

                    //foreach(var udfValidValue in udfValidValues)
                    //{
                    //    oUDFMD.ValidValues.Value = udfValidValue.Key;
                    //    oUDFMD.ValidValues.Value = udfValidValue.Key;
                    //    oUDFMD.ValidValues.Add();
                    //}

                    AddOnUtilities.IRetCode = oUDFMD.Add();

                    AddOnUtilities.DIErrorHandler(String.Format("UDF {0} Created.", udfName));
                }
                catch (Exception ex)
                {
                    AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
                }
                finally
                {
                    //Important - release COM Object
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDFMD);
                    GC.Collect();
                    oUDFMD = null;
                }
            }
        }

        private static bool CheckUDOExists(String udoName)
        {
            SAPbobsCOM.UserObjectsMD oUdtMD = null;
            bool blnFlag = false;

            try
            {
                oUdtMD = AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                if (oUdtMD.GetByKey(udoName))
                    blnFlag = true;
            }
            catch (Exception ex)
            {
                blnFlag = false;
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdtMD);
                oUdtMD = null;
                GC.Collect();
            }

            return blnFlag;
        }

        private static bool CheckTableExists(String tableName)
        {
            SAPbobsCOM.IUserTablesMD oUdtMD = null;
            bool blnFlag = false;

            try
            {
                tableName = tableName.Replace("@", "");
                oUdtMD = AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                if (oUdtMD.GetByKey(tableName))
                    blnFlag = true;
            }
            catch (Exception ex)
            {
                blnFlag = false;
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdtMD);
                oUdtMD = null;
                GC.Collect();
            }

            return blnFlag;
        }

        private static bool CheckFieldExists(String tableName, String fieldName)
        {
            SAPbobsCOM.IUserFieldsMD oUserFieldsMD = null;
            bool blnFlag = false;

            try
            {
                fieldName = fieldName.Replace("U_", "");
                oUserFieldsMD = AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                int fieldID = GetFieldIDByName(tableName, fieldName);
                
                if (oUserFieldsMD.GetByKey(tableName, fieldID))
                    blnFlag = true;
            }
            catch (Exception ex)
            {
                blnFlag = false;
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
            }

            return blnFlag;
        }

        private static int GetFieldIDByName(String tableName, String fieldName)
        {
            int index = -1;
            SAPbobsCOM.Recordset ors = null;

            try
            {
                ors = AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (AddOnUtilities.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    ors.DoQuery("SELECT \"FieldID\" FROM \"CUFD\" WHERE \"TableID\" = '" + tableName + "' AND \"AliasID\" = '" + fieldName + "'");
                else
                    ors.DoQuery("SELECT FieldID FROM CUFD WHERE TableID = '" + tableName + "' AND AliasID = '" + fieldName + "'");

                if (!ors.EoF)
                    index = ors.Fields.Item("FieldID").Value;
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
                ors = null;
                GC.Collect();
            }

            return index;
        }
    }
}
