using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AccessMatrix.DI;
using AccessMatrix.Enum;

namespace AccessMatrix.Core
{
    public static class GenericEventHandler
    {
        public static void RegisterEventHandler()
        {
            //events handled by AddOnUtilities.oApplication_AppEvent
            AddOnUtilities.oApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(oApplication_AppEvent);
            //events handled by AddOnUtilities.oApplication_MenuEvent
            AddOnUtilities.oApplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(oApplication_MenuEvent);
            //events handled by AddOnUtilities.oApplication_ItemEvent
            AddOnUtilities.oApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(oApplication_ItemEvent);
            //events handled by AddOnUtilities.oApplication_ProgressBarEvent
            AddOnUtilities.oApplication.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(oApplication_ProgressBarEvent);
            //events handled by AddOnUtilities.oApplication_StatusBarEvent
            AddOnUtilities.oApplication.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(oApplication_StatusBarEvent);
        }


        private static void oApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

            // ********************************************************************************
            //  the following are the events sent by the application
            //  (Ignore aet_ServerTermination)
            //  in order to implement your own code upon each of the events
            //  place you code instead of the matching message box statement
            // ********************************************************************************


            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:

                    //AddOnUtilities.MsgBoxWrapper("A Shut Down Event has been caught");

                    // **************************************************************
                    // 
                    //  Take care of terminating your AddOn application
                    // 
                    // **************************************************************

                    System.Windows.Forms.Application.Exit();


                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:

                    //AddOnUtilities.MsgBoxWrapper("A Company Change Event has been caught");

                    // **************************************************************
                    //  Check the new company name, if your add on was not meant for
                    //  the new company terminate your AddOn
                    //     If oApplication.Company.Name Is Not "Company1" then
                    //          Close
                    //     End If
                    // **************************************************************

                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:

                    //AddOnUtilities.MsgBoxWrapper("A Languge Change Event has been caught");

                    break;
            }
        }



        private static void oApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {

            // ********************************************************************************
            //  in order to activate your own forms instead of SAP Business One system forms
            //  process the menu event by your self
            //  change BubbleEvent to False so that SAP Business One won't process it
            // ********************************************************************************
            BubbleEvent = true;
            if (pVal.BeforeAction == true)
            {

                //oApplication.SetStatusBarMessage("Menu item: " + pVal.MenuUID + " sent an event BEFORE SAP Business One processes it.", SAPbouiCOM.BoMessageTime.bmt_Long, true);

                //  to stop SAP Business One from processing this event
                //  unmark the following statement

                //  BubbleEvent = False

            }
            else
            {

                //oApplication.SetStatusBarMessage("Menu item: " + pVal.MenuUID + " sent an event AFTER SAP Business One processes it.", SAPbouiCOM.BoMessageTime.bmt_Long, true);

            }

        }


        private static void oApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //For UDOs, FormType is 0 and FormUID starts with UDO.
            if (pVal.FormType != 0 || FormUID.Contains("UDO"))
            {
                
                SAPbouiCOM.BoEventTypes EventEnum = 0;
                EventEnum = pVal.EventType;

                // To prevent an endless loop of MessageBoxes,
                // we'll not notify et_FORM_ACTIVATE and et_FORM_LOAD events
                //if ( ( EventEnum != SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE ) & ( EventEnum != SAPbouiCOM.BoEventTypes.et_FORM_LOAD ) )
                if(pVal.BeforeAction)
                {
                    string input = "";

                    //if FormType is 0, it's UDO and use FormType as FormUID (Variable)
                    string formType = pVal.FormType.ToString() == "0" ? FormUID : pVal.FormType.ToString();

                    if (EventEnum == SAPbouiCOM.BoEventTypes.et_VALIDATE || EventEnum == BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbobsCOM.Recordset rsConfigData = AccessMatrixEngine.GetConfiguration(formType, pVal.ItemUID, pVal.ColUID, (int)pVal.EventType);
                        string accessMatrixKey = String.Empty, controlType = String.Empty;

                        if (rsConfigData.RecordCount > 0)
                        {
                            accessMatrixKey = rsConfigData.Fields.Item("U_AccessMatrixKey").Value.ToString();
                            controlType = rsConfigData.Fields.Item("U_ControlType").Value.ToString();
                            string controlName = rsConfigData.Fields.Item("U_ControlName").Value.ToString();
                            string itemId = rsConfigData.Fields.Item("U_ItemId").Value.ToString();
                            string colId = rsConfigData.Fields.Item("U_ColumnId").Value.ToString();
                            string result = AccessMatrixEngine.CheckIfValidOrNot(rsConfigData);

                            input = AccessMatrixEngine.GetInputValue(itemId, colId, controlType);

                            if (input != "" && result != "")
                            {
                                BubbleEvent = false;
                                AddOnUtilities.oApplication.ActivateMenuItem("7425");
                                return;
                            }

                            if (result != "")
                            {
                                BubbleEvent = false;
                                AddOnUtilities.MsgBoxWrapper(result, Enum.MsgBoxType.B1StatusBar, BoMessageTime.bmt_Short, true);
                                return;
                            }



                            //Set Names from Codes in UDOs.
                            string relatedColumn = rsConfigData.Fields.Item("U_RelatedColumn").Value.ToString();
                            string relatedItem = rsConfigData.Fields.Item("U_RelatedItem").Value.ToString();

                            if (input != "" && pVal.FormType.ToString() == "0")
                            {
                                if (relatedColumn != "" && relatedItem != "")
                                {
                                    string strName = AccessMatrixEngine.GetNameByCode(input, accessMatrixKey);
                                    AccessMatrixEngine.SetValueToMatrix(relatedItem, relatedColumn, strName);
                                }

                            }

                            BubbleEvent = true;
                            return;
                        }
                    }
                    else if (EventEnum == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.CharPressed == 9)
                    {
                        SAPbobsCOM.Recordset rsConfigData = AccessMatrixEngine.GetConfiguration(formType, pVal.ItemUID, pVal.ColUID, (int)pVal.EventType);
                        string accessMatrixKey = String.Empty, controlType = String.Empty;

                        if(rsConfigData.RecordCount > 0)
                        {

                            accessMatrixKey = rsConfigData.Fields.Item("U_AccessMatrixKey").Value.ToString();
                            controlType = rsConfigData.Fields.Item("U_ControlType").Value.ToString();
                            string controlName = rsConfigData.Fields.Item("U_ControlName").Value.ToString();
                            string itemId = rsConfigData.Fields.Item("U_ItemId").Value.ToString();
                            string colId = rsConfigData.Fields.Item("U_ColumnId").Value.ToString();
                            string result = AccessMatrixEngine.CheckIfValidOrNot(rsConfigData);

                            input = AccessMatrixEngine.GetInputValue(itemId, colId, controlType);

                            if (input == "")
                            {
                                BubbleEvent = false;
                                AddOnUtilities.oApplication.ActivateMenuItem("7425");
                                return;
                            }
                            if(result != "")
                            {
                                BubbleEvent = false;
                                AddOnUtilities.oApplication.ActivateMenuItem("7425");
                                return;
                            }
                            BubbleEvent = true;
                            return;
                        }

                    }
                    else if(EventEnum == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                         SAPbobsCOM.Recordset rsConfigData = AccessMatrixEngine.GetConfigurationForLineLoop(formType, pVal.ItemUID, pVal.ColUID, (int)pVal.EventType);
                        int recordCount = rsConfigData.RecordCount;
                        string accessMatrixKey = String.Empty, controlType = String.Empty;

                        while (!rsConfigData.EoF)
                        {
                            //To Do : Update following statement to be dynamic.
                            // I = 1 (Item), S = 2 (Service)

                            //Some Document do not have Item or Service Type. Those controls are always ComboBox.
                            //Check if it is ComboBox and if it's ComboBox, assign ControlType as 3 which is not Item or Service.
                            dynamic tmpControl = AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item("3").Specific;
                            string documentType = "";
                            if (tmpControl is SAPbouiCOM.ComboBox)
                            {
                                documentType = ((SAPbouiCOM.ComboBox)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item("3").Specific).Value;
                                documentType = (documentType == "I" ? "1" : "2");
                            }
                            else
                            {
                                documentType = "3";
                            }
                            
                            
                            string configDocumentType = rsConfigData.Fields.Item("U_DocType").Value.ToString();

                            if (documentType != configDocumentType)
                            {
                                rsConfigData.MoveNext();
                                continue;
                            }

                            //if(rsConfigData.Fields.Item("U_DocType").Value.ToString() != documentType && !rsConfigData.EoF)
                            //{
                            //rsConfigData.MoveNext();
                            //}
                            string result = AccessMatrixEngine.CheckIfValidOrNotAllLine(rsConfigData, documentType);
                            rsConfigData.MoveNext();
                            if (result != "")
                            {
                                BubbleEvent = false;
                                AddOnUtilities.MsgBoxWrapper(result, Enum.MsgBoxType.B1StatusBar, BoMessageTime.bmt_Short, true);
                                return;
                            }
                        }
                    }
                }
            }
            BubbleEvent = true;
        }

        //private static string GetInputValue(ItemEvent pVal, string controlType, string relatedCol = "")
        //{
        //    string input = String.Empty;
        //    int rowIndex;
        //    SAPbouiCOM.Matrix oMatrix = null;
        //    switch (controlType)
        //    {
        //        case "1": //Textbox
        //            input = ((SAPbouiCOM.EditText)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(pVal.ItemUID).Specific).Value;
        //            break;
        //        case "2": //Checkbox (Only for BP Selection Window)
        //            string colIdForDescription = (relatedCol == "" ? pVal.ColUID : relatedCol);
        //            oMatrix = (SAPbouiCOM.Matrix)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(pVal.ItemUID).Specific;
        //            rowIndex = oMatrix.GetCellFocus().rowIndex;
        //            input = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(colIdForDescription).Cells.Item(rowIndex).Specific).Value;
        //            break;
        //        case "3": //Dropdown List
        //            break;
        //        case "4": //Matrix
        //            oMatrix = (SAPbouiCOM.Matrix)AddOnUtilities.oApplication.Forms.ActiveForm.Items.Item(pVal.ItemUID).Specific;
        //            rowIndex = oMatrix.GetCellFocus().rowIndex;
        //            input = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(rowIndex).Specific).Value;
        //            break;
        //        default:
        //            break;
        //    }
        //    return input;
        //}


        private static void oApplication_ProgressBarEvent(ref SAPbouiCOM.ProgressBarEvent pVal, out bool BubbleEvent)
        {
            SAPbouiCOM.BoProgressBarEventTypes EventEnum = 0;
            EventEnum = pVal.EventType;
            BubbleEvent = true;
            //AddOnUtilities.MsgBoxWrapper("The event " + EventEnum.ToString() + " has been sent", 1, "Ok", "", "");
        }


        private static void oApplication_StatusBarEvent(string Text, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {
            //AddOnUtilities.MsgBoxWrapper(@"Status bar event with message: """ + Text + @""" has been sent", 1, "Ok", "", "");
        }
    }
}
