using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AccessMatrix.Core;

namespace AccessMatrix.DI
{
    public class RecordSet
    {
        public void BrowserTop3CustomerSample()
        {
            String query = @"Select Top 3 CardCode as 'Code', CardName as 'Name', Balance as 'Balance' From OCRD Where CardType = 'C' Order by Balance DESC";

            try
            {
                Recordset oRS = ExecuteQuery(query);
                SAPbobsCOM.BusinessPartners oBP = AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                // set the RecordSet to Browser of BO
                oBP.Browser.Recordset = oRS;

                String bpCode = String.Empty;
                String bpName = String.Empty;
                Double balance = Double.MinValue;
                String msg = "Customer Code : {0}. BP Name: {1}. Balance: {2}";

                oBP.Browser.MoveFirst();
                while (!oBP.Browser.EoF)
                {
                    bpCode = oBP.CardCode;
                    bpName = oBP.CardName;
                    balance = oBP.CurrentAccountBalance; // read from DI API properties
                    AddOnUtilities.MsgBoxWrapper(String.Format(msg, bpCode, bpName, balance));
                    oBP.SaveXML(oBP.CardCode + ".xml");
                    oBP.Browser.MoveNext();
                }

            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
        }

        public void QueryTop3CustomersSample()
        {
            String query = @"Select Top 3 CardCode as 'Code', CardName as 'Name', Balance as 'Balance' From OCRD Where CardType = 'C' Order by Balance DESC";

            try
            {
                Recordset oRS = ExecuteQuery(query);

                // loop recordset
                if (oRS != null)
                {
                    String bpCode = String.Empty;
                    String bpName = String.Empty;
                    Double balance = Double.MinValue;
                    String msg = "Customer Code : {0}. BP Name: {1}. Balance: {2}";

                    while (!oRS.EoF)
                    {
                        bpCode = oRS.Fields.Item("Code").Value;
                        bpName = oRS.Fields.Item("Name").Value;
                        balance = oRS.Fields.Item("Balance").Value;
                        AddOnUtilities.MsgBoxWrapper(String.Format(msg, bpCode, bpName, balance));
                        oRS.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
        }

        public SAPbobsCOM.Recordset ExecuteQuery(String query)
        {
            SAPbobsCOM.Recordset oRS = null;

            try
            {
                oRS = AddOnUtilities.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(query);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            return oRS;
        }
    }
}
