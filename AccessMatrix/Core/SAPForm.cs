using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace AccessMatrix.Core
{
    class SAPForm
    {
        public SAPbouiCOM.Form CreateFormViaXML(String xmlData)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                SAPbouiCOM.FormCreationParams oFCP = AddOnUtilities.oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                oFCP.XmlData = xmlData;

                // Option 1 - to load the file
                //AddOn.oApplication.LoadBatchActions(xmlData);
                // Option 2 - to load the file
                oForm = AddOnUtilities.oApplication.Forms.AddEx(oFCP);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oForm;
        }

        public SAPbouiCOM.Form CreateFormXmlFile(String filePath)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                if (!File.Exists(filePath))
                {
                    AddOnUtilities.MsgBoxWrapper(String.Format("File does not exist.", filePath));
                    return null;
                }

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(filePath);
                oForm = CreateFormViaXML(xmlDoc.InnerXml);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return oForm;
        }

        public SAPbouiCOM.Form CreateQCForm()
        {
            SAPbouiCOM.Form oForm = null;
            oForm = CreateFormXmlFile("C:\\Users\\CEL0035\\Desktop\\Customized Form\\QC_SDK_2018.xml");
            ManageSeries(ref oForm, "Item_3", "@TH_OQC", "1_U_E");

            //Guidelines : Events of Company Change/ Language
            //Detect user Language
            return oForm;
        }

        public void SetValueViaDBS(ref SAPbouiCOM.Form oForm, String table, String field, Object value)
        {
            try
            {
                SAPbouiCOM.DBDataSource oDBS = oForm.DataSources.DBDataSources.Item(table);
                oDBS.SetValue(field, 0, value.ToString());
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
        }

        public void ManageSeries(ref SAPbouiCOM.Form oForm, String seriesUID, String headerTable, String docNumUID)
        {
            SAPbouiCOM.ComboBox cbxSeries = oForm.Items.Item(seriesUID).Specific;

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                cbxSeries.ValidValues.LoadSeries(oForm.BusinessObject.Type, SAPbouiCOM.BoSeriesMode.sf_Add);
                cbxSeries.Select("Primary", SAPbouiCOM.BoSearchKey.psk_ByDescription);

                //Getting the next number
                int nextDocNum = oForm.BusinessObject.GetNextSerialNumber("Primary");
                SetValueViaDBS(ref oForm, headerTable, "DocNum", nextDocNum);
            }
            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            {
                cbxSeries.ValidValues.LoadSeries(oForm.BusinessObject.Type, SAPbouiCOM.BoSeriesMode.sf_View);
                SetValueViaDBS(ref oForm, headerTable, "Series", String.Empty);
                oForm.Items.Item(docNumUID).Enabled = true;
            }
        }
    }
}
