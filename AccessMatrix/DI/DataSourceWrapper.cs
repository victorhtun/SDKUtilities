using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AccessMatrix.Core;

namespace AccessMatrix.DI
{
    public static class DataSourceWrapper
    {
        public static String GetValue(ref SAPbouiCOM.Form oForm, String tableName, String fieldIndex, int rowIndex = 0)
        {
            try
            {
                SAPbouiCOM.DBDataSource dbDataSource = oForm.DataSources.DBDataSources.Item(tableName);
                return dbDataSource.GetValue(fieldIndex, rowIndex);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            return String.Empty;
        }

        public static void SetValue(ref SAPbouiCOM.Form oForm, String tableName, String newValue, String fieldIndex, int rowIndex = 0)
        {
            try
            {
                SAPbouiCOM.DBDataSource dbDataSource = oForm.DataSources.DBDataSources.Item(tableName);
                dbDataSource.SetValue(fieldIndex, rowIndex, newValue);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
        }
    }
}
