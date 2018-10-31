using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using SBODI_Server;
using AccessMatrix.Properties;
using AccessMatrix.Core;

namespace AccessMatrix.DI
{
    public class DIServerWrapper
    {
        private INode node = null;
        private String sSessionID = String.Empty;

        #region [DI Server Login]
        public void DIServerLoginSample()
        {
            Login("CEL0035", "SBODemoGB", "dst_MSSQL2014", "manager", "1234", "ln_English", "CEL0035:30000");
        }

        public String Login(String DbServer, String DbName, String DbType, String User, String Password, String Language, String License)
        {
            String sessionID = String.Empty;

            try
            {
                // 1. Get Soap Request + Pass Parameters
                String soapLoginRequest = String.Format(Resources.DIServerLoginSOAP, DbServer, DbName, DbType, User, Password, Language, License);
                // 2. Send to DI Server Node
                String soapLoginResponse = Interact(soapLoginRequest);

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(soapLoginResponse);

                XmlNode xmlNode = xmlDoc.SelectSingleNode("//*[local-name()='SessionID']");

                if (xmlNode == null)
                {
                    // Error
                    AddOnUtilities.MsgBoxWrapper("DI Server Login failed.");
                }
                else
                {
                    // Success
                    sessionID = xmlNode.InnerText;
                    sSessionID = sessionID;
                    AddOnUtilities.MsgBoxWrapper(string.Format("DI Server Login Success. {0}", sessionID));
                }
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return sessionID;
        }

        #endregion

        public struct DIServerResult
        {
            public bool Success;
            //if success, msg return retkey of BO (e.g. CardCode, DocEntry, etc.)

            public String Message;

            public DIServerResult(bool bSuccess, String sMsg)
            {
                Success = bSuccess;
                Message = sMsg;
            }
        }

        public DIServerResult AddObject(String sessionID, String commandID, string sBOM)
        {
            DIServerResult result = new DIServerResult();
            try
            {
                String soapResponse = Interact(String.Format(Resources.AddObjectSoapRequest, sessionID, commandID, sBOM));

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(soapResponse);
                XmlNode node = xmlDoc.SelectSingleNode("//*[local-name()='RetKey']");

                if (node == null)
                {
                    //error
                    AddOnUtilities.MsgBoxWrapper("DI Server - Add Obj Failed.");
                    result.Success = false;
                }
                else
                {
                    AddOnUtilities.MsgBoxWrapper(String.Format("DI Server - Add Obj Success. Return Key: {0}.", node.InnerText));
                    result.Success = true;
                }
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }

            return result;
        }

        #region [Add Business Partner Sample]

        //public void DIServerAddBPSample()
        //{
        //    DIServerLoginSample();
        //    DIServerAddBP(sSessionID, "C20000", "Added from DI Server");
        //}

        //public void DIServerAddBP(String sessionID, String bpCode, String bpName)
        //{
        //    try
        //    {
        //        String bpXMLBOM = String.Format(Resources.BPSoapRequest, bpCode, bpName);
        //        DIServerResult response = AddObject(sessionID, "Add BP", bpXMLBOM);
        //    }
        //    catch (Exception ex)
        //    {
        //        AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
        //    }
        //}

        #endregion
            

        public INode GetDIServerNode()
        {
            if (node == null) // there can be multiple nodes (Load balancer)
                node = new Node();

            return node;
        }

        public String Interact(String soapRequest)
        {
            String soapResponse = String.Empty;

            try
            {
                soapResponse = GetDIServerNode().Interact(soapRequest);
            }
            catch (Exception ex)
            {
                AddOnUtilities.MsgBoxWrapper(ex.Message + " " + ex.StackTrace);
            }
            return soapResponse;
        }

        public string BatchInteract(String soapRequest)
        {
            return GetDIServerNode().BatchInteract(soapRequest);
        }
    }

}
