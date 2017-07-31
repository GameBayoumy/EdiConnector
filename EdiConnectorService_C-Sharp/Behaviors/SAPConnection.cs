using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using SAPbobsCOM;

namespace EdiConnectorService_C_Sharp
{
    class SAPConnection : AConnection
    {
        /// <summary>
        /// Gets or sets the company.
        /// </summary>
        public SAPbobsCOM.Company Company { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [connected to sap].
        /// </summary>
        public bool ConnectedToSAP { get; set; }

        public SAPConnection()
        {
            Company = new Company();
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="SAPConnection"/> class.
        /// </summary>
        ~SAPConnection()
        {
            EdiConnectorService.ClearObject(Company);
        }

        /// <summary>
        /// Sets the specified XML node.
        /// </summary>
        /// <param name="_xmlNode">The XML node.</param>
        public override void Set(XmlNode _xmlNode)
        {
            try
            {
                XmlNode xmlNode = _xmlNode;
                Company.Server = xmlNode["Server"].InnerText;
                Company.LicenseServer = xmlNode["LicenceServer"].InnerText;
                Company.UserName = xmlNode["Username"].InnerText;
                Company.Password = xmlNode["Password"].InnerText;
                Company.CompanyDB = xmlNode["CompanyDB"].InnerText;
                if (xmlNode["Test"].InnerText == "Y")
                    Company.CompanyDB = "TEST_" + Company.CompanyDB;
                if (xmlNode["DbServerType"].InnerText == "HANA")
                    Company.DbServerType = BoDataServerTypes.dst_HANADB;
                else if (xmlNode["DbServerType"].InnerText == "2005")
                    Company.DbServerType = BoDataServerTypes.dst_MSSQL2005;
                else if (xmlNode["DbServerType"].InnerText == "2008")
                    Company.DbServerType = BoDataServerTypes.dst_MSSQL2008;
                else if (xmlNode["DbServerType"].InnerText == "2012")
                    Company.DbServerType = BoDataServerTypes.dst_MSSQL2012;
                else
                    Company.DbServerType = BoDataServerTypes.dst_MSSQL;
                Company.DbUserName = xmlNode["DbUsername"].InnerText;
                Company.DbPassword = xmlNode["DbPassword"].InnerText;
                Company.UseTrusted = true;
                
                EventLogger.getInstance().EventInfo("Set connection - Server: " + Company.Server);
            }
            catch (Exception e)
            {
                EventLogger.getInstance().EventError("Error setting connection: " + e.Message);
            }
        }

        public override bool Connect()
        {
            try
            {
                if (Company.Connect() == 0)
                {
                    EventLogger.getInstance().EventInfo("Connected to SAP - Server: " + Company.Server);
                    ConnectedToSAP = true;
                    return ConnectedToSAP;
                }
                else if (Company.Connected == true)
                {
                    EventLogger.getInstance().EventInfo("Already connected to SAP - Server: " + Company.Server);
                    ConnectedToSAP = true;
                    return ConnectedToSAP;
                }
                else
                {
                    EventLogger.getInstance().EventError(Company.GetLastErrorDescription() + Company.Server);
                    ConnectedToSAP = false;
                    return ConnectedToSAP;
                }
            }
            catch (Exception e)
            {
                EventLogger.getInstance().EventError(e.Message);
                ConnectedToSAP = false;
                return ConnectedToSAP;
            }
        }

        public override bool Disconnect()
        {
            try
            {
                Company.Disconnect();
                EventLogger.getInstance().EventInfo("Disconnected to SAP - Server: " + Company.Server);
                ConnectedToSAP = false;
                return ConnectedToSAP;
            }
            catch
            {
                return ConnectedToSAP;
            }
        }
    }
}
