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
        XmlNode xmlNode;

        public SAPConnection()
        {

        }

        /// <summary>
        /// Sets the specified XML node.
        /// </summary>
        /// <param name="_xmlNode">The XML node.</param>
        public override void Set(XmlNode _xmlNode)
        {
            this.xmlNode = _xmlNode;
            Company.DbServerType = BoDataServerTypes.dst_HANADB;
            Company.DbUserName = xmlNode["DbUserName"].InnerText;
            Company.DbPassword = xmlNode["DbPassword"].InnerText;
            Company.Server = xmlNode["Server"].InnerText;
            Company.LicenseServer = xmlNode["LicenceServer"].InnerText;
            Company.CompanyDB = xmlNode["CompanyDB"].InnerText;
            if (xmlNode["Test"].InnerText == "Y")
                Company.CompanyDB = "TEST_" + Company.CompanyDB;
            Company.UserName = xmlNode["UserName"].InnerText;
            Company.Password = xmlNode["PassWord"].InnerText;
        }

        public override bool Connect()
        {
            try
            {
                if (Company.Connect() == 0)
                {
                    ConnectedToSAP = true;
                    return ConnectedToSAP;
                }
                else if (Company.Connected == true)
                {
                    ConnectedToSAP = true;
                    return ConnectedToSAP;
                }
                else
                {
                    ConnectedToSAP = false;
                    EventLogger.getInstance().EventInfo("");
                    return ConnectedToSAP;
                }
            }
            catch (Exception e)
            {
                ConnectedToSAP = false;
                return ConnectedToSAP;
            }
        }

        public override bool Disconnect()
        {
            try
            {
                Company.Disconnect();
                ConnectedToSAP = false;
                return ConnectedToSAP;
            }
            catch
            {
                return ConnectedToSAP;
            }
        }

        /// <summary>
        /// Gets or sets the company.
        /// </summary>
        /// <value>
        /// The company.
        /// </value>
        public SAPbobsCOM.Company Company { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether [connected to sap].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [connected to sap]; otherwise, <c>false</c>.
        /// </value>
        public bool ConnectedToSAP { get; set; }
    }
}
