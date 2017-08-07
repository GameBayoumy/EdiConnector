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
        /// Gets or sets the file path.
        /// </summary>
        public string MessagesFilePath { get; set; }

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

                MessagesFilePath = xmlNode["MessagesFilePath"].InnerText;
                if (System.IO.Directory.Exists(MessagesFilePath))
                {
                    if (!System.IO.Directory.Exists(MessagesFilePath + EdiConnectorData.getInstance().sProcessedDirName))
                    {
                        System.IO.Directory.CreateDirectory(MessagesFilePath + EdiConnectorData.getInstance().sProcessedDirName);
                    }
                }

                EventLogger.getInstance().EventInfo("Set connection - Server: " + Company.Server);
            }
            catch (Exception e)
            {
                EventLogger.getInstance().EventError("Error setting connection: " + e.Message);
            }
        }

        /// <summary>
        /// Connects this instance.
        /// </summary>
        /// <returns></returns>
        public override bool Connect()
        {
            try
            {
                if (Company.Connect() == 0)
                {
                    EventLogger.getInstance().EventInfo("Connected to SAP - Server: " + Company.Server);
                    return ConnectedToSAP = true;
                }
                else if (Company.Connected == true)
                {
                    EventLogger.getInstance().EventInfo("Already connected to SAP - Server: " + Company.Server);
                    return ConnectedToSAP = true;
                }
                else
                {
                    EventLogger.getInstance().EventError(Company.GetLastErrorDescription() + Company.Server);
                    return ConnectedToSAP = false;
                }
            }
            catch (Exception e)
            {
                EventLogger.getInstance().EventError(e.Message);
                return ConnectedToSAP = false;
            }
        }

        /// <summary>
        /// Disconnects this instance.
        /// </summary>
        /// <returns></returns>
        public override bool Disconnect()
        {
            try
            {
                if (Company.Connected == true)
                {
                    Company.Disconnect();
                    EventLogger.getInstance().EventInfo("Disconnected to SAP - Server: " + Company.Server);
                    ConnectedToSAP = false;
                }
                return ConnectedToSAP;
            }
            catch (Exception e)
            {
                EventLogger.getInstance().EventError(e.Message);
                return ConnectedToSAP = false;
            }
        }

        /// <summary>
        /// Sends the mail notification through SBO Mailer.
        /// </summary>
        /// <param name="_subject">The subject.</param>
        /// <param name="_body">The body.</param>
        /// <param name="_mailAddress">The mail address.</param>
        public void SendMailNotification(string _subject, string _body, string _mailAddress)
        {
            Messages objMsg = (Messages)Company.GetBusinessObject(BoObjectTypes.oMessages);
            objMsg.Subject = _subject;
            objMsg.MessageText = _body;
            objMsg.Recipients.Add();
            objMsg.Recipients.SetCurrentLine(0);
            objMsg.Recipients.UserCode = Company.UserName;
            objMsg.Recipients.NameTo = Company.UserName;
            objMsg.Recipients.UserType = BoMsgRcpTypes.rt_InternalUser;
            objMsg.Recipients.SendInternal = BoYesNoEnum.tNO;
            objMsg.Recipients.SendEmail = BoYesNoEnum.tYES;
            objMsg.Recipients.EmailAddress = _mailAddress;
            objMsg.Priority = BoMsgPriorities.pr_Normal;
            //objMsg.Attachments.Add();
            //objMsg.Attachments.Item(0).FileName = Attachment;

            if (objMsg.Add() != 0)
                EventLogger.getInstance().EventInfo("Error sending mail notification from: " + Company.Server);
            else
                EventLogger.getInstance().EventInfo("Message send from: " + Company.Server);

            EdiConnectorService.ClearObject(objMsg);
        }
    }
}
