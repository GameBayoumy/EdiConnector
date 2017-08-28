using System;
using SAPbobsCOM;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class SAPConnection : AConnection
    {
        /// <summary>
        /// Gets or sets the edi profile.
        /// </summary>
        /// <value>
        /// The edi profile.
        /// </value>
        public string EdiProfile { get; set; }

        /// <summary>
        /// Gets or sets the company.
        /// </summary>
        /// <value>
        /// The company.
        /// </value>
        public SAPbobsCOM.Company Company { get; set; }

        /// <summary>
        /// Gets or sets the file path.
        /// </summary>
        /// <value>
        /// The messages file path.
        /// </value>
        public string MessagesFilePath { get; set; }


        /// <summary>
        /// Gets or sets the udf file path.
        /// </summary>
        /// <value>
        /// The udf file path.
        /// </value>
        public string UdfFilePath { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [connected to sap].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [connected to sap]; otherwise, <c>false</c>.
        /// </value>
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
        /// <param name="_xEle">The XML node.</param>
        public override void Set(XElement _xEle)
        {
            try
            {
                XElement xEle = _xEle;
                EdiProfile = xEle.Element("EdiProfile").Value;
                Company.Server = xEle.Element("Server").Value;
                Company.LicenseServer = xEle.Element("LicenceServer").Value;
                Company.UserName = xEle.Element("Username").Value;
                Company.Password = xEle.Element("Password").Value;
                Company.CompanyDB = xEle.Element("CompanyDB").Value;
                if (xEle.Element("Test").Value == "Y")
                    Company.CompanyDB = "TEST_" + Company.CompanyDB;
                if (xEle.Element("DbServerType").Value == "HANA")
                    Company.DbServerType = BoDataServerTypes.dst_HANADB;
                else if (xEle.Element("DbServerType").Value == "2005")
                    Company.DbServerType = BoDataServerTypes.dst_MSSQL2005;
                else if (xEle.Element("DbServerType").Value == "2008")
                    Company.DbServerType = BoDataServerTypes.dst_MSSQL2008;
                else if (xEle.Element("DbServerType").Value == "2012")
                    Company.DbServerType = BoDataServerTypes.dst_MSSQL2012;
                else
                    Company.DbServerType = BoDataServerTypes.dst_MSSQL;
                Company.DbUserName = xEle.Element("DbUsername").Value;
                Company.DbPassword = xEle.Element("DbPassword").Value;
                Company.UseTrusted = true;

                MessagesFilePath = xEle.Element("MessagesFilePath").Value;
                if (System.IO.Directory.Exists(MessagesFilePath))
                {
                    if (!System.IO.Directory.Exists(MessagesFilePath + EdiConnectorData.getInstance().sProcessedDirName))
                    {
                        System.IO.Directory.CreateDirectory(MessagesFilePath + EdiConnectorData.getInstance().sProcessedDirName);
                    }
                }
                else
                {
                    EventLogger.getInstance().EventWarning($"Server: {Company.Server}. Messages File Path not found! {MessagesFilePath}");
                }

                UdfFilePath = xEle.Element("UdfFilePath").Value;
                if (!System.IO.File.Exists(UdfFilePath))
                {
                    EventLogger.getInstance().EventWarning($"Server: {Company.Server}. Udf File Path not found! {UdfFilePath}");
                }

                EventLogger.getInstance().EventInfo($"Server: {Company.Server}. Connection set");
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
                    string serverName = Company.Server;
                    Company.Disconnect();
                    EventLogger.getInstance().EventInfo("Disconnected to SAP - Server: " + serverName);
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
                EventLogger.getInstance().EventInfo($"Server: {Company.Server}. Error sending mail notification to: {_mailAddress}");
            else
                EventLogger.getInstance().EventInfo($"Server: {Company.Server}. Message send to: {_mailAddress}");

            EdiConnectorService.ClearObject(objMsg);
        }
    }
}
