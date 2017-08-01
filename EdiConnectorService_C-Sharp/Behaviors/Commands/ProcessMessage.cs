using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    public class ProcessMessage : ICommand
    {
        string filePath;
        string fileName;
        string connectedServer;
        string userTableCode;

        public ProcessMessage(string _connectedServer, string _filePath, string _fileName)
        {
            connectedServer = _connectedServer;
            filePath = _filePath;
            fileName = _fileName;
        }

        public void execute()
        {
            userTableCode = AddIncomingXmlMessage(connectedServer, filePath, fileName, "Processing...", "", DateTime.Now);
            XDocument xDoc = XDocument.Load(filePath + fileName);
            XElement xMessages = xDoc.Element("Messages");
            EdiDocument ediDocument = new EdiDocument();
            Object ediDocumentData = new Object();

            // Checks which kind of document type gets through the system
            if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "3").Count() > 0)
            {
                ediDocument.SetDocumentType(new OrderDocument());
            }
            else if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "5").Count() > 0)
            {
                ediDocument.SetDocumentType(new OrderResponseDocument());
            }
            else if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "8").Count() > 0)
            {
                ediDocument.SetDocumentType(new InvoiceDocument());
            }

            // Reads the XML Data for the specified document type
            ediDocumentData = ediDocument.ReadXMLData(xMessages);

            // Save the data object for the specified document type to SAP
            ediDocument.SaveToSAP(ediDocumentData);
        }

        private string AddIncomingXmlMessage(string _connectedServer, string _filePath, string _fileName, string _status, string _logMessage, DateTime _createDate)
        {
            SAPbobsCOM.UserTable oUDT;
            oUDT = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.UserTables.Item("0_SWS_EDI");

            try
            {
                oUDT.UserFields.Fields.Item("U_XML_FILE_PATH").Value = _filePath;
                oUDT.UserFields.Fields.Item("U_XML_FILE_NAME").Value = _fileName;
                oUDT.UserFields.Fields.Item("U_STATUS").Value = _status;
                oUDT.UserFields.Fields.Item("U_LOG_MESSAGE").Value = _logMessage;
                oUDT.UserFields.Fields.Item("U_CREATE_DATE").Value = _createDate;

                if (oUDT.Add() == 0)
                {
                    EventLogger.getInstance().EventInfo("Succesfully added incoming xml message to UDT: " + oUDT.Name);
                    return oUDT.Code;
                }
                else
                {
                    EventLogger.getInstance().EventError("Error adding items to UDT: " + ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription());
                    return null;
                }
            }
            catch
            {
                EventLogger.getInstance().EventError("Error adding items to UDF: " + ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription());
                return null;
            }
            finally
            {
                EdiConnectorService.ClearObject(oUDT);
            }
        }

        private void UpdateIncomingXmlMessage()
        {

        }
    }
}
