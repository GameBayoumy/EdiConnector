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
        string recordCode;
        XDocument xDoc;

        public ProcessMessage(string _connectedServer, string _filePath, string _fileName)
        {
            connectedServer = _connectedServer;
            filePath = _filePath;
            fileName = _fileName;
        }

        public void execute()
        {
            xDoc = XDocument.Load(filePath + fileName);
            XElement xMessages = xDoc.Element("Messages");
            EdiDocument ediDocument = new EdiDocument();
            Object ediDocumentData = new Object();
            AddIncomingXmlMessage("Processing..", "Loaded new document: " + fileName, DateTime.Now);

            // Checks which kind of document type gets through the system
            try
            {
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

                UpdateIncomingXmlMessage("Processing..", "Set document type to: " + ediDocument.GetDocumentType().ToString());
            }
            catch (Exception e)
            {
                UpdateIncomingXmlMessage("Error!", "Error setting document type with XML MessageType: " + xDoc.Element("MessageType").Value.ToString() + ". Exception: " + e.Message);
                EventLogger.getInstance().EventError("Error setting message - Error setting document type with XML MessageType: " + xDoc.Element("MessageType").Value.ToString() + ". Exception: " + e.Message);
            }


            // Reads the XML Data for the specified document type
            ediDocumentData = ediDocument.ReadXMLData(xMessages, out Exception exR);
            if (exR != null)
            {
                UpdateIncomingXmlMessage("Error!", "Error reading document with type: " + ediDocument.GetDocumentType().ToString() + " Error: " + exR.Message + " XML node probably missing/incorrect!!!");
                EventLogger.getInstance().EventError("Error reading message - Error reading document with type: " + ediDocument.GetDocumentType().ToString() + " Error: " + exR.Message + " XML node probably missing / incorrect!!!");
            }
            else
                UpdateIncomingXmlMessage("Processing..", "Read document with type: " + ediDocument.GetDocumentType().ToString());

            // Save the data object for the specified document type to SAP
            ediDocument.SaveToSAP(ediDocumentData, connectedServer, out Exception exS);
            if(exS != null)
            {
                UpdateIncomingXmlMessage("Error!", "Saving document " + fileName + " with document type: " + ediDocument.GetDocumentType().ToString() + " Error: " + exS.Message);
                EventLogger.getInstance().EventError("Error saving document - Error saving document" + fileName + " with document type: " + ediDocument.GetDocumentType().ToString() + " Error: " + exS.Message);
            }
            else
                UpdateIncomingXmlMessage("Processed.", "Saved document " + fileName + " with document type: " + ediDocument.GetDocumentType().ToString());
        }

        private void AddIncomingXmlMessage(string _status, string _logMessage, DateTime _createDateTime)
        {
            SAPbobsCOM.UserTable oUDT = ConnectionManager.getInstance().GetConnection(connectedServer).Company.UserTables.Item("0_SWS_EDI_LOG");
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)(ConnectionManager.getInstance().GetConnection(connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

            try
            {
                oUDT.Name = fileName;
                oUDT.UserFields.Fields.Item("U_XML_ATTACHMENT").Value = xDoc.ToString();
                oUDT.UserFields.Fields.Item("U_STATUS").Value = _status;
                oUDT.UserFields.Fields.Item("U_LOG_MESSAGE").Value = _logMessage;
                oUDT.UserFields.Fields.Item("U_CREATE_DATE").Value = _createDateTime.Date;
                oUDT.UserFields.Fields.Item("U_CREATE_TIME").Value = _createDateTime.ToShortTimeString();

                if (oUDT.Add() == 0)
                {
                    oRs.DoQuery(@"SELECT Max(""Code"") FROM ""@0_SWS_EDI_LOG""");
                    EventLogger.getInstance().EventInfo("Succesfully added incoming xml message to UDT: " + oUDT.TableName);
                    recordCode = oRs.Fields.Item(0).Value.ToString();
                }
                else
                {
                    EventLogger.getInstance().EventError("Error adding items to UDT: " + ConnectionManager.getInstance().GetConnection(connectedServer).Company.GetLastErrorDescription());
                    recordCode = null;
                }
            }
            catch
            {
                ConnectionManager.getInstance().GetConnection(connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                EventLogger.getInstance().EventError("Error adding items to UDF: " + errMsg);
                recordCode = null;
            }
            finally
            {
                EdiConnectorService.ClearObject(oUDT);
                EdiConnectorService.ClearObject(oRs);
            }
        }

        private void UpdateIncomingXmlMessage(string _status, string _logMessage)
        {
            SAPbobsCOM.UserTable oUDT = ConnectionManager.getInstance().GetConnection(connectedServer).Company.UserTables.Item("0_SWS_EDI_LOG");

            oUDT.GetByKey(recordCode);

            oUDT.UserFields.Fields.Item("U_STATUS").Value = _status;
            oUDT.UserFields.Fields.Item("U_LOG_MESSAGE").Value = _logMessage;

            if (oUDT.Update() != 0)
            {
                EventLogger.getInstance().EventError("Error updating items to UDT: " + ConnectionManager.getInstance().GetConnection(connectedServer).Company.GetLastErrorDescription());
            }

            EdiConnectorService.ClearObject(oUDT);
        }
    }
}
