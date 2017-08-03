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

        public ProcessMessage(string _connectedServer, string _filePath, string _fileName)
        {
            connectedServer = _connectedServer;
            filePath = _filePath;
            fileName = _fileName;
        }

        public void execute()
        {
            XDocument xDoc = XDocument.Load(filePath + fileName);
            XElement xMessages = xDoc.Element("Messages");
            EdiDocument ediDocument = new EdiDocument();
            Object ediDocumentData = new Object();
            recordCode = AddIncomingXmlMessage(connectedServer, filePath, fileName, "Processing..", "Loaded new document", DateTime.Now);

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
            //UpdateIncomingXmlMessage(connectedServer, "Processing..", "Set document type to: " + ediDocument.GetDocumentType().ToString());

            // Reads the XML Data for the specified document type
            ediDocumentData = ediDocument.ReadXMLData(xMessages);
            //UpdateIncomingXmlMessage(connectedServer, "Processing..", "Read document with type: " + ediDocument.GetDocumentType().ToString());

            // Save the data object for the specified document type to SAP
            ediDocument.SaveToSAP(ediDocumentData, connectedServer);
            //UpdateIncomingXmlMessage(connectedServer, "Saved.", "Saved document with type: " + ediDocument.GetDocumentType().ToString());
        }

        private string AddIncomingXmlMessage(string _connectedServer, string _filePath, string _fileName, string _status, string _logMessage, DateTime _createDate)
        {
            SAPbobsCOM.UserTable oUDT = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.UserTables.Item("0_SWS_EDI_LOG");
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            SAPbobsCOM.Attachments2 oAtt = (SAPbobsCOM.Attachments2)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2));

            try
            {

                oAtt.Lines.SourcePath = _filePath;
                oAtt.Lines.FileName = _fileName;
                oAtt.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;

                if (oAtt.Add() == 0)
                {
                    int attEnt = int.Parse(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetNewObjectKey());
                    oUDT.UserFields.Fields.Item("U_XML_ATTACHMENT").Value = attEnt;
                    //DeliveryV.UserFields.Fields.Item("U_ReadyForInvoice").Value = "Y";
                }
                else
                {
                    ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                    EventLogger.getInstance().EventError("Error adding attachment to UDF: " + errMsg);
                }

                oUDT.UserFields.Fields.Item("U_STATUS").Value = _status;
                oUDT.UserFields.Fields.Item("U_LOG_MESSAGE").Value = _logMessage;
                oUDT.UserFields.Fields.Item("U_CREATE_DATE").Value = _createDate;

                if (oUDT.Add() == 0)
                {
                    oRs.DoQuery(@"SELECT Max(""Code"") FROM ""@0_SWS_EDI_LOG""");
                    recordCode = oRs.Fields.Item(0).Value.ToString();
                    EventLogger.getInstance().EventInfo("Succesfully added incoming xml message to UDT: " + oUDT.TableName);
                    return recordCode;
                }
                else
                {
                    EventLogger.getInstance().EventError("Error adding items to UDT: " + ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription());
                    return recordCode = null;
                }
            }
            catch
            {
                ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                EventLogger.getInstance().EventError("Error adding items to UDF: " + errMsg);
                return recordCode = null;
            }
            finally
            {
                EdiConnectorService.ClearObject(oAtt);
                EdiConnectorService.ClearObject(oUDT);
                EdiConnectorService.ClearObject(oRs);
            }
        }

        private void UpdateIncomingXmlMessage(string _connectedServer, string _status, string _logMessage)
        {
            SAPbobsCOM.UserTable oUDT;
            oUDT = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.UserTables.Item("0_SWS_EDI_LOG");

            oUDT.GetByKey(recordCode);

            oUDT.UserFields.Fields.Item("U_STATUS").Value = _status;
            oUDT.UserFields.Fields.Item("U_LOG_MESSAGE").Value = _logMessage;

            if (oUDT.Update() == 0)
            {
                EventLogger.getInstance().EventInfo("Succesfully updated incoming xml message to UDT: " + oUDT.TableName);
            }
            else
            {
                EventLogger.getInstance().EventError("Error updating items to UDT: " + ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription());
            }
        }
    }
}
