using System;
using System.Linq;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    public class ProcessMessage : ICommand
    {
        string filePath;
        string fileName;
        string connectedServer;
        XDocument xDoc;

        public ProcessMessage(string _connectedServer, string _fileName)
        {
            connectedServer = _connectedServer;
            fileName = _fileName;
        }

        public void execute()
        {
            filePath = ConnectionManager.getInstance().GetConnection(connectedServer).MessagesFilePath;
            xDoc = XDocument.Load(filePath + fileName);
            XElement xMessages = xDoc.Element("Messages");
            EdiDocument ediDocument = new EdiDocument();
            Object ediDocumentData = new Object();
            string recordReference = EventLogger.getInstance().CreateSAPLogMessage(connectedServer, fileName, xDoc, "Loaded new document: " + fileName, "Processing..");
            EdiConnectorData.getInstance().sRecordReference = recordReference;
            if (System.IO.File.Exists(filePath + fileName))
            {
                System.IO.File.Copy((filePath + fileName), (filePath + EdiConnectorData.getInstance().sProcessedDirName + @"\" + fileName), true);
                System.IO.File.Delete((filePath + fileName));
            }

            // Checks which kind of document type gets through the system
            try
            {
                if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "3").Count() > 0)
                {
                    if(ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "demo")
                        ediDocument.SetDocumentType(new OrderDocument());
                    else if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "plastica")
                        ediDocument.SetDocumentType(new PlasticaOrderDoc());
                }
                else if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "5").Count() > 0)
                {
                    if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "demo")
                        ediDocument.SetDocumentType(new OrderResponseDocument());
                    //else if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "plastica")
                        //ediDocument.SetDocumentType(new PlasticaResponseDoc());
                }
                else if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "8").Count() > 0)
                {
                    if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "demo")
                        ediDocument.SetDocumentType(new InvoiceDocument());
                    //else if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "plastica")
                        //ediDocument.SetDocumentType(new PlasticaInvoiceDoc());
                }

                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Set document type to: " + ediDocument.GetDocumentTypeName(), "Processing..");
            }
            catch (Exception e)
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Error setting document type with XML MessageType: " + xDoc.Element("MessageType").Value.ToString() + ". EXCEPTION: " + e.Message, "Error!");
                EventLogger.getInstance().EventError("Server: " + connectedServer + ". Error setting document type with XML MessageType: " + xDoc.Element("MessageType").Value.ToString() + ". EXCEPTION: " + e.Message);
            }

            // Reads the XML Data for the specified document type
            ediDocumentData = ediDocument.ReadXMLData(xMessages, out string exceptionRead);
            if (string.IsNullOrWhiteSpace(exceptionRead))
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Read document with type: " + ediDocument.GetDocumentTypeName(), "Processing..");
            else
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Error reading document with type: " + ediDocument.GetDocumentTypeName() + " ERROR: " + exceptionRead + " XML node probably missing/incorrect!!!", "Error!");
                EventLogger.getInstance().EventError("Server: " + connectedServer + ". Error reading document with type: " + ediDocument.GetDocumentTypeName() + " ERROR: " + exceptionRead + " XML node probably missing / incorrect!!!");
            }

            // Save the data object for the specified document type to SAP
            if(ediDocumentData != null)
            {
                ediDocument.SaveToSAP(ediDocumentData, connectedServer, out string exceptionSave);
                if(string.IsNullOrWhiteSpace(exceptionSave))
                    EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Saved document " + fileName + " with document type: " + ediDocument.GetDocumentTypeName(), "Processed.");
                else
                {
                    EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Error saving document: " + fileName + " with document type: " + ediDocument.GetDocumentTypeName() + " ERROR: " + exceptionSave, "Error!");
                    EventLogger.getInstance().EventError("Server: " + connectedServer + ". Error saving document: " + fileName + " with document type: " + ediDocument.GetDocumentTypeName() + " ERROR: " + exceptionSave);
                }
            }
            else
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Error saving document: " + fileName + " data object is empty!", "Error!");
                EventLogger.getInstance().EventError("Server: " + connectedServer + ". Error saving document: " + fileName + " data object is empty!");
            }
        }
    }
}
