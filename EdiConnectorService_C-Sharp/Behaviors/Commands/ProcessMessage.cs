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
            string logRecordReference = EventLogger.getInstance().CreateSAPLogMessage(connectedServer, fileName, xDoc, "Loaded new document: " + fileName, "Processing..");
            ediDocument.SetLogRecordReference(logRecordReference);

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

                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, logRecordReference, "Set document type to: " + ediDocument.GetDocumentTypeName(), "Processing..");
            }
            catch (Exception e)
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, logRecordReference, "Error setting document type with XML MessageType: " + xDoc.Element("MessageType").Value.ToString() + ". EXCEPTION: " + e.Message, "Error!");
                EventLogger.getInstance().EventError("Server: " + connectedServer + ". Error setting message - Error setting document type with XML MessageType: " + xDoc.Element("MessageType").Value.ToString() + ". EXCEPTION: " + e.Message);
            }


            // Reads the XML Data for the specified document type
            ediDocumentData = ediDocument.ReadXMLData(xMessages, out Exception exR);
            if (exR != null)
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, logRecordReference, "Error reading document with type: " + ediDocument.GetDocumentTypeName() + " ERROR: " + exR.Message + " XML node probably missing/incorrect!!!", "Error!");
                EventLogger.getInstance().EventError("Server: " + connectedServer + ". Error reading message - Error reading document with type: " + ediDocument.GetDocumentTypeName() + " ERROR: " + exR.Message + " XML node probably missing / incorrect!!!");
            }
            else
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, logRecordReference, "Read document with type: " + ediDocument.GetDocumentTypeName(), "Processing..");

            // Save the data object for the specified document type to SAP
            ediDocument.SaveToSAP(ediDocumentData, connectedServer, out Exception exS);
            if(exS != null)
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, logRecordReference, "Saving document " + fileName + " with document type: " + ediDocument.GetDocumentTypeName() + " ERROR: " + exS.Message, "Error!");
                EventLogger.getInstance().EventError("Server: " + connectedServer + ". Error saving document - Error saving document " + fileName + " with document type: " + ediDocument.GetDocumentTypeName() + " ERROR: " + exS.Message);
            }
            else
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, logRecordReference, "Saved document " + fileName + " with document type: " + ediDocument.GetDocumentTypeName(), "Processed.");
        }
    }
}
