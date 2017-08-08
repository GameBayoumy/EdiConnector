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

                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Set document type to: " + ediDocument.GetDocumentType().ToString(), "Processing..");
            }
            catch (Exception e)
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Error setting document type with XML MessageType: " + xDoc.Element("MessageType").Value.ToString() + ". EXCEPTION: " + e.Message, "Error!");
                EventLogger.getInstance().EventError("Error setting message - Error setting document type with XML MessageType: " + xDoc.Element("MessageType").Value.ToString() + ". EXCEPTION: " + e.Message);
            }


            // Reads the XML Data for the specified document type
            ediDocumentData = ediDocument.ReadXMLData(xMessages, out Exception exR);
            if (exR != null)
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Error reading document with type: " + ediDocument.GetDocumentType().ToString() + " ERROR: " + exR.Message + " XML node probably missing/incorrect!!!", "Error!");
                EventLogger.getInstance().EventError("Error reading message - Error reading document with type: " + ediDocument.GetDocumentType().ToString() + " ERROR: " + exR.Message + " XML node probably missing / incorrect!!!");
            }
            else
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Read document with type: " + ediDocument.GetDocumentType().ToString(), "Processing..");

            // Save the data object for the specified document type to SAP
            ediDocument.SaveToSAP(ediDocumentData, connectedServer, out Exception exS);
            if(exS != null)
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Saving document " + fileName + " with document type: " + ediDocument.GetDocumentType().ToString() + " ERROR: " + exS.Message, "Error!");
                EventLogger.getInstance().EventError("Error saving document - Error saving document " + fileName + " with document type: " + ediDocument.GetDocumentType().ToString() + " ERROR: " + exS.Message);
            }
            else
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Saved document " + fileName + " with document type: " + ediDocument.GetDocumentType().ToString(), "Processed.");
        }
    }
}
