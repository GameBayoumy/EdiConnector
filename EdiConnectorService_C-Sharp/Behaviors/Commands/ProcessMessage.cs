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
        string recordReference;
        XDocument xDoc;
        EdiDocument ediDocument;
        Object ediDocumentData;

        // Constructor with connected server and file name as parameter
        public ProcessMessage(string _connectedServer, string _fileName)
        {
            connectedServer = _connectedServer;
            fileName = _fileName;
        }

        public void execute()
        {
            // Load incoming XML message from the messages file path
            filePath = ConnectionManager.getInstance().GetConnection(connectedServer).MessagesFilePath;
            xDoc = XDocument.Load(filePath + fileName);
            XElement xMessages = xDoc.Element("Messages");

            // Create SAP log message and save it's reference
            recordReference = EventLogger.getInstance().CreateSAPLogMessage(connectedServer, fileName, xDoc, "Loaded new document: " + fileName, "Processing..");
            EdiConnectorData.GetInstance().RecordReference = recordReference;
            
            // Create generalized EdiDocument and data object
            ediDocument = new EdiDocument();
            ediDocumentData = new Object();

            // Check if the incoming XML message exists to prevent crashes
            if (System.IO.File.Exists(filePath + fileName))
            {
                // Copy the xml message to the processed directory and delete the file from the messages file path
                System.IO.File.Copy((filePath + fileName), (filePath + EdiConnectorData.GetInstance().ProcessedDirName + @"\" + fileName), true);
                System.IO.File.Delete((filePath + fileName));
            }

            /// <summary>
            /// Checks which kind of document type gets through the system by looking for the <MessageType> value
            /// Value 3 = Sales order
            /// Value 5 = Sales order response
            /// Value 8 = Invoice
            /// 
            /// It will apply the correct document type according to the message type value
            /// But first it will check which EDI profile is being used by the connection 
            /// and then it will set the EDI document with the specified document type
            /// </summary>
            try
            {
                if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "3").Count() > 0)
                {
                    if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "norm")
                        ediDocument.SetDocumentType(new OrderDocument());
                    else if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "demo")
                        ediDocument.SetDocumentType(new DemoOrderDoc());
                    else if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "plastica")
                        ediDocument.SetDocumentType(new PlasticaOrderDoc());
                }
                else if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "5").Count() > 0)
                {
                    if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "norm")
                        ediDocument.SetDocumentType(new OrderResponseDocument());
                    //else if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "plastica")
                    //ediDocument.SetDocumentType(new PlasticaResponseDoc());
                }
                else if (xMessages.Elements().Where(x => x.Element("MessageType").Value == "8").Count() > 0)
                {
                    if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "norm")
                        ediDocument.SetDocumentType(new InvoiceDocument());
                    //else if (ConnectionManager.getInstance().GetConnection(connectedServer).EdiProfile == "plastica")
                    //ediDocument.SetDocumentType(new PlasticaInvoiceDoc());
                }

                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Set document type to: " + ediDocument.GetDocumentTypeName(), "Processing..");
            }
            catch (Exception e)
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, $"Error setting document type with XML MessageType: {xDoc.Element("MessageType").Value.ToString()}. EXCEPTION: {e.Message}", "Error!");
                EventLogger.getInstance().EventError($"Server: {connectedServer}. Error setting document type with XML MessageType: {xDoc.Element("MessageType").Value.ToString()}. EXCEPTION: {e.Message}");
            }


            // Reads the XML message for the specified document type and return the data object
            ediDocumentData = ediDocument.ReadXMLData(xMessages, out string exceptionRead);

            // Check if reading XML data method has caught an exception and log its error to SAP
            if (string.IsNullOrWhiteSpace(exceptionRead))
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, "Read document with type: " + ediDocument.GetDocumentTypeName(), "Processing..");
            else
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, $"Error reading document with type: {ediDocument.GetDocumentTypeName()} ERROR: {exceptionRead} XML node probably missing/incorrect!!!", "Error!");
                EventLogger.getInstance().EventError($"Server: {connectedServer}. Error reading document with type: {ediDocument.GetDocumentTypeName()} ERROR: {exceptionRead} XML node probably missing / incorrect!!!");
            }

            // Check if the data object is not empty
            if (ediDocumentData != null)
            {
                // Save the data object for the specified document type to SAP
                ediDocument.SaveToSAP(ediDocumentData, connectedServer, out string exceptionSave);

                // Check if saving data object has caught an exception and log its error to SAP
                if (string.IsNullOrWhiteSpace(exceptionSave))
                    EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, $"Saved document {fileName} with document type: {ediDocument.GetDocumentTypeName()}", "Processed.");
                else
                {
                    EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, $"Error saving document: {fileName} with document type: {ediDocument.GetDocumentTypeName()} ERROR: {exceptionSave}", "Error!");
                    EventLogger.getInstance().EventError($"Server: {connectedServer}. Error saving document: {fileName} with document type: {ediDocument.GetDocumentTypeName()} ERROR: {exceptionSave}");
                }
            }
            else
            {
                EventLogger.getInstance().UpdateSAPLogMessage(connectedServer, recordReference, $"Error saving document: {fileName} data object is empty!", "Error!");
                EventLogger.getInstance().EventError($"Server: {connectedServer}. Error saving document: {fileName} data object is empty!");
            }
        }
    }
}
