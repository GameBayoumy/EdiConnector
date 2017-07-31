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

        public ProcessMessage(string _filePath, string _fileName)
        {
            filePath = _filePath;
            fileName = _fileName;
        }

        public void execute()
        {
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
    }
}
