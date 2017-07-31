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

            ediDocumentData = ediDocument.ReadXMLData(xMessages);

        }
    }
}
