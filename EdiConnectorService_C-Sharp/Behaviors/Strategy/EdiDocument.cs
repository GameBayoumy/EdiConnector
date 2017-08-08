using System;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class EdiDocument
    {
        IEdiDocumentType documentType;

        public EdiDocument()
        {

        }

        /// <summary>
        /// Sets the type of the document.
        /// </summary>
        /// <param name="_documentType">Type of the document.</param>
        public void SetDocumentType(IEdiDocumentType _documentType)
        {
            this.documentType = _documentType;
        }

        public IEdiDocumentType GetDocumentType()
        {
            return documentType;
        }

        public Object ReadXMLData(XElement _xMessages, out Exception _ex)
        {
            return documentType.ReadXMLData(_xMessages, out _ex);
        }

        public void SaveToSAP(Object _dataObject, string _connectedServer, out Exception _ex)
        {
            documentType.SaveToSAP(_dataObject, _connectedServer, out _ex);
        }
    }
}
