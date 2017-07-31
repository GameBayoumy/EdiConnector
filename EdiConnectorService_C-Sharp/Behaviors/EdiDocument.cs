using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Collections;

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

        public Object ReadXMLData(XElement _xMessages)
        {
            return documentType.ReadXMLData(_xMessages);
        }

        public void SaveToSAP(Object _dataObject)
        {

        }
    }
}
