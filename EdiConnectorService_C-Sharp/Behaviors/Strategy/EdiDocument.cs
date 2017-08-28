using System;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{

    /// <summary>
    /// Generalized EDI document which can have different document types containing different implementations for the defined functions in this class.
    /// By setting the document type, this object will implement the functions of the set document type.
    /// When calling upon a function with this object it will refer to the functions defined in the document type implementation.
    /// User can define specific implementations by creating a new object and inheriting "IEdiDocumentType".
    /// </summary>
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

        /// <summary>
        /// Gets the type of the document.
        /// </summary>
        /// <returns>The document type</returns>
        public IEdiDocumentType GetDocumentType()
        {
            return documentType;
        }

        /// <summary>
        /// Gets the name of the document type.
        /// </summary>
        /// <returns>The document type name</returns>
        public string GetDocumentTypeName()
        {
            return documentType.TypeName;
        }

        /// <summary>
        /// Reads the XML data.
        /// </summary>
        /// <param name="_xMessages">The x messages.</param>
        /// <param name="_exception">The exception.</param>
        /// <returns></returns>
        public Object ReadXMLData(XElement _xMessages, out string _exception)
        {
            return documentType.ReadXMLData(_xMessages, out _exception);
        }

        /// <summary>
        /// Saves specific document data object to SAP.
        /// </summary>
        /// <param name="_dataObject">The data object.</param>
        /// <param name="_connectedServer">The connected server.</param>
        /// <param name="_exception">The exception.</param>
        public void SaveToSAP(Object _dataObject, string _connectedServer, out string _exception)
        {
            documentType.SaveToSAP(_dataObject, _connectedServer, out _exception);
        }
    }
}
