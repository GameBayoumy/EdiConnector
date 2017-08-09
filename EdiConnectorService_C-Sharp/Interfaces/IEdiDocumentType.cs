using System;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    interface IEdiDocumentType
    {
        string TypeName { get; set; }

        Object ReadXMLData(XElement _xMessages, out Exception _ex);
        void SaveToSAP(Object _dataObject, string _connectedServer, out Exception _ex);
    }
}
