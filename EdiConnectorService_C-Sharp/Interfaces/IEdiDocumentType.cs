﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    interface IEdiDocumentType
    {
        Object ReadXMLData(XElement _xMessages);
        void SaveToSAP(Object _dataObject, string _connectedServer);
    }
}
