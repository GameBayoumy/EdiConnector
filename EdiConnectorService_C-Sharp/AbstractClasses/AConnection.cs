using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// The 'Connection' abstract class
    /// </summary>
    abstract class AConnection
    {
        /// <summary>
        /// Sets this instance.
        /// </summary>
        public abstract void Set(XmlNode _xmlNode);
        /// <summary>
        /// Connects this instance.
        /// </summary>
        /// <returns></returns>
        public abstract bool Connect();
        /// <summary>
        /// Disconnects this instance.
        /// </summary>
        /// <returns></returns>
        public abstract bool Disconnect();
    }
}
