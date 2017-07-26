using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
        public abstract void Set();
        /// <summary>
        /// Connects this instance.
        /// </summary>
        /// <returns></returns>
        public abstract bool Connect();
        /// <summary>
        /// Disconnects this instance.
        /// </summary>
        /// <returns></returns>
        public abstract void Disconnect();
    }
}
