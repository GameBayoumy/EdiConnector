using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EdiConnectorService_C_Sharp
{
    class ConnectionManager : AConnection
    {
        public Dictionary<string, SAPConnection> Connections { get; set; }

        public ConnectionManager()
        {
            Connections = new Dictionary<string, SAPConnection>();
        }

        /// <summary>
        /// Connects all connections in the list.
        /// </summary>
        public void ConnectAll()
        {
            foreach (SAPConnection connection in Connections.Values)
            {
                connection.Connect();
            }
        }

        /// <summary>
        /// Disconnects all connections in the list.
        /// </summary>
        public void DisconnectAll()
        {
            foreach (SAPConnection connection in Connections.Values)
            {
                connection.Disconnect();
            }
        }

        public string GetAllConnectedServers()
        {
            string connected = "";
            foreach (SAPConnection connection in Connections.Values)
            {
                if (connection.ConnectedToSAP)
                    connected += connection.Company.Server;
            }
            return connected;
        }

        public string GetAllDisconnectedServers()
        {
            string disconnected = "";
            foreach (SAPConnection connection in Connections.Values)
            {
                if (!connection.ConnectedToSAP)
                    disconnected += connection.Company.Server;
            }
            return disconnected;
        }

        public string GetAllServerStatus()
        {
            return "Connected: " + GetAllConnectedServers() + " Disconnected: " + GetAllDisconnectedServers();
        }
    }
}
