using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EdiConnectorService_C_Sharp
{
    class ConnectionManager
    {
        public Dictionary<string, SAPConnection> Connections { get; set; }
        private static ConnectionManager instance = null;

        ConnectionManager()
        {
            Connections = new Dictionary<string, SAPConnection>();
        }

        public static ConnectionManager getInstance()
        {
            if (instance == null)
            {
                instance = new ConnectionManager();
            }
            return instance;
        }

        /// <summary>
        /// Connects all disconnected servers in the list.
        /// </summary>
        public void ConnectAll()
        {
            foreach (string disconnectedServer in GetAllDisconnectedServers())
            {
                GetConnection(disconnectedServer).Connect();
            }
        }

        /// <summary>
        /// Disconnects all connected servers in the list.
        /// </summary>
        public void DisconnectAll()
        {
            foreach (string connectedServer in GetAllConnectedServers())
            {
                GetConnection(connectedServer).Disconnect();
            }
        }

        public List<string> GetAllConnectedServers()
        {
            List<string> connected = new List<string>();
            foreach (SAPConnection connection in Connections.Values)
            {
                if (connection.ConnectedToSAP)
                    connected.Add(connection.Company.Server);
            }
            return connected;
        }

        public List<string> GetAllDisconnectedServers()
        {
            List<string> disconnected = new List<string>();
            foreach (SAPConnection connection in Connections.Values)
            {
                if (!connection.ConnectedToSAP)
                    disconnected.Add(connection.Company.Server);
            }
            return disconnected;
        }

        public string GetAllServerStatus()
        {
            string connectedStatus = "";
            foreach (string c in GetAllConnectedServers())
                connectedStatus += (c + " -");

            string disconnectedStatus = "";
            foreach (string d in GetAllDisconnectedServers())
                disconnectedStatus += (d + " -");

            return "Connected: " + connectedStatus + " Disconnected: " + disconnectedStatus;
        }

        public SAPConnection GetConnection(string _serverName)
        {
            SAPConnection value;
            if (Connections.TryGetValue(_serverName, out value))
                return value;
            else
            {
                EventLogger.getInstance().EventError("Can't find connection - Server: " + _serverName);
                return null;
            }
        }
    }
}
