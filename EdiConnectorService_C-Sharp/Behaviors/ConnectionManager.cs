using System.Collections.Generic;

namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// This class is used to manage all connections that are added in the Connections dictionary.
    /// 
    /// </summary>
    class ConnectionManager
    {
        /// Collection of the connections where the string parameter is the key and its value is the connection object it self.
        /// The server name of a connection is used as a key.
        public Dictionary<string, SAPConnection> Connections { get; set; }

        // Create a static instance of the manager
        private static ConnectionManager instance = null;

        // Initialize Connections with constructor
        ConnectionManager()
        {
            Connections = new Dictionary<string, SAPConnection>();
        }

        /// <summary>
        /// Gets the instance.
        /// </summary>
        /// <returns></returns>
        public static ConnectionManager getInstance()
        {
            if (instance == null) // Create only 1 instance of the connection manager 
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

        /// <summary>
        /// Gets all connected servers.
        /// </summary>
        /// <returns>List of all connected servers</returns>
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

        /// <summary>
        /// Gets all disconnected servers.
        /// </summary>
        /// <returns>List of all disconnected servers</returns>
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

        /// <summary>
        /// Gets all server status.
        /// </summary>
        /// <returns>String of status of all connections</returns>
        public string GetAllServerStatus()
        {
            string connectedStatus = "";
            foreach (string c in GetAllConnectedServers())
                connectedStatus += (c + " -");

            string disconnectedStatus = "";
            foreach (string d in GetAllDisconnectedServers())
                disconnectedStatus += (d + " -");

            return $"Connected: {connectedStatus} Disconnected: {disconnectedStatus}";
        }

        /// <summary>
        /// Gets the connection.
        /// </summary>
        /// <param name="_serverName">Name of the server.</param>
        /// <returns>The connection object</returns>
        public SAPConnection GetConnection(string _serverName)
        {
            if (Connections.TryGetValue(_serverName, out SAPConnection value))
                return value;
            else
            {
                EventLogger.getInstance().EventError($"Can't find connection - Server: {_serverName}");
                return null;
            }
        }
    }
}
