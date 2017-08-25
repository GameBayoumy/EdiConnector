using System.Linq;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// Concrete Command class
    /// This command is used to load the configuration xml file and create connections which will be added and set in a connection list, held by the ConnectionManager.
    /// It will only create connections if every child element in the <Connection> has been filled in.
    /// </summary>
    /// <seealso cref="EdiConnectorService_C_Sharp.ICommand" />
    public class CreateConnectionsCommand : ICommand
    {
        public CreateConnectionsCommand()
        {

        }

        public void execute()
        {
            // Load configuration xml document
            XDocument xDoc = XDocument.Load(EdiConnectorData.getInstance().sApplicationPath + @"config.xml");

            // Iterate through every <Connection> element in <Connections>
            foreach (XElement xEle in xDoc.Element("Connections").Elements("Connection"))
            {
                bool emptyFieldFound = false;
                // Iterate through every child element in "Connection" element
                foreach(XElement xChild in xEle.Descendants())
                {
                    // Check if the child value is empty
                    if(xChild.Value == "")
                    {
                        emptyFieldFound = true;
                        break;
                    }
                }

                // If an empty child element has been found
                if (!emptyFieldFound)
                {
                    // Create a new connection with the "Server" element value as the key
                    ConnectionManager.getInstance().Connections.Add(xEle.Element("Server").Value, new SAPConnection());

                    // Use the last added connection object and set its value
                    ConnectionManager.getInstance().Connections.Last().Value.Set(xEle);
                }
                else
                    EventLogger.getInstance().EventError("Error creating connections: Empty field found in configuration " + xEle.Element("Server").Value);
            }
        }
    }
}
