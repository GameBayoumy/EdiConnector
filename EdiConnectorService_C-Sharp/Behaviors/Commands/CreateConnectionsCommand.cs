using System.Linq;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// Concrete Command class
    /// </summary>
    /// <seealso cref="EdiConnectorService_C_Sharp.ICommand" />
    public class CreateConnectionsCommand : ICommand
    {
        public CreateConnectionsCommand()
        {

        }

        public void execute()
        {
            // Create and set connections from config.xml
            XDocument xDoc = XDocument.Load(EdiConnectorData.getInstance().sApplicationPath + @"config.xml");
            foreach (XElement xEle in xDoc.Element("Connections").Elements("Connection"))
            {
                bool emptyFieldFound = false;
                foreach(XElement xChild in xEle.Descendants())
                {
                    if(xChild.Value == "")
                    {
                        emptyFieldFound = true;
                        break;
                    }
                }
                if (!emptyFieldFound)
                {
                    ConnectionManager.getInstance().Connections.Add(xEle.Element("Server").Value, new SAPConnection());
                    ConnectionManager.getInstance().Connections.Last().Value.Set(xEle);
                }
                else
                    EventLogger.getInstance().EventError("Error creating connections: Empty field found in configuration " + xEle.Element("Server"));
            }
        }
    }
}
