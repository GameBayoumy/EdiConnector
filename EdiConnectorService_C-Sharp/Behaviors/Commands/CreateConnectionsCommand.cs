using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

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
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(EdiConnectorData.getInstance().sApplicationPath + @"config.xml");
            XmlNodeList xmlList = xmlDoc.SelectNodes("/Connections/Connection");
            for (int i = 0; i < xmlList.Count; i++)
            {
                ConnectionManager.getInstance().Connections.Add(xmlList[i]["Server"].InnerText, new SAPConnection());
                ConnectionManager.getInstance().Connections.Last().Value.Set(xmlList[i]);
            }
        }
    }
}
