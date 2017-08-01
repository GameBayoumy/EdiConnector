﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// Concrete Command class
    /// </summary>
    /// <seealso cref="EdiConnectorService_C_Sharp.Command" />
    public class CreateUserDefinitionsCommand : Command
    {
        public CreateUserDefinitionsCommand()
        {

        }

        public void execute()
        {
            // For each connected server it will try to create udf fields for that server
            foreach (string connectedServer in ConnectionManager.getInstance().GetAllConnectedServers())
            {
                UserDefined.CreateTables(connectedServer);
                UserDefined.CreateFields(connectedServer);
            }
        }
    }
}
