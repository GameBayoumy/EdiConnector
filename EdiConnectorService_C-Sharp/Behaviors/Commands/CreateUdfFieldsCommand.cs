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
    /// <seealso cref="EdiConnectorService_C_Sharp.Command" />
    public class CreateUfdFieldsCommand : Command
    {
        public CreateUfdFieldsCommand()
        {

        }

        public void execute()
        {
            UdfFields.CreateUdfFields();
        }
    }
}
