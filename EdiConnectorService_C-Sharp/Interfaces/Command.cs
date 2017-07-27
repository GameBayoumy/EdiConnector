using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// Command interface
    /// </summary>
    public interface Command
    {
        void execute();
    }
}
