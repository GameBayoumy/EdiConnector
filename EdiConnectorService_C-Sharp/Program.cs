using System.ServiceProcess;

namespace EdiConnectorService_C_Sharp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
			{ 
				new EdiConnectorService() 
			};
            ServiceBase.Run(ServicesToRun);
        }

    }
}
