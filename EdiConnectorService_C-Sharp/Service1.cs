using System;
using System.ServiceProcess;
using System.Threading;
using System.IO;
using SAPbobsCOM;
using System.Net.Mail;
using System.Net;

namespace EdiConnectorService_C_Sharp
{
    public partial class EdiConnectorService : ServiceBase
    {
        private bool stopping;
        private ManualResetEvent stoppedEvent;

        public Agent agent;

        public EdiConnectorService()
        {
            InitializeComponent();
            this.stoppedEvent = new ManualResetEvent(false);
            this.stopping = false;
        }

        /// <summary>
        /// The function is executed when a Start command is sent to the service
        /// by the SCM or when the operating system starts (for a service that 
        /// starts automatically). It specifies actions to take when the service 
        /// starts. OnStart logs a service-start message to 
        /// the Application log, and queues the main service function for 
        /// execution in a thread pool worker thread.
        /// </summary>
        /// <param name="args">Command line arguments</param>
        /// <remarks>
        /// A service application is designed to be long running. Therefore, it 
        /// usually polls or monitors something in the system. The monitoring is 
        /// set up in the OnStart method. However, OnStart does not actually do 
        /// the monitoring. The OnStart method must return to the operating 
        /// system after the service's operation has begun. It must not loop 
        /// forever or block. To set up a simple monitoring mechanism, one 
        /// general solution is to create a timer in OnStart. The timer would 
        /// then raise events in your code periodically, at which time your 
        /// service could do its monitoring. The other solution is to spawn a 
        /// new thread to perform the main service functions.
        /// </remarks>
        protected override void OnStart(string[] args)
        {
            EventLogger.getInstance().EventInfo("EdiService in OnStart.");

            // Initialize objects
            EdiConnectorData.GetInstance();
            ConnectionManager.getInstance();
            agent = new Agent();
            EdiConnectorData.GetInstance().ApplicationPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + @"\";
            EdiConnectorData.GetInstance().ProcessedDirName = "Processed";

            // Create SAP connections
            agent.QueueCommand(new CreateConnectionsCommand());
            // Try to connect to all servers that are created
            ConnectionManager.getInstance().ConnectAll();

            //Creates udf fields for every connected server

            agent.QueueCommand(new CreateUserDefinitionsCommand());

            ThreadPool.QueueUserWorkItem(new WaitCallback(ServiceWorkerThread));
        }

        /// <summary>
        /// The method performs the main function of the service. It runs on a 
        /// thread pool worker thread.
        /// </summary>
        /// <param name="state"></param>
        private void ServiceWorkerThread(Object state)
        {
            // Periodically check if the service is stopping.
            do
            {
                // Perform main service function here...

                // Processes incoming messages
                foreach (string connectedServer in ConnectionManager.getInstance().GetAllConnectedServers())
                {
                    string messagesFilePath = ConnectionManager.getInstance().GetConnection(connectedServer).MessagesFilePath;
                    if (Directory.Exists(messagesFilePath))
                    {
                        DirectoryInfo dirInfo = new DirectoryInfo(messagesFilePath);

                        foreach (var file in dirInfo.EnumerateFiles("*.xml", SearchOption.TopDirectoryOnly))
                        {
                            agent.QueueCommand(new ProcessMessage(connectedServer, file.Name));
                        }
                    }
                    else
                    {
                        EventLogger.getInstance().EventError($"Server: {connectedServer}. Messages file path does not exist!: {messagesFilePath}");
                    }
                }

                // After performing functions, put the thread in sleep
                Thread.Sleep(10000);
            }
            while (!stopping);

            this.stoppedEvent.Set();
        }

        /// <summary>
        /// The function is executed when a Stop command is sent to the service 
        /// by SCM. It specifies actions to take when a service stops running.
        /// OnStop logs a service-stop message to the Application log, 
        /// and waits for the finish of the main service function.
        /// </summary>
        protected override void OnStop()
        {
            //Add code here to perform any tear-down necessary to stop your service.
            ConnectionManager.getInstance().DisconnectAll();

            //Log a service stop message to the Application log.
            this.eventLog1.WriteEntry("EdiService in OnStop.");

            //Indicate that the service is stopping and wait for the finish of 
            //the main service function (ServiceWorkerThread).
            this.stopping = true;
            this.stoppedEvent.WaitOne();
        }

        public static void ClearObject(object t)
        {
            if (t != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(t);
                t = null;
                GC.Collect();
            }
        }
    }
}
