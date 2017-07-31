using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.IO;
using System.Data.SqlClient;
using SAPbobsCOM;
using System.Net.Mail;
using System.Net;
using System.Xml.Linq;

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

            // Initialize objects
            EdiConnectorData.getInstance();
            ConnectionManager.getInstance();
            agent = new Agent();

            EdiConnectorData.getInstance().sApplicationPath = @"H:\Projecten\Sharif\GitKraken\EdiConnector\EdiConnectorService_C-Sharp";

            agent.QueueCommand(new ProcessMessage(EdiConnectorData.getInstance().sApplicationPath, @"\orders.xml"));

        }

        /* <summary>
    ''' The function is executed when a Start command is sent to the service
    ''' by the SCM or when the operating system starts (for a service that 
    ''' starts automatically). It specifies actions to take when the service 
    ''' starts. In this code sample, OnStart logs a service-start message to 
    ''' the Application log, and queues the main service function for 
    ''' execution in a thread pool worker thread.
    ''' </summary>
    ''' <param name="args">Command line arguments</param>
    ''' <remarks>
    ''' A service application is designed to be long running. Therefore, it 
    ''' usually polls or monitors something in the system. The monitoring is 
    ''' set up in the OnStart method. However, OnStart does not actually do 
    ''' the monitoring. The OnStart method must return to the operating 
    ''' system after the service's operation has begun. It must not loop 
    ''' forever or block. To set up a simple monitoring mechanism, one 
    ''' general solution is to create a timer in OnStart. The timer would 
    ''' then raise events in your code periodically, at which time your 
    ''' service could do its monitoring. The other solution is to spawn a 
    ''' new thread to perform the main service functions, which is 
    ''' demonstrated in this code sample.
    ''' </remarks>
         * */
        protected override void OnStart(string[] args)
        {
            EventLogger.getInstance().EventInfo("EdiService in OnStart.");

            EdiConnectorData.getInstance().sApplicationPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);

            // Create connections from config.xml and try to connect all servers
            agent.QueueCommand(new CreateConnectionsCommand());
            ConnectionManager.getInstance().ConnectAll();

            // Creates udf fields for every connected server
            agent.QueueCommand(new CreateUfdFieldsCommand());

            //ReadSettings();
            //ConnectToSAP();
            //CreateUdfFieldsText();

            ThreadPool.QueueUserWorkItem(new WaitCallback(ServiceWorkerThread));
        }

        /*''' <summary>
    ''' The method performs the main function of the service. It runs on a 
    ''' thread pool worker thread.
    ''' </summary>
    ''' <param name="state"></param>
         * */
        private void ServiceWorkerThread(Object state)
        {
            //Periodically check if the service is stopping.
            do
            {
                //Perform main service function here...

                //If there are any servers connected to SAP
                if(ConnectionManager.getInstance().GetAllConnectedServers().Count > 0)
                {

                }


                if (EdiConnectorData.getInstance().cmp.Connected == true && EdiConnectorData.getInstance().cn.State == ConnectionState.Open)
                {
                    CheckAndExportDelivery();
                    CheckAndExportInvoice();
                    SplitOrder();
                    ReadSOFile();
                }
                else
                {
                    ConnectToSAP();
                }

                Thread.Sleep(6000);

            }
            while (!stopping);

            this.stoppedEvent.Set();
        }

        /*''' <summary>
    ''' The function is executed when a Stop command is sent to the service 
    ''' by SCM. It specifies actions to take when a service stops running. In 
    ''' this code sample, OnStop logs a service-stop message to the 
    ''' Application log, and waits for the finish of the main service 
    ''' function.
    ''' </summary>
             * */
        protected override void OnStop()
        {
            //Add code here to perform any tear-down necessary to stop your service.
            DisconnectToSAP();

            //Log a service stop message to the Application log.
            this.eventLog1.WriteEntry("EdiService in OnStop.");

            //Indicate that the service is stopping and wait for the finish of 
            //the main service function (ServiceWorkerThread).
            this.stopping = true;
            this.stoppedEvent.WaitOne();
        }

        #region
        public bool ConnectToSAP()
        {
            try
            {
                if (EdiConnectorData.getInstance().cmp.Connected == true)
                {
                    Log("V", "SAP is already Connected", "ConnectToSAP");
                    return true;
                }

                bool bln = ConnectToDatabase();

                if (bln == false)
                {
                    return false;
                }

                EdiConnectorData.getInstance().cmp.DbServerType = EdiConnectorData.getInstance().bstDBServerType;
                EdiConnectorData.getInstance().cmp.Server = EdiConnectorData.getInstance().sServer;
                EdiConnectorData.getInstance().cmp.DbUserName = EdiConnectorData.getInstance().sDBUserName;
                EdiConnectorData.getInstance().cmp.DbPassword = EdiConnectorData.getInstance().sDBPassword;
                EdiConnectorData.getInstance().cmp.CompanyDB = EdiConnectorData.getInstance().sCompanyDB;
                EdiConnectorData.getInstance().cmp.UserName = EdiConnectorData.getInstance().sUserName;
                EdiConnectorData.getInstance().cmp.Password = EdiConnectorData.getInstance().sPassword;
                EdiConnectorData.getInstance().cmp.UseTrusted = false;

                if (EdiConnectorData.getInstance().cmp.Connect() != 0)
                {
                    int errCode;
                    string errMsg = "";
                    EdiConnectorData.getInstance().cmp.GetLastError(out errCode, out errMsg);
                    if (errCode != 0)
                    {
                        Log("X", "Error: " + errMsg + "(" + errCode.ToString() + ")", "ConnectToSAP");
                        return false;
                    }
                }
                else
                {
                    Log("V", "Connected to SAP", "ConnectToSAP");
                    return true;
                }
            }
            catch (Exception e)
            {
                Log("X", e.Message, "ConnectToSAP");
                return false;
            }

            return false;
        }

        private bool ConnectToDatabase()
        {
            try
            {
                EdiConnectorData.getInstance().cn.ConnectionString = "Data Source=" + EdiConnectorData.getInstance().sServer +
                    ";Initial Catalog=" + EdiConnectorData.getInstance().sCompanyDB +
                    ";User ID=" + EdiConnectorData.getInstance().sDBUserName +
                    ";Password=" + EdiConnectorData.getInstance().sDBPassword;

                if (EdiConnectorData.getInstance().sSQLVersion == "2005")
                    EdiConnectorData.getInstance().bstDBServerType = BoDataServerTypes.dst_MSSQL2005;
                else if (EdiConnectorData.getInstance().sSQLVersion == "2008")
                    EdiConnectorData.getInstance().bstDBServerType = BoDataServerTypes.dst_MSSQL2008;
                else if (EdiConnectorData.getInstance().sSQLVersion == "2012")
                    EdiConnectorData.getInstance().bstDBServerType = BoDataServerTypes.dst_MSSQL2012;
                else
                    EdiConnectorData.getInstance().bstDBServerType = BoDataServerTypes.dst_MSSQL;

                EdiConnectorData.getInstance().cn.Open();

                Log("V", "Connected to database " + EdiConnectorData.getInstance().sCompanyDB + ".", "ConnectToDatabase");
                return true;
            }
            catch (Exception e)
            {
                Log("X", e.Message, "ConnectToDatabase");
                return false;
            }

        }

        public void ReadSettings()
        {
            try
            {
                DataSet dataSet = new DataSet();
                dataSet.ReadXml(EdiConnectorData.getInstance().sApplicationPath + @"\settings.xml");

                if (dataSet.Tables["server"].Rows[0]["sqlversion"].ToString() == "2005")
                    EdiConnectorData.getInstance().bstDBServerType = BoDataServerTypes.dst_MSSQL2005;
                else if (dataSet.Tables["server"].Rows[0]["sqlversion"].ToString() == "2008")
                    EdiConnectorData.getInstance().bstDBServerType = BoDataServerTypes.dst_MSSQL2008;
                else if (dataSet.Tables["server"].Rows[0]["sqlversion"].ToString() == "2012")
                    EdiConnectorData.getInstance().bstDBServerType = BoDataServerTypes.dst_MSSQL2012;
                else
                    EdiConnectorData.getInstance().bstDBServerType = BoDataServerTypes.dst_MSSQL;

                EdiConnectorData.getInstance().sServer = dataSet.Tables["server"].Rows[0]["name"].ToString();
                EdiConnectorData.getInstance().sDBUserName = dataSet.Tables["server"].Rows[0]["userid"].ToString();
                EdiConnectorData.getInstance().sDBPassword = dataSet.Tables["server"].Rows[0]["password"].ToString();
                EdiConnectorData.getInstance().sCompanyDB = dataSet.Tables["server"].Rows[0]["catalog"].ToString();
                EdiConnectorData.getInstance().sUserName = dataSet.Tables["server"].Rows[0]["sapuser"].ToString();
                EdiConnectorData.getInstance().sPassword = dataSet.Tables["server"].Rows[0]["sappassword"].ToString();
                EdiConnectorData.getInstance().sSQLVersion = dataSet.Tables["server"].Rows[0]["sqlversion"].ToString();
                EdiConnectorData.getInstance().sDesAdvLevel = dataSet.Tables["server"].Rows[0]["desadvlevel"].ToString();

                EdiConnectorData.getInstance().sSOPath = dataSet.Tables["folders"].Rows[0]["so"].ToString();
                EdiConnectorData.getInstance().sSOTempPath = dataSet.Tables["folders"].Rows[0]["sotemp"].ToString();
                EdiConnectorData.getInstance().sSODonePath = dataSet.Tables["folders"].Rows[0]["sodone"].ToString();
                EdiConnectorData.getInstance().sSOErrorPath = dataSet.Tables["folders"].Rows[0]["soerror"].ToString();
                EdiConnectorData.getInstance().sInvoicePath = dataSet.Tables["folders"].Rows[0]["invoice"].ToString();
                EdiConnectorData.getInstance().sDeliveryPath = dataSet.Tables["folders"].Rows[0]["delivery"].ToString();

                EdiConnectorData.getInstance().iSendNotification = Convert.ToInt32(dataSet.Tables["email"].Rows[0]["send_notification"]);

                EdiConnectorData.getInstance().sSmpt = dataSet.Tables["email"].Rows[0]["smtp"].ToString();
                EdiConnectorData.getInstance().iSmtpPort = Convert.ToInt32(dataSet.Tables["email"].Rows[0]["port"]);

                if (dataSet.Tables["email"].Rows[0]["security"].ToString() == "0")
                    EdiConnectorData.getInstance().bSmtpUserSecurity = false;
                else
                    EdiConnectorData.getInstance().bSmtpUserSecurity = true;

                EdiConnectorData.getInstance().sSmtpUser = dataSet.Tables["email"].Rows[0]["user"].ToString();
                EdiConnectorData.getInstance().sSmtpPassword = dataSet.Tables["email"].Rows[0]["password"].ToString();

                EdiConnectorData.getInstance().sSenderEmail = dataSet.Tables["email"].Rows[0]["emailaddress"].ToString();
                EdiConnectorData.getInstance().sSenderName = dataSet.Tables["email"].Rows[0]["fullname"].ToString();

                EdiConnectorData.getInstance().sOrderMailTo = dataSet.Tables["email"].Rows[0]["emailaddress_order"].ToString();
                EdiConnectorData.getInstance().sOrderMailToFullName = dataSet.Tables["email"].Rows[0]["fullname_order"].ToString();

                EdiConnectorData.getInstance().sDeliveryMailTo = dataSet.Tables["email"].Rows[0]["emailaddress_delivery"].ToString();
                EdiConnectorData.getInstance().sDeliveryMailToFullName = dataSet.Tables["email"].Rows[0]["fullname_delivery"].ToString();

                EdiConnectorData.getInstance().sInvoiceMailTo = dataSet.Tables["email"].Rows[0]["emailaddress_invoice"].ToString();
                EdiConnectorData.getInstance().sInvoiceMailToFullName = dataSet.Tables["email"].Rows[0]["fullname_invoice"].ToString();

                dataSet.Dispose();

                ReadInterface();
            }
            catch (Exception e)
            {
                Log("X", "Settings are not in the right format! - " + e.Message, "ReadSettings");
            }
        }

        public void ReadInterface()
        {
            DataSet dataSet = new DataSet();
            dataSet.ReadXml(EdiConnectorData.getInstance().sApplicationPath + @"\orderfld.xml");

            for (int orderHead = 0; orderHead < dataSet.Tables["OK"].Rows.Count - 1; orderHead++ )
            {
                EdiConnectorData.getInstance().OK_POS[orderHead] = Convert.ToInt32(dataSet.Tables["OK"].Rows[orderHead][0]);
                EdiConnectorData.getInstance().OK_LEN[orderHead] = Convert.ToInt32(dataSet.Tables["OK"].Rows[orderHead][1]);
            }

            for (int orderItem = 0; orderItem < dataSet.Tables["OR"].Rows.Count - 1; orderItem++)
            {
                EdiConnectorData.getInstance().OK_POS[orderItem] = Convert.ToInt32(dataSet.Tables["OR"].Rows[orderItem][0]);
                EdiConnectorData.getInstance().OK_LEN[orderItem] = Convert.ToInt32(dataSet.Tables["OR"].Rows[orderItem][1]);
            }

            dataSet.Dispose();
        }

        public void Log(string sType, string msg, string functionSender)
        {
            switch(sType)
            {
                case "V":
                    this.eventLog1.WriteEntry(sType + " - " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " - " + functionSender.Replace("_", " ") + " - " + msg, System.Diagnostics.EventLogEntryType.Information);
                    break;
                case "X":
                    this.eventLog1.WriteEntry(sType + " - " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " - " + functionSender.Replace("_", " ") + " - " + msg, System.Diagnostics.EventLogEntryType.Error);
                    break;
                default:
                    this.eventLog1.WriteEntry(sType + " - " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " - " + functionSender.Replace("_", " ") + " - " + msg);
                    break;
            }
        }

        public void DisconnectToSAP()
        {
            if (EdiConnectorData.getInstance().cmp.Connected == true)
            {
                EdiConnectorData.getInstance().cmp.Disconnect();
                EdiConnectorData.getInstance().cn.Close();
                Log("V", "SAP is disconneted.", "DisconnectToSAP");
            }
            else
                Log("X", "SAP is already disconnected.", "DisconnectToSAP");
        }

        public void SplitOrder()
        {
            string fileSize = "";
            DirectoryInfo dirI = new DirectoryInfo(EdiConnectorData.getInstance().sSOPath);
            FileInfo[] arrayFileI = dirI.GetFiles("*.DAT");
            string newFile = "";
            string BGMnumber = "";

            foreach(FileInfo fileInfo in arrayFileI)
            {
                string fileLength = (fileInfo.Length / 1024).ToString();
                fileSize = String.Format(fileLength, "##0.00");
                Log("V", "Begin splitting sales order filename " + fileInfo.Name + " - file size: " + fileSize + " kb", "Split_Order");

                string line;

                using (StreamReader reader = new StreamReader(fileInfo.FullName))
                {
                    while (true)
                    {
                        line = reader.ReadLine();
                        if (String.IsNullOrWhiteSpace(line))
                            break;

                        if (line.Substring(1, 1) == "0")
                        {
                            if (newFile.Length > 0)
                            {
                                WriteNewFile(newFile, BGMnumber);
                                newFile = "";
                            }

                            BGMnumber = line.Trim(line.Substring(30, 35).ToCharArray());
                        }

                        newFile += line + System.Environment.NewLine;
                    }

                    if (newFile.Length > 0)
                        WriteNewFile(newFile, BGMnumber);
                }

                try
                {
                    File.Delete(fileInfo.FullName);
                    Log("V", "Splitted Sales order filename " + fileInfo.Name + " deleted.", "SplitOrder");
                }
                catch (Exception e)
                {
                    Log("X", e.Message, "SplitOrder");
                }
            }
        }

        public void WriteNewFile(string newFile, string BGMnumber)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(EdiConnectorData.getInstance().sSOTempPath + @"\" + "ORDER" + "_" + BGMnumber + ".DAT"))
                {
                    writer.Write(newFile);
                    writer.Close();
                    Log("V", "New splitted order created " + BGMnumber, "WriteNewFile");
                }
            }
            catch (Exception e)
            {
                Log("X", e.Message, "WriteNewFile");
            }
        }

        public void ReadSOFile()
        {
            DirectoryInfo dirI = new DirectoryInfo(EdiConnectorData.getInstance().sSOTempPath);
            FileInfo[] arrayFileI = dirI.GetFiles("*.DAT");

            foreach (FileInfo fileInfo in arrayFileI)
            {
                try
                {
                    Log("V", "Begin reading sales order filename " + fileInfo.Name + ".", "Read_SO_file");

                    StreamReader reader = new StreamReader(fileInfo.FullName, false);
                    EdiConnectorData.getInstance().SO_FILE = reader.ReadToEnd();
                    EdiConnectorData.getInstance().SO_FILENAME = fileInfo.Name;
                    reader.Close();
                }
                catch (Exception e)
                {
                    Log("X", e.Message, "ReadSOFile");
                }

                MatchSOdata();
            }
        }

        public bool MatchSOdata()
        {
            try
            {
                string line = "";

                using (StringReader reader = new StringReader(EdiConnectorData.getInstance().SO_FILE))
                {
                    line = reader.ReadLine();
                    switch (line.Substring(1, 1))
                    {
                        case "0":
                            EdiConnectorData.getInstance().OK_K_EANCODE = line.Substring(EdiConnectorData.getInstance().OK_POS[0], EdiConnectorData.getInstance().OK_LEN[0]).Trim();
                            EdiConnectorData.getInstance().OK_TEST = line.Substring(EdiConnectorData.getInstance().OK_POS[1], EdiConnectorData.getInstance().OK_LEN[1]).Trim();
                            EdiConnectorData.getInstance().OK_KNAAM = line.Substring(EdiConnectorData.getInstance().OK_POS[2], EdiConnectorData.getInstance().OK_LEN[2]).Trim();
                            EdiConnectorData.getInstance().OK_BGM = line.Substring(EdiConnectorData.getInstance().OK_POS[3], EdiConnectorData.getInstance().OK_LEN[3]).Trim();
                            if (line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Trim().Length == 8)
                                EdiConnectorData.getInstance().OK_K_ORDDAT = Convert.ToDateTime(line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(7, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(1, 4));
                            else
                                EdiConnectorData.getInstance().OK_K_ORDDAT = Convert.ToDateTime("01-01-0001");

                            if (line.Substring(EdiConnectorData.getInstance().OK_POS[5], EdiConnectorData.getInstance().OK_LEN[5]).Trim().Length == 8)
                                EdiConnectorData.getInstance().OK_DTM_2 = Convert.ToDateTime(line.Substring(EdiConnectorData.getInstance().OK_POS[5], EdiConnectorData.getInstance().OK_LEN[5]).Substring(7, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(1, 4));
                            else
                                EdiConnectorData.getInstance().OK_DTM_2 = Convert.ToDateTime("01-01-0001");
                            EdiConnectorData.getInstance().OK_TIJD_2 = line.Substring(EdiConnectorData.getInstance().OK_POS[6], EdiConnectorData.getInstance().OK_LEN[6]).Trim();

                            if (line.Substring(EdiConnectorData.getInstance().OK_POS[7], EdiConnectorData.getInstance().OK_LEN[7]).Trim().Length == 8)
                                EdiConnectorData.getInstance().OK_DTM_17 = Convert.ToDateTime(line.Substring(EdiConnectorData.getInstance().OK_POS[7], EdiConnectorData.getInstance().OK_LEN[7]).Substring(7, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(1, 4));
                            else
                                EdiConnectorData.getInstance().OK_DTM_17 = Convert.ToDateTime("01-01-0001");
                            EdiConnectorData.getInstance().OK_TIJD_17 = line.Substring(EdiConnectorData.getInstance().OK_POS[8], EdiConnectorData.getInstance().OK_LEN[8]).Trim();

                            if (line.Substring(EdiConnectorData.getInstance().OK_POS[9], EdiConnectorData.getInstance().OK_LEN[9]).Trim().Length == 8)
                                EdiConnectorData.getInstance().OK_DTM_64 = Convert.ToDateTime(line.Substring(EdiConnectorData.getInstance().OK_POS[9], EdiConnectorData.getInstance().OK_LEN[9]).Substring(7, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(1, 4));
                            else
                                EdiConnectorData.getInstance().OK_DTM_64 = Convert.ToDateTime("01-01-0001");
                            EdiConnectorData.getInstance().OK_TIJD_64 = line.Substring(EdiConnectorData.getInstance().OK_POS[10], EdiConnectorData.getInstance().OK_LEN[10]).Trim();

                            if (line.Substring(EdiConnectorData.getInstance().OK_POS[11], EdiConnectorData.getInstance().OK_LEN[11]).Trim().Length == 8)
                                EdiConnectorData.getInstance().OK_DTM_63 = Convert.ToDateTime(line.Substring(EdiConnectorData.getInstance().OK_POS[11], EdiConnectorData.getInstance().OK_LEN[11]).Substring(7, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(1, 4));
                            else
                                EdiConnectorData.getInstance().OK_DTM_63 = Convert.ToDateTime("01-01-0001");
                            EdiConnectorData.getInstance().OK_TIJD_63 = line.Substring(EdiConnectorData.getInstance().OK_POS[12], EdiConnectorData.getInstance().OK_LEN[12]).Trim();

                            EdiConnectorData.getInstance().OK_RFF_BO = line.Substring(EdiConnectorData.getInstance().OK_POS[13], EdiConnectorData.getInstance().OK_LEN[13]).Trim();
                            EdiConnectorData.getInstance().OK_RFF_CR = line.Substring(EdiConnectorData.getInstance().OK_POS[14], EdiConnectorData.getInstance().OK_LEN[14]).Trim();
                            EdiConnectorData.getInstance().OK_RFF_PD = line.Substring(EdiConnectorData.getInstance().OK_POS[15], EdiConnectorData.getInstance().OK_LEN[15]).Trim();
                            EdiConnectorData.getInstance().OK_RFFCT = line.Substring(EdiConnectorData.getInstance().OK_POS[16], EdiConnectorData.getInstance().OK_LEN[16]).Trim();
                            if (line.Substring(EdiConnectorData.getInstance().OK_POS[17], EdiConnectorData.getInstance().OK_LEN[17]).Trim().Length == 8)
                                EdiConnectorData.getInstance().OK_DTMCT = Convert.ToDateTime(line.Substring(EdiConnectorData.getInstance().OK_POS[17], EdiConnectorData.getInstance().OK_LEN[17]).Substring(7, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OK_POS[4], EdiConnectorData.getInstance().OK_LEN[4]).Substring(1, 4));
                            else
                                EdiConnectorData.getInstance().OK_DTMCT = Convert.ToDateTime("01-01-0001");

                            EdiConnectorData.getInstance().OK_FLAGS[0] = line.Substring(EdiConnectorData.getInstance().OK_POS[18], EdiConnectorData.getInstance().OK_LEN[18]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[1] = line.Substring(EdiConnectorData.getInstance().OK_POS[19], EdiConnectorData.getInstance().OK_LEN[19]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[2] = line.Substring(EdiConnectorData.getInstance().OK_POS[20], EdiConnectorData.getInstance().OK_LEN[20]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[3] = line.Substring(EdiConnectorData.getInstance().OK_POS[21], EdiConnectorData.getInstance().OK_LEN[21]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[4] = line.Substring(EdiConnectorData.getInstance().OK_POS[22], EdiConnectorData.getInstance().OK_LEN[22]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[5] = line.Substring(EdiConnectorData.getInstance().OK_POS[23], EdiConnectorData.getInstance().OK_LEN[23]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[6] = line.Substring(EdiConnectorData.getInstance().OK_POS[24], EdiConnectorData.getInstance().OK_LEN[24]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[7] = line.Substring(EdiConnectorData.getInstance().OK_POS[25], EdiConnectorData.getInstance().OK_LEN[25]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[8] = line.Substring(EdiConnectorData.getInstance().OK_POS[26], EdiConnectorData.getInstance().OK_LEN[26]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[9] = line.Substring(EdiConnectorData.getInstance().OK_POS[27], EdiConnectorData.getInstance().OK_LEN[27]).Trim();
                            EdiConnectorData.getInstance().OK_FLAGS[10] = line.Substring(EdiConnectorData.getInstance().OK_POS[28], EdiConnectorData.getInstance().OK_LEN[28]).Trim();
                            EdiConnectorData.getInstance().OK_FTXDSI = line.Substring(EdiConnectorData.getInstance().OK_POS[29], EdiConnectorData.getInstance().OK_LEN[29]).Trim();
                            EdiConnectorData.getInstance().OK_NAD_BY = line.Substring(EdiConnectorData.getInstance().OK_POS[30], EdiConnectorData.getInstance().OK_LEN[30]).Trim();
                            EdiConnectorData.getInstance().OK_NAD_DP = line.Substring(EdiConnectorData.getInstance().OK_POS[31], EdiConnectorData.getInstance().OK_LEN[31]).Trim();
                            EdiConnectorData.getInstance().OK_NAD_IV = line.Substring(EdiConnectorData.getInstance().OK_POS[32], EdiConnectorData.getInstance().OK_LEN[32]).Trim();
                            EdiConnectorData.getInstance().OK_NAD_SF = line.Substring(EdiConnectorData.getInstance().OK_POS[33], EdiConnectorData.getInstance().OK_LEN[33]).Trim();
                            EdiConnectorData.getInstance().OK_NAD_SU = line.Substring(EdiConnectorData.getInstance().OK_POS[34], EdiConnectorData.getInstance().OK_LEN[34]).Trim();
                            EdiConnectorData.getInstance().OK_NAD_UC = line.Substring(EdiConnectorData.getInstance().OK_POS[35], EdiConnectorData.getInstance().OK_LEN[35]).Trim();
                            EdiConnectorData.getInstance().OK_NAD_BCO = line.Substring(EdiConnectorData.getInstance().OK_POS[36], EdiConnectorData.getInstance().OK_LEN[36]).Trim();
                            EdiConnectorData.getInstance().OK_RECEIVER = line.Substring(EdiConnectorData.getInstance().OK_POS[37], EdiConnectorData.getInstance().OK_LEN[37]).Trim();
                            bool matchHead = WriteSOhead();
                            if (matchHead == false)
                            {
                                // Move file
                                File.Move(EdiConnectorData.getInstance().sSOTempPath + @"\" + EdiConnectorData.getInstance().SO_FILENAME, EdiConnectorData.getInstance().sSOErrorPath + @"\" + DateTime.Now.ToString("HHmmss") + "_" + EdiConnectorData.getInstance().SO_FILENAME);
                                Log("X", EdiConnectorData.getInstance().SO_FILENAME + " copied to the errors folder!", "MatchSOdata");
                                return false;
                            }
                            break;

                        case "1":
                            EdiConnectorData.getInstance().OR_DEUAC = line.Substring(EdiConnectorData.getInstance().OR_POS[0], EdiConnectorData.getInstance().OR_LEN[0]).Trim();
                            EdiConnectorData.getInstance().OR_QTY = Convert.ToDouble(line.Substring(EdiConnectorData.getInstance().OR_POS[1], EdiConnectorData.getInstance().OR_LEN[1]).Trim());
                            EdiConnectorData.getInstance().OR_LEVARTCODE = line.Substring(EdiConnectorData.getInstance().OR_POS[2], EdiConnectorData.getInstance().OR_LEN[2]).Trim();
                            EdiConnectorData.getInstance().OR_DEARTOM = line.Substring(EdiConnectorData.getInstance().OR_POS[3], EdiConnectorData.getInstance().OR_LEN[3]).Trim();
                            EdiConnectorData.getInstance().OR_COLOR = line.Substring(EdiConnectorData.getInstance().OR_POS[4], EdiConnectorData.getInstance().OR_LEN[4]).Trim();
                            EdiConnectorData.getInstance().OR_LENGTH = line.Substring(EdiConnectorData.getInstance().OR_POS[5], EdiConnectorData.getInstance().OR_LEN[5]).Trim();
                            EdiConnectorData.getInstance().OR_WIDTH = line.Substring(EdiConnectorData.getInstance().OR_POS[6], EdiConnectorData.getInstance().OR_LEN[6]).Trim();
                            EdiConnectorData.getInstance().OR_HEIGHT = line.Substring(EdiConnectorData.getInstance().OR_POS[7], EdiConnectorData.getInstance().OR_LEN[7]).Trim();
                            EdiConnectorData.getInstance().OR_CUX = line.Substring(EdiConnectorData.getInstance().OR_POS[8], EdiConnectorData.getInstance().OR_LEN[8]).Trim();
                            EdiConnectorData.getInstance().OR_PIA = line.Substring(EdiConnectorData.getInstance().OR_POS[9], EdiConnectorData.getInstance().OR_LEN[9]).Trim();
                            EdiConnectorData.getInstance().OR_RFFLI1 = line.Substring(EdiConnectorData.getInstance().OR_POS[10], EdiConnectorData.getInstance().OR_LEN[10]).Trim();
                            EdiConnectorData.getInstance().OR_RFFLI2 = line.Substring(EdiConnectorData.getInstance().OR_POS[11], EdiConnectorData.getInstance().OR_LEN[11]).Trim();
                            if (line.Substring(EdiConnectorData.getInstance().OR_POS[12], EdiConnectorData.getInstance().OR_LEN[12]).Trim().Length == 8)
                                EdiConnectorData.getInstance().OR_DTM_2 = Convert.ToDateTime(line.Substring(EdiConnectorData.getInstance().OR_POS[12], EdiConnectorData.getInstance().OR_LEN[12]).Substring(7, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OR_POS[12], EdiConnectorData.getInstance().OR_LEN[12]).Substring(5, 2) + "-" + line.Substring(EdiConnectorData.getInstance().OR_POS[12], EdiConnectorData.getInstance().OR_LEN[12]).Substring(1, 4));
                            else
                                EdiConnectorData.getInstance().OR_DTM_2 = Convert.ToDateTime("01-01-0001");
                            EdiConnectorData.getInstance().OR_LINNR = line.Substring(EdiConnectorData.getInstance().OR_POS[13], EdiConnectorData.getInstance().OR_LEN[13]).Trim();
                            EdiConnectorData.getInstance().OR_PRI = Convert.ToDouble(line.Substring(EdiConnectorData.getInstance().OR_POS[14], EdiConnectorData.getInstance().OR_LEN[14]).Trim());
                            bool matchItem = WriteSOitems();
                            if (matchItem == false)
                            {
                                File.Move(EdiConnectorData.getInstance().sSOTempPath + @"\" + EdiConnectorData.getInstance().SO_FILENAME, EdiConnectorData.getInstance().sSOErrorPath + @"\" + DateTime.Now.ToString("HHmmss") + "_" + EdiConnectorData.getInstance().SO_FILENAME);
                                Log("X", EdiConnectorData.getInstance().SO_FILENAME + " copied to the errors folder!", "MatchSOdata");
                                return false;
                            }
                            break;
                        } // End switch
                } // End using 

                OrderSave();

                EdiConnectorData.getInstance().SO_FILE = "";
                EdiConnectorData.getInstance().SO_FILENAME = "";
                return true;
            } // End try
            catch(Exception e)
            {
                File.Move(EdiConnectorData.getInstance().sSOTempPath + @"\" + EdiConnectorData.getInstance().SO_FILENAME, EdiConnectorData.getInstance().sSOErrorPath + @"\" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "_" + EdiConnectorData.getInstance().SO_FILENAME);
                Log("X", e.Message, "MatchSOdata");
                Log("X", EdiConnectorData.getInstance().SO_FILENAME + " copied to the errors folder!", "MatchSOdata");
                return false;
            }
        }

        public bool WriteSOhead()
        {
            try
            {
                bool matchHead = CheckSOhead();
                if (matchHead == false)
                    return false;

                EdiConnectorData.getInstance().oOrder = (Documents)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.oOrders);

                EdiConnectorData.getInstance().oOrder.CardCode = EdiConnectorData.getInstance().CARDCODE;
                EdiConnectorData.getInstance().oOrder.NumAtCard = EdiConnectorData.getInstance().OK_BGM;
                EdiConnectorData.getInstance().oOrder.DocDate =  EdiConnectorData.getInstance().OK_K_ORDDAT;
                EdiConnectorData.getInstance().oOrder.TaxDate = EdiConnectorData.getInstance().OK_K_ORDDAT;

                // Reference Rosis DocDueDate
                if (EdiConnectorData.getInstance().OK_DTM_2.Year > 1999)
                    EdiConnectorData.getInstance().oOrder.DocDueDate = EdiConnectorData.getInstance().OK_DTM_2;
                else if(EdiConnectorData.getInstance().OK_DTM_17.Year > 1999)
                    EdiConnectorData.getInstance().oOrder.DocDueDate = EdiConnectorData.getInstance().OK_DTM_17;
                else if (EdiConnectorData.getInstance().OK_DTM_64.Year > 1999)
                    EdiConnectorData.getInstance().oOrder.DocDueDate = EdiConnectorData.getInstance().OK_DTM_64;
                else
                    EdiConnectorData.getInstance().oOrder.DocDueDate = Convert.ToDateTime("31/12/" + DateTime.Now.Year);

                EdiConnectorData.getInstance().oOrder.ShipToCode = EdiConnectorData.getInstance().OK_NAD_DP;
                EdiConnectorData.getInstance().oOrder.PayToCode = EdiConnectorData.getInstance().OK_NAD_IV;

                // Loop UDF
                foreach (Field field in EdiConnectorData.getInstance().oOrder.UserFields.Fields)
                {
                    switch (field.Name)
                    {
                        case "U_BGM": field.Value = EdiConnectorData.getInstance().OK_BGM; break;
                        case "U_RFF": field.Value = EdiConnectorData.getInstance().OK_BGM; break;

                        case "U_K_EANCODE": field.Value = EdiConnectorData.getInstance().OK_K_EANCODE; break;
                        case "U_KNAAM": field.Value = EdiConnectorData.getInstance().OK_KNAAM; break;
                        case "U_TEST":
                            if (EdiConnectorData.getInstance().OK_TEST == "1")
                                field.Value = "J";
                            else
                                field.Value = "N";
                            break;

                        case "U_DTM_2": if (EdiConnectorData.getInstance().OK_DTM_2.ToString() != "1-1-0001 0:00:00") field.Value = EdiConnectorData.getInstance().OK_DTM_2.ToString().Replace(" 0:00:00", ""); break;
                        case "U_TIJD_2": field.Value = EdiConnectorData.getInstance().OK_TIJD_2; break;

                        case "U_DTM_17": if (EdiConnectorData.getInstance().OK_DTM_17.ToString() != "1-1-0001 0:00:00") field.Value = EdiConnectorData.getInstance().OK_DTM_17.ToString().Replace(" 0:00:00", ""); break;
                        case "U_TIJD_17": field.Value = EdiConnectorData.getInstance().OK_TIJD_17; break;

                        case "U_DTM_64": if (EdiConnectorData.getInstance().OK_DTM_64.ToString() != "1-1-0001 0:00:00") field.Value = EdiConnectorData.getInstance().OK_DTM_64.ToString().Replace(" 0:00:00", ""); break;
                        case "U_TIJD_64": field.Value = EdiConnectorData.getInstance().OK_TIJD_64; break;

                        case "U_RFF_BO": field.Value = EdiConnectorData.getInstance().OK_RFF_BO.ToString(); break;
                        case "U_RFF_CR": field.Value = EdiConnectorData.getInstance().OK_RFF_CR.ToString(); break;
                        case "U_RFF_PD": field.Value = EdiConnectorData.getInstance().OK_RFF_PD.ToString(); break;
                        case "U_RFFCT": field.Value = EdiConnectorData.getInstance().OK_RFFCT.ToString(); break;

                        case "U_DTMCT": if (EdiConnectorData.getInstance().OK_DTMCT.ToString() != "1-1-0001 0:00:00") field.Value = EdiConnectorData.getInstance().OK_DTMCT.ToString().Replace(" 0:00:00", ""); break;

                        case "U_FLAG0": field.Value = EdiConnectorData.getInstance().OK_FLAGS[0].ToString(); break;
                        case "U_FLAG1": field.Value = EdiConnectorData.getInstance().OK_FLAGS[1].ToString(); break;
                        case "U_FLAG2": field.Value = EdiConnectorData.getInstance().OK_FLAGS[2].ToString(); break;
                        case "U_FLAG3": field.Value = EdiConnectorData.getInstance().OK_FLAGS[3].ToString(); break;
                        case "U_FLAG4": field.Value = EdiConnectorData.getInstance().OK_FLAGS[4].ToString(); break;
                        case "U_FLAG5": field.Value = EdiConnectorData.getInstance().OK_FLAGS[5].ToString(); break;
                        case "U_FLAG6": field.Value = EdiConnectorData.getInstance().OK_FLAGS[6].ToString(); break;
                        case "U_FLAG7": field.Value = EdiConnectorData.getInstance().OK_FLAGS[7].ToString(); break;
                        case "U_FLAG8": field.Value = EdiConnectorData.getInstance().OK_FLAGS[8].ToString(); break;
                        case "U_FLAG9": field.Value = EdiConnectorData.getInstance().OK_FLAGS[9].ToString(); break;
                        case "U_FLAG10": field.Value = EdiConnectorData.getInstance().OK_FLAGS[10].ToString(); break;

                        case "U_NAD_BY": field.Value = EdiConnectorData.getInstance().OK_NAD_BY.ToString(); break;
                        case "U_NAD_DP": field.Value = EdiConnectorData.getInstance().OK_NAD_DP.ToString(); break;
                        case "U_NAD_IV": field.Value = EdiConnectorData.getInstance().OK_NAD_IV.ToString(); break;
                        case "U_NAD_SF": field.Value = EdiConnectorData.getInstance().OK_NAD_SF.ToString(); break;
                        case "U_NAD_SU": field.Value = EdiConnectorData.getInstance().OK_NAD_SU.ToString(); break;
                        case "U_NAD_UC": field.Value = EdiConnectorData.getInstance().OK_NAD_UC.ToString(); break;
                        case "U_NAD_BCO": field.Value = EdiConnectorData.getInstance().OK_NAD_BCO.ToString(); break;

                        case "U_ONTVANGER": field.Value = EdiConnectorData.getInstance().OK_RECEIVER.ToString(); break;
                        case "U_EDI_BERICHT": field.Value = "Ja"; break;
                        case "U_EDI_IMP_TIJD": field.Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"); break;
                        case "U_EDI_EXPORT": field.Value = "Nee"; break;

                    }
                } // End UDF

                return true;
            }
            catch (Exception e)
            {
                Log("X", e.Message, "WriteSOhead");
                return false;
            }
        }

        public string CheckBuyerAddress(string buyerEANCode)
        {
            string buyerAddress = "";
            Recordset oRecordset;
            oRecordset = (Recordset)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRecordset.DoQuery("SELECT LicTradNum FROM OCRD WHERE U_K_EANCODE = '" + buyerEANCode + "'");
            if (oRecordset.RecordCount == 1)
                buyerAddress = oRecordset.Fields.Item(0).Value.ToString();
            else if (oRecordset.RecordCount > 1)
                Log("X", "Error: Duplicate EANcode Buyers found!", "CheckBuyerAddress");
            else
                Log("X", "Error: Match EANcode Buyers NOT found!", "CheckBuyerAddress");

            oRecordset = null;
            return buyerAddress;
        }

        public bool CheckSOhead()
        {
            EdiConnectorData.getInstance().CARDCODE = "";
            Recordset oRecordset;
            oRecordset = (Recordset)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRecordset.DoQuery("SELECT CardCode FROM OCRD WHERE U_K_EANCODE = '" + EdiConnectorData.getInstance().OK_K_EANCODE + "'");
                if (oRecordset.RecordCount == 1)
                {
                    EdiConnectorData.getInstance().CARDCODE = oRecordset.Fields.Item(0).Value.ToString();
                    return true;
                }
                else if (oRecordset.RecordCount > 1)
                {
                    Log("X", "Error: Duplicate customers found!", "CheckSOhead");
                    return false;
                }
                else
                {
                    Log("X", "Error: Match customer NOT found!", "CheckSOhead");
                    return false;
                }
            }
            finally
            {
                oRecordset = null;
            }
        }

        public bool WriteSOitems()
        {
            try
            {
                bool matchItems = CheckSOitems();
                if (matchItems == false)
                    return false;

                EdiConnectorData.getInstance().oOrder.Lines.ItemCode = EdiConnectorData.getInstance().ITEMCODE;
                string replace = EdiConnectorData.getInstance().OR_QTY.ToString().Replace(",", ".");
                EdiConnectorData.getInstance().oOrder.Lines.Quantity = Convert.ToDouble(String.Format(EdiConnectorData.getInstance().SALPACKUN, replace));
                EdiConnectorData.getInstance().oOrder.Lines.ShipDate = EdiConnectorData.getInstance().OK_DTM_2;

                foreach (Field field in EdiConnectorData.getInstance().oOrder.Lines.UserFields.Fields)
                {
                    switch (field.Name)
                    {
                        case "U_DEUAC": field.Value = EdiConnectorData.getInstance().OR_DEUAC.ToString(); break;
                        case "U_LEVARTCODE": field.Value = EdiConnectorData.getInstance().OR_LEVARTCODE.ToString(); break;
                        case "U_DEARTOM": field.Value = EdiConnectorData.getInstance().OR_DEARTOM.ToString(); break;
                        case "U_KLEUR": field.Value = EdiConnectorData.getInstance().OR_COLOR.ToString(); break;
                        case "U_LENGTE": field.Value = EdiConnectorData.getInstance().OR_LENGTH.ToString(); break;
                        case "U_BREEDTE": field.Value = EdiConnectorData.getInstance().OR_WIDTH.ToString(); break;
                        case "U_HOOGTE": field.Value = EdiConnectorData.getInstance().OR_HEIGHT.ToString(); break;
                        case "U_CUX": field.Value = EdiConnectorData.getInstance().OR_CUX.ToString(); break;
                        case "U_PIA": field.Value = EdiConnectorData.getInstance().OR_PIA.ToString(); break;
                        case "U_RFFLI1": field.Value = EdiConnectorData.getInstance().OR_RFFLI1.ToString(); break;
                        case "U_RFFLI2": field.Value = EdiConnectorData.getInstance().OR_RFFLI2.ToString(); break;
                        case "U_LINNR": field.Value = EdiConnectorData.getInstance().OR_LINNR.ToString(); break;
                        case "U_PRI": field.Value = EdiConnectorData.getInstance().OR_PRI.ToString(); break;

                    }
                }
                EdiConnectorData.getInstance().oOrder.Lines.Add();
                return true;
            }
            catch (Exception e)
            {
                Log("X", e.Message, "WriteSOitems");
                return false;
            }
        }

        public bool CheckSOitems()
        {
            EdiConnectorData.getInstance().ITEMCODE = "";
            EdiConnectorData.getInstance().ITEMNAME = "";
            EdiConnectorData.getInstance().SALPACKUN = "1";
            Recordset oRecordset;
            oRecordset = (Recordset)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRecordset.DoQuery("SELECT ItemCode, ItemName, SalPackUn FROM OITM WHERE U_EAN_Handels_EH = '" + EdiConnectorData.getInstance().OR_DEUAC + "'");

                if (oRecordset.RecordCount == 1)
                {
                    EdiConnectorData.getInstance().ITEMCODE = oRecordset.Fields.Item(0).Value.ToString();
                    EdiConnectorData.getInstance().ITEMNAME = oRecordset.Fields.Item(1).Value.ToString();
                    EdiConnectorData.getInstance().SALPACKUN = oRecordset.Fields.Item(2).Value.ToString();
                    return true;
                }
                else if (oRecordset.RecordCount > 1)
                {
                    Log("X", "Error: Duplicate EANcodes found! Eancode " + EdiConnectorData.getInstance().OR_DEUAC, "CheckSOitems");
                    return false;
                }
                else
                {
                    Log("X", "Error: Match EANcode NOT found!", "CheckSOitems");
                    return false;
                }
            }
            finally
            {
                oRecordset = null;
            }
        }

        public void OrderSave()
        {
            int errCode;
            string errMsg = "";

            if (EdiConnectorData.getInstance().oOrder.Add() != 0)
            {
                EdiConnectorData.getInstance().cmp.GetLastError(out errCode, out errMsg);

                File.Move(EdiConnectorData.getInstance().sSOTempPath + @"\" + EdiConnectorData.getInstance().SO_FILENAME, EdiConnectorData.getInstance().sSOErrorPath + @"\" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "_" + EdiConnectorData.getInstance().SO_FILENAME);
                Log("X", "Error: " + errMsg + "(" + errCode + ")", "OrderSave");
                Log("X", EdiConnectorData.getInstance().SO_FILENAME + " copied to the errors folder!", "OrderSave");
            }
            else
            {
                if (EdiConnectorData.getInstance().iSendNotification == 1)
                    MailToSOreceiver(EdiConnectorData.getInstance().oOrder.DocEntry, EdiConnectorData.getInstance().oOrder.DocDate.ToString(), EdiConnectorData.getInstance().oOrder.DocNum);

                Log("V", "Order written in SAP BO!", "OrderSave");
                File.Move(EdiConnectorData.getInstance().sSOTempPath + @"\" + EdiConnectorData.getInstance().SO_FILENAME, EdiConnectorData.getInstance().sSODonePath + @"\" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "_" + EdiConnectorData.getInstance().SO_FILENAME);
            }   
        }

        public void CreateUdfFieldsText()
        {
            if(File.Exists(EdiConnectorData.getInstance().sApplicationPath + @"\udf.xml") == false)
            {
                return;
            }

            DataSet dataSet = new DataSet();
            dataSet.ReadXml(EdiConnectorData.getInstance().sApplicationPath + @"\udf.xml");

            try
            {
                if (dataSet.Tables["udf"].Rows.Count > 0)
                {
                    for (int i = 0; i < dataSet.Tables["udf"].Rows.Count - 1; i++)
                        CreateUdf(dataSet.Tables["udf"].Rows[i][0].ToString(), dataSet.Tables["udf"].Rows[i][1].ToString(), dataSet.Tables["udf"].Rows[i][2].ToString(), 
                            BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, Convert.ToInt32(dataSet.Tables["udf"].Rows[i][3]), false, false, "");
                }
            }
            catch (Exception e)
            {
                Log("X", "Error: " + e.Message, "CreateUdfFieldsText");
                dataSet.Dispose();
            }
        }

        private void CreateUdf(string _tableName, string _fieldName, string _description,
            BoFieldTypes _boType, BoFldSubTypes _boSubType, int _editSize, bool _mandatory, bool _default, string _defaultValue)
        {
            IUserFieldsMD oUDF;
            oUDF = (SAPbobsCOM.IUserFieldsMD)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.oUserFields);
            int errCode;
            string errMsg = "";

            oUDF.TableName = _tableName;
            oUDF.Name = _fieldName;
            oUDF.Description = _description;
            oUDF.Type = _boType;
            oUDF.SubType = _boSubType;
            oUDF.EditSize = _editSize;
            oUDF.Size = _editSize;
            if (_default == true)
                oUDF.DefaultValue = _defaultValue;

            if (_mandatory == true)
                oUDF.Mandatory = BoYesNoEnum.tYES;
            else
                oUDF.Mandatory = BoYesNoEnum.tNO;

            if (oUDF.Add() != 0)
            {
                EdiConnectorData.getInstance().cmp.GetLastError(out errCode, out errMsg);
                Log("X", "Error: " + errMsg + "(" + errCode + ")", "CreateUdf");
            }
            else
            {
                Log("V", "Udf " + _fieldName + " successfully created!", "CreateUdf");
            }

            oUDF = null;
        }

        public bool MailToSOreceiver(int _docEntry, string _docDate, int _docNumber)
        {
            try
            {
                using (MailMessage mailMsg = new MailMessage())
                {
                    SmtpClient smtpMail = new SmtpClient(EdiConnectorData.getInstance().sSmpt, EdiConnectorData.getInstance().iSmtpPort);

                    if (EdiConnectorData.getInstance().bSmtpUserSecurity == true)
                        smtpMail.Credentials = new NetworkCredential(EdiConnectorData.getInstance().sSmtpUser, EdiConnectorData.getInstance().sSmtpPassword);

                    mailMsg.From = new MailAddress(EdiConnectorData.getInstance().sSenderEmail, EdiConnectorData.getInstance().sSenderName);
                    mailMsg.To.Add(EdiConnectorData.getInstance().sOrderMailTo);
                    mailMsg.Subject = String.Format("Order {0} imported", _docEntry);

                    using(StreamReader reader = new StreamReader(EdiConnectorData.getInstance().sApplicationPath + @"\email_o.txt"))
                    {
                        string body = reader.ReadToEnd();
                        body = body.Replace("::NAME::", EdiConnectorData.getInstance().sOrderMailToFullName);
                        body = body.Replace("::DOCENTRY::", _docEntry.ToString());
                        body = body.Replace("::DOCDATE::", _docDate);
                        body = body.Replace("::DOCNUM::", _docNumber.ToString());
                        reader.Close();

                        mailMsg.Body = body;
                    }// End using reader

                    smtpMail.Send(mailMsg);
                }// End using mailMsg

                Log("V", "Order notification sent!", "MailToSOreceiver");
                return true;
            }
            catch
            {
                Log("X", "Order notification was not sent!", "MailToSOreceiver");
                return false;
            }
        }

        public void CheckAndExportDelivery()
        {
            Recordset oRecordset;
            oRecordset = (Recordset)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                oRecordset.DoQuery("SELECT DocEntry FROM EDLN WHERE Canceled='N' AND U_EDI_BERICHT = 'Ja' " +
                    "AND U_EDI_EXPORT = 'Ja' AND (U_EDI_DEL_EXP is NULL OR U_EDI_DEL_EXP = 'Nee')");

                if (oRecordset.RecordCount == 0)
                    Log("V", "No Delivery notes found to export!", "CheckAndExportDelivery");
                else if (oRecordset.RecordCount > 0)
                {
                    oRecordset.MoveFirst();
                    for (int i = 1; i < oRecordset.RecordCount; i++)
                    {
                        CreateDeliveryFile(Convert.ToInt32(oRecordset.Fields.Item(0).Value));
                        oRecordset.MoveNext();
                    }
                }
            }
            catch (Exception e)
            {
                Log("X", e.Message, "CheckAndExportDelivery");
            }
            finally
            {
                oRecordset = null;
            }
        }

        public void CreateDeliveryFile(int _docEntry)
        {
            Documents oDelivery;
            oDelivery = (Documents)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
            bool foundKey = oDelivery.GetByKey(_docEntry);

            if (oDelivery.UserFields.Fields.Item("U_DESADV").Value.ToString() == "1")
            {
                if (foundKey == true)
                {
                    try
                    {
                        using (StreamWriter writer = new StreamWriter(EdiConnectorData.getInstance().sDeliveryPath + @"\" + "DEL_" + oDelivery.DocNum + ".DAT"))
                        {

                            string header = "";
                            string line = "";

                            EdiConnectorData.getInstance().DK_K_EANCODE = "";
                            EdiConnectorData.getInstance().DK_PBTEST = "";
                            EdiConnectorData.getInstance().DK_KNAAM = "";
                            EdiConnectorData.getInstance().DK_BGM = "";
                            EdiConnectorData.getInstance().DK_DTM_137 = "";
                            EdiConnectorData.getInstance().DK_DTM_2 = "";
                            EdiConnectorData.getInstance().DK_TIJD_2 = "";
                            EdiConnectorData.getInstance().DK_DTM_17 = "";
                            EdiConnectorData.getInstance().DK_TIJD_17 = "";
                            EdiConnectorData.getInstance().DK_DTM_64 = "";
                            EdiConnectorData.getInstance().DK_TIJD_64 = "";
                            EdiConnectorData.getInstance().DK_DTM_63 = "";
                            EdiConnectorData.getInstance().DK_TIJD_63 = "";
                            EdiConnectorData.getInstance().DK_BH_DAT = "";
                            EdiConnectorData.getInstance().DK_BH_TIJD = "";
                            EdiConnectorData.getInstance().DK_RFF = "";
                            EdiConnectorData.getInstance().DK_RFFVN = "";
                            EdiConnectorData.getInstance().DK_BH_EAN = "";
                            EdiConnectorData.getInstance().DK_NAD_BY = "";
                            EdiConnectorData.getInstance().DK_NAD_SU = "";
                            EdiConnectorData.getInstance().DK_DESATYPE = "";
                            EdiConnectorData.getInstance().DK_ONTVANGER = "";

                            EdiConnectorData.getInstance().DK_K_EANCODE = FormatField(oDelivery.UserFields.Fields.Item("U_K_EANCODE").Value.ToString(), 13, "STRING");

                            if (oDelivery.UserFields.Fields.Item("U_TEST").Value.ToString() == "J")
                                EdiConnectorData.getInstance().DK_PBTEST = FormatField("1", 1, "STRING");
                            else
                                EdiConnectorData.getInstance().DK_PBTEST = FormatField(" ", 1, "STRING");

                            EdiConnectorData.getInstance().DK_KNAAM = FormatField(oDelivery.UserFields.Fields.Item("U_KNAAM").Value.ToString(), 14, "STRING");
                            EdiConnectorData.getInstance().DK_BGM = FormatField(oDelivery.DocNum.ToString(), 35, "STRING");
                            EdiConnectorData.getInstance().DK_DTM_137 = FormatField(oDelivery.DocDate.ToString("yyyyMMdd"), 8, "STIRNG");

                            if (oDelivery.UserFields.Fields.Item("U_DTM_2").Value.ToString().Trim().Length > 0)
                                EdiConnectorData.getInstance().DK_DTM_2 = FormatField(Convert.ToDateTime(oDelivery.UserFields.Fields.Item("U_DTM_2").Value).ToString("yyyyMMdd"), 8, "STRING");
                            else
                                EdiConnectorData.getInstance().DK_DTM_2 = FormatField(" ", 8, "STRING");
                            EdiConnectorData.getInstance().DK_TIJD_2 = FormatField(oDelivery.UserFields.Fields.Item("U_TIJD_2").Value.ToString(), 5, "STRING");

                            if (oDelivery.UserFields.Fields.Item("U_DTM_17").Value.ToString().Trim().Length > 0)
                                EdiConnectorData.getInstance().DK_DTM_17 = FormatField(Convert.ToDateTime(oDelivery.UserFields.Fields.Item("U_DTM_17").Value).ToString("yyyyMMdd"), 8, "STRING");
                            else
                                EdiConnectorData.getInstance().DK_DTM_17 = FormatField(" ", 8, "STRING");
                            EdiConnectorData.getInstance().DK_TIJD_17 = FormatField(oDelivery.UserFields.Fields.Item("U_TIJD_17").Value.ToString(), 5, "STRING");

                            if (oDelivery.UserFields.Fields.Item("U_DTM_64").Value.ToString().Trim().Length > 0)
                                EdiConnectorData.getInstance().DK_DTM_64 = FormatField(Convert.ToDateTime(oDelivery.UserFields.Fields.Item("U_DTM_64").Value).ToString("yyyyMMdd"), 8, "STRING");
                            else
                                EdiConnectorData.getInstance().DK_DTM_64 = FormatField(" ", 8, "STRING");
                            EdiConnectorData.getInstance().DK_TIJD_64 = FormatField(oDelivery.UserFields.Fields.Item("U_TIJD_64").Value.ToString(), 5, "STRING");

                            if (oDelivery.UserFields.Fields.Item("U_DTM_63").Value.ToString().Trim().Length > 0)
                                EdiConnectorData.getInstance().DK_DTM_63 = FormatField(Convert.ToDateTime(oDelivery.UserFields.Fields.Item("U_DTM_63").Value).ToString("yyyyMMdd"), 8, "STRING");
                            else
                                EdiConnectorData.getInstance().DK_DTM_63 = FormatField(" ", 8, "STRING");
                            EdiConnectorData.getInstance().DK_TIJD_63 = FormatField(oDelivery.UserFields.Fields.Item("U_TIJD_63").Value.ToString(), 5, "STRING");

                            EdiConnectorData.getInstance().DK_BH_DAT = FormatField(oDelivery.UserFields.Fields.Item("U_BH_DAT").Value.ToString(), 8, "STRING");
                            EdiConnectorData.getInstance().DK_BH_TIJD = FormatField(oDelivery.UserFields.Fields.Item("U_BH_TIJD").Value.ToString(), 5, "STRING");
                            EdiConnectorData.getInstance().DK_RFF = FormatField(oDelivery.UserFields.Fields.Item("U_RFF").Value.ToString(), 35, "STRING");
                            EdiConnectorData.getInstance().DK_RFFVN = FormatField(oDelivery.UserFields.Fields.Item("U_RFFVN").Value.ToString(), 35, "STRING");
                            EdiConnectorData.getInstance().DK_BH_EAN = FormatField(oDelivery.UserFields.Fields.Item("U_BH_EAN").Value.ToString(), 13, "STRING");
                            EdiConnectorData.getInstance().DK_NAD_BY = FormatField(oDelivery.UserFields.Fields.Item("U_NAD_BY").Value.ToString(), 13, "STRING");
                            EdiConnectorData.getInstance().DK_NAD_DP = FormatField(oDelivery.UserFields.Fields.Item("U_NAD_DP").Value.ToString(), 13, "STRING");
                            EdiConnectorData.getInstance().DK_NAD_SU = FormatField(oDelivery.UserFields.Fields.Item("U_NAD_SU").Value.ToString(), 13, "STRING");
                            EdiConnectorData.getInstance().DK_NAD_UC = FormatField(oDelivery.UserFields.Fields.Item("U_NAD_UC").Value.ToString(), 13, "STRING");
                            //DK_DESATYPE = strDesAdvLevel //see desadvlevel in settings file
                            EdiConnectorData.getInstance().DK_DESATYPE = oDelivery.UserFields.Fields.Item("U_DESADV").Value.ToString();
                            EdiConnectorData.getInstance().DK_ONTVANGER = FormatField(oDelivery.UserFields.Fields.Item("U_ONTVANGER").Value.ToString(), 13, "STRING");

                            header = EdiConnectorData.getInstance().DK_K_EANCODE +
                                EdiConnectorData.getInstance().DK_PBTEST +
                                EdiConnectorData.getInstance().DK_KNAAM +
                                EdiConnectorData.getInstance().DK_BGM +
                                EdiConnectorData.getInstance().DK_DTM_137 +
                                EdiConnectorData.getInstance().DK_DTM_2 +
                                EdiConnectorData.getInstance().DK_TIJD_2 +
                                EdiConnectorData.getInstance().DK_DTM_17 +
                                EdiConnectorData.getInstance().DK_TIJD_17 +
                                EdiConnectorData.getInstance().DK_DTM_64 +
                                EdiConnectorData.getInstance().DK_TIJD_64 +
                                EdiConnectorData.getInstance().DK_DTM_63 +
                                EdiConnectorData.getInstance().DK_TIJD_63 +
                                EdiConnectorData.getInstance().DK_BH_DAT +
                                EdiConnectorData.getInstance().DK_BH_TIJD +
                                EdiConnectorData.getInstance().DK_RFF +
                                EdiConnectorData.getInstance().DK_RFFVN +
                                EdiConnectorData.getInstance().DK_BH_EAN +
                                EdiConnectorData.getInstance().DK_NAD_BY +
                                EdiConnectorData.getInstance().DK_NAD_DP +
                                EdiConnectorData.getInstance().DK_NAD_SU +
                                EdiConnectorData.getInstance().DK_NAD_UC +
                                EdiConnectorData.getInstance().DK_DESATYPE +
                                EdiConnectorData.getInstance().DK_ONTVANGER;

                            writer.WriteLine("0" + header);

                            for (int i = 0; i < oDelivery.Lines.Count - 1; i++)
                            {
                                line = "";

                                EdiConnectorData.getInstance().DR_DEUAC = "";
                                EdiConnectorData.getInstance().DR_OLDUAC = "";
                                EdiConnectorData.getInstance().DR_DEARTNR = "";
                                EdiConnectorData.getInstance().DR_DEARTOM = "";
                                EdiConnectorData.getInstance().DR_PIA = "";
                                EdiConnectorData.getInstance().DR_BATCH = "";
                                EdiConnectorData.getInstance().DR_QTY = "";
                                EdiConnectorData.getInstance().DR_ARTEENHEID = "";
                                EdiConnectorData.getInstance().DR_RFFONID = "";
                                EdiConnectorData.getInstance().DR_RFFONORD = "";
                                EdiConnectorData.getInstance().DR_DTM_23E = "";
                                EdiConnectorData.getInstance().DR_TGTDATUM = "";
                                EdiConnectorData.getInstance().DR_GEWICHT = "";
                                EdiConnectorData.getInstance().DR_FEENHEID = "";
                                EdiConnectorData.getInstance().DR_QTY_AFW = "";
                                EdiConnectorData.getInstance().DR_REDEN = "";
                                EdiConnectorData.getInstance().DR_GINTYPE = "";
                                EdiConnectorData.getInstance().DR_GINID = "";
                                EdiConnectorData.getInstance().DR_BATCHH = "";

                                oDelivery.Lines.SetCurrentLine(i);

                                EdiConnectorData.getInstance().DR_DEUAC = FormatField(ItemEAN(oDelivery.Lines.ItemCode), 14, "STRING");
                                EdiConnectorData.getInstance().DR_OLDUAC = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_OLDUAC").Value.ToString(), 14, "STRING");
                                EdiConnectorData.getInstance().DR_DEARTNR = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_DEARTNR").Value.ToString(), 9, "STRING");
                                EdiConnectorData.getInstance().DR_DEARTOM = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_DEARTOM").Value.ToString(), 35, "STRING");
                                EdiConnectorData.getInstance().DR_PIA = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_PIA").Value.ToString(), 10, "STRING");
                                EdiConnectorData.getInstance().DR_BATCH = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_BATCH").Value.ToString(), 35, "STRING");

                                EdiConnectorData.getInstance().DR_QTY = FormatField(oDelivery.Lines.Quantity.ToString(), 17, "STRING");

                                EdiConnectorData.getInstance().DR_ARTEENHEID = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_ARTEENHEID").Value.ToString(), 3, "STRING");
                                EdiConnectorData.getInstance().DR_RFFONID = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_RFFONID").Value.ToString(), 6, "STRING");
                                EdiConnectorData.getInstance().DR_RFFONORD = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_RFFONORD").Value.ToString(), 35, "STRING");
                                EdiConnectorData.getInstance().DR_DTM_23E = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_DTM_23E").Value.ToString(), 8, "STRING");
                                EdiConnectorData.getInstance().DR_TGTDATUM = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_TGTDATUM").Value.ToString(), 1, "STRING");
                                EdiConnectorData.getInstance().DR_GEWICHT = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_GEWICHT").Value.ToString(), 6, "STRING");
                                EdiConnectorData.getInstance().DR_FEENHEID = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_FEENHEID").Value.ToString(), 3, "STRING");
                                EdiConnectorData.getInstance().DR_QTY_AFW = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_QTY_AFW").Value.ToString(), 17, "STRING");
                                EdiConnectorData.getInstance().DR_REDEN = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_REDEN").Value.ToString(), 3, "STRING");
                                EdiConnectorData.getInstance().DR_GINTYPE = FormatField(" ", 3, "STRING");
                                EdiConnectorData.getInstance().DR_GINID = FormatField(" ", 19, "STRING");
                                EdiConnectorData.getInstance().DR_BATCHH = FormatField(oDelivery.Lines.UserFields.Fields.Item("U_BATCHH").Value.ToString(), 35, "STRING");

                                line = EdiConnectorData.getInstance().DR_DEUAC +
                                    EdiConnectorData.getInstance().DR_OLDUAC +
                                    EdiConnectorData.getInstance().DR_DEARTNR +
                                    EdiConnectorData.getInstance().DR_DEARTOM +
                                    EdiConnectorData.getInstance().DR_PIA +
                                    EdiConnectorData.getInstance().DR_BATCH +
                                    EdiConnectorData.getInstance().DR_QTY +
                                    EdiConnectorData.getInstance().DR_ARTEENHEID +
                                    EdiConnectorData.getInstance().DR_RFFONID +
                                    EdiConnectorData.getInstance().DR_RFFONORD +
                                    EdiConnectorData.getInstance().DR_DTM_23E +
                                    EdiConnectorData.getInstance().DR_TGTDATUM +
                                    EdiConnectorData.getInstance().DR_GEWICHT +
                                    EdiConnectorData.getInstance().DR_FEENHEID +
                                    EdiConnectorData.getInstance().DR_QTY_AFW +
                                    EdiConnectorData.getInstance().DR_REDEN +
                                    EdiConnectorData.getInstance().DR_GINTYPE +
                                    EdiConnectorData.getInstance().DR_GINID +
                                    EdiConnectorData.getInstance().DR_BATCHH;

                                if (oDelivery.Lines.TreeType != BoItemTreeTypes.iIngredient)
                                    writer.WriteLine("1" + line);
                            }// End for loop

                            writer.Close();

                            oDelivery.UserFields.Fields.Item("U_EDI_DEL_EXP").Value = "Ja";
                            oDelivery.UserFields.Fields.Item("U_EDI_DELEXP_TIJD").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                            oDelivery.Update();

                            if (EdiConnectorData.getInstance().iSendNotification == 1)
                                MailToDLreceiver(oDelivery.DocEntry, oDelivery.DocDate.ToString(), oDelivery.DocNum);

                            Log("V", "Delivery note file created!", "CreateDeliveryFile");
                        }// End using
                    }// End try
                    catch (Exception e)
                    {
                        File.Delete(EdiConnectorData.getInstance().sDeliveryPath + @"\" + "DEL_" + oDelivery.DocNum + ".DAT");
                        Log("X", e.Message, "CreateDeliveryFile");
                    }

                }// End if foundKey
                else
                {
                    Log("X", "Delivery note " + _docEntry + " not found!", "CreateDeliveryFile");
                }
            }// End if "U_DESADV"
        }

        public string SSCC(int _docEntry, int _lineNr, string _itemCode)
        {
            string newSSCC;
            Recordset oRecordset;
            oRecordset = (Recordset)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRecordset.DoQuery("SELECT SSCC_Code FROM SSCC WHERE DocEntry = " + _docEntry + " AND RegelNr = " + _lineNr + " AND ItemCode = '" + _itemCode + "'");
            if (oRecordset.RecordCount > 0)
            {
                newSSCC = oRecordset.Fields.Item(0).Value.ToString();
                Log("V", "SSCC code found!", "SSCC");
            }
            else
            {
                newSSCC = "";
                Log("X", "SSCC code for document " + _docEntry + " not found!", "SSCC");
            }

            oRecordset = null;
            return newSSCC;
        }

        public string SSCCqty(int _docEntry, int _lineNr)
        {
            string newSSCCqty;
            Recordset oRecordset;
            oRecordset = (Recordset)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRecordset.DoQuery("SELECT Quantity FROM SSCC WHERE DocEntry = " + _docEntry + " AND RegelNr = " + _lineNr);
            if (oRecordset.RecordCount > 0)
            {
                newSSCCqty = oRecordset.Fields.Item(0).Value.ToString();
                Log("V", "SSCC code (qty) found!", "SSCC");
            }
            else
            {
                newSSCCqty = "";
                Log("X", "SSCC code (qty) for document " + _docEntry + " not found!", "SSCC");
            }

            oRecordset = null;
            return newSSCCqty;
        }

        public string ItemEAN(string _iCode)
        {
            string i_EAN;
            Recordset oRecordset;
            oRecordset = (Recordset)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRecordset.DoQuery("SELECT U_EAN_Handels_EH FROM OITM WHERE ItemCode = '" + _iCode + "'");
            if (oRecordset.RecordCount > 0)
            {
                i_EAN = oRecordset.Fields.Item(0).Value.ToString();
                Log("V", "ItemEAN code found!", "ItemEAN");
            }
            else
            {
                i_EAN = "";
                Log("X", "ItemEAN not found!", "ItemEAN");
            }

            oRecordset = null;
            return i_EAN;
        }

        public void CheckAndExportInvoice()
        {
            Recordset oRecordset;
            oRecordset = (Recordset)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRecordset.DoQuery("SELECT DocEntry FROM OINV WHERE U_EDI_BERICHT = 'Ja' AND U_EDI_EXPORT = 'Ja' AND (U_EDI_INV_EXP is NULL OR U_EDI_INV_EXP = 'Nee')");
                if (oRecordset.RecordCount == 0)
                    Log("V", "No Invoices note found to export!", "CheckAndExportInvoice");
                else if (oRecordset.RecordCount > 0)
                {
                    oRecordset.MoveFirst();
                    for (int i = 1; i < oRecordset.RecordCount; i++)
                    {
                        CreateInvoiceFile(oRecordset.Fields.Item(0).Value.ToString());
                        oRecordset.MoveNext();
                    }
                }
            }
            catch (Exception e)
            {
                Log("X", e.Message, "CheckAndExportInvoice");
            }
            finally
            {
                oRecordset = null;
            }
        }

        public void CreateInvoiceFile(string _docEntry)
        {
            Documents oInvoice;
            oInvoice = (Documents)EdiConnectorData.getInstance().cmp.GetBusinessObject(BoObjectTypes.oInvoices);
            bool foundKey = oInvoice.GetByKey(Convert.ToInt32(_docEntry));

            if (foundKey == true)
            {
                try
                {
                    using (StreamWriter writer = new StreamWriter(EdiConnectorData.getInstance().sInvoicePath + @"\" + "INV_" + oInvoice.DocNum + ".DAT"))
                    {
                        string header = "";
                        string line = "";

                        string alcHeader = "";
                        //string alcLine = "";

                        EdiConnectorData.getInstance().FK_K_EANCODE = "";
                        EdiConnectorData.getInstance().FK_FAKTEST = "";
                        EdiConnectorData.getInstance().FK_KNAAM = "";
                        EdiConnectorData.getInstance().FK_F_SOORT = "";
                        EdiConnectorData.getInstance().FK_FAKT_NUM = "";
                        EdiConnectorData.getInstance().FK_FAKT_DATUM = "";
                        EdiConnectorData.getInstance().FK_AFL_DATUM = "";
                        EdiConnectorData.getInstance().FK_RFFIV = "";
                        EdiConnectorData.getInstance().FK_K_ORDERNR = "";
                        EdiConnectorData.getInstance().FK_K_ORDDAT = "";
                        EdiConnectorData.getInstance().FK_PAKBONNR = "";
                        EdiConnectorData.getInstance().FK_RFFCDN = "";
                        EdiConnectorData.getInstance().FK_RFFALO = "";
                        EdiConnectorData.getInstance().FK_RFFVN = "";
                        EdiConnectorData.getInstance().FK_RFFVNDAT = "";
                        EdiConnectorData.getInstance().FK_NAD_BY = "";
                        EdiConnectorData.getInstance().FK_A_EANCODE = "";
                        EdiConnectorData.getInstance().FK_F_EANCODE = "";
                        EdiConnectorData.getInstance().FK_NAD_SF = "";
                        EdiConnectorData.getInstance().FK_NAD_SU = "";
                        EdiConnectorData.getInstance().FK_NAD_UC = "";
                        EdiConnectorData.getInstance().FK_NAD_PE = "";
                        EdiConnectorData.getInstance().FK_OBNUMMER = "";
                        EdiConnectorData.getInstance().FK_ACT = "";
                        EdiConnectorData.getInstance().FK_CUX = "";
                        EdiConnectorData.getInstance().FK_DAGEN = "";
                        EdiConnectorData.getInstance().FK_KORTPERC = "";
                        EdiConnectorData.getInstance().FK_KORTBEDR = "";
                        EdiConnectorData.getInstance().FK_ONTVANGER = "";

                        EdiConnectorData.getInstance().FK_K_EANCODE = FormatField(oInvoice.UserFields.Fields.Item("U_K_EANCODE").Value.ToString(), 13, "STRING");

                        if (oInvoice.UserFields.Fields.Item("U_TEST").Value.ToString() == "J")
                            EdiConnectorData.getInstance().FK_FAKTEST = FormatField("1", 1, "STRING");
                        else
                            EdiConnectorData.getInstance().FK_FAKTEST = FormatField(" ", 1, "STRING");

                        EdiConnectorData.getInstance().FK_KNAAM = FormatField(oInvoice.UserFields.Fields.Item("U_KNAAM").Value.ToString(), 14, "STRING");

                        if (oInvoice.UserFields.Fields.Item("U_F_SOORT").Value.ToString().Trim().Length > 0)
                            EdiConnectorData.getInstance().FK_F_SOORT = FormatField(oInvoice.UserFields.Fields.Item("U_F_SOORT").Value.ToString(), 3, "STRING");
                        else
                            EdiConnectorData.getInstance().FK_F_SOORT = FormatField("380", 3, "STRING");

                        EdiConnectorData.getInstance().FK_FAKT_NUM = FormatField(oInvoice.DocNum.ToString(), 12, "STRING");

                        EdiConnectorData.getInstance().FK_FAKT_DATUM = FormatField(oInvoice.DocDate.ToString("yyyyMMdd"), 8, "STRING");
                        EdiConnectorData.getInstance().FK_AFL_DATUM = FormatField(Convert.ToDateTime(oInvoice.UserFields.Fields.Item("U_AFL_DATUM").Value).ToString("yyyyMMdd"), 8, "STRING");
                        EdiConnectorData.getInstance().FK_RFFIV = FormatField(oInvoice.UserFields.Fields.Item("U_RFFIV").Value.ToString(), 35, "STRING");
                        EdiConnectorData.getInstance().FK_K_ORDERNR = FormatField(oInvoice.UserFields.Fields.Item("U_K_ORDERNR").Value.ToString(), 35, "STRING");
                        EdiConnectorData.getInstance().FK_K_ORDDAT = FormatField(oInvoice.UserFields.Fields.Item("U_K_ORDDAT").Value.ToString(), 8, "STRING");
                        EdiConnectorData.getInstance().FK_PAKBONNR = FormatField(oInvoice.UserFields.Fields.Item("U_PAKBONNR").Value.ToString(), 35, "STRING");
                        EdiConnectorData.getInstance().FK_RFFCDN = FormatField(oInvoice.UserFields.Fields.Item("U_RFFCDN").Value.ToString(), 35, "STRING");
                        EdiConnectorData.getInstance().FK_RFFALO = FormatField(oInvoice.UserFields.Fields.Item("U_RFFALO").Value.ToString(), 35, "STRING");
                        EdiConnectorData.getInstance().FK_RFFVN = FormatField(oInvoice.UserFields.Fields.Item("U_RFFVN").Value.ToString(), 35, "STRING");
                        EdiConnectorData.getInstance().FK_RFFVNDAT = FormatField(oInvoice.UserFields.Fields.Item("U_RFFVNDAT").Value.ToString(), 8, "STRING");
                        EdiConnectorData.getInstance().FK_NAD_BY = FormatField(oInvoice.UserFields.Fields.Item("U_NAD_BY").Value.ToString(), 13, "STRING");
                        EdiConnectorData.getInstance().FK_A_EANCODE = FormatField(oInvoice.ShipToCode, 13, "STRING");
                        EdiConnectorData.getInstance().FK_F_EANCODE = FormatField(oInvoice.PayToCode, 13, "STRING");
                        EdiConnectorData.getInstance().FK_NAD_SF = FormatField(oInvoice.UserFields.Fields.Item("U_NAD_SF").Value.ToString(), 13, "STRING");
                        EdiConnectorData.getInstance().FK_NAD_SU = FormatField(oInvoice.UserFields.Fields.Item("U_NAD_SU").Value.ToString(), 13, "STRING");
                        EdiConnectorData.getInstance().FK_NAD_UC = FormatField(oInvoice.UserFields.Fields.Item("U_NAD_UC").Value.ToString(), 13, "STRING");
                        EdiConnectorData.getInstance().FK_NAD_PE = FormatField(oInvoice.UserFields.Fields.Item("U_NAD_PE").Value.ToString(), 13, "STRING");
                        EdiConnectorData.getInstance().FK_OBNUMMER = FormatField(CheckBuyerAddress(oInvoice.UserFields.Fields.Item("U_K_EANCODE").Value.ToString()), 15, "STRING");

                        EdiConnectorData.getInstance().FK_ACT = FormatField(oInvoice.UserFields.Fields.Item("U_ACT").Value.ToString(), 1, "STRING");
                        EdiConnectorData.getInstance().FK_CUX = FormatField(oInvoice.DocCurrency, 3, "STRING");
                        EdiConnectorData.getInstance().FK_DAGEN = FormatField(oInvoice.UserFields.Fields.Item("U_DAGEN").Value.ToString(), 3, "STRING");
                        EdiConnectorData.getInstance().FK_KORTPERC = FormatField(oInvoice.UserFields.Fields.Item("U_KORTPERC").Value.ToString(), 8, "STRING");
                        EdiConnectorData.getInstance().FK_KORTBEDR = FormatField(oInvoice.UserFields.Fields.Item("U_KORTBEDR").Value.ToString(), 9, "STRING");
                        EdiConnectorData.getInstance().FK_ONTVANGER = FormatField(oInvoice.UserFields.Fields.Item("U_ONTVANGER").Value.ToString(), 13, "STRING");

                        header = EdiConnectorData.getInstance().FK_K_EANCODE +
                            EdiConnectorData.getInstance().FK_FAKTEST +
                            EdiConnectorData.getInstance().FK_KNAAM +
                            EdiConnectorData.getInstance().FK_F_SOORT +
                            EdiConnectorData.getInstance().FK_FAKT_NUM +
                            EdiConnectorData.getInstance().FK_FAKT_DATUM +
                            EdiConnectorData.getInstance().FK_AFL_DATUM +
                            EdiConnectorData.getInstance().FK_RFFIV +
                            EdiConnectorData.getInstance().FK_K_ORDERNR +
                            EdiConnectorData.getInstance().FK_K_ORDDAT +
                            EdiConnectorData.getInstance().FK_PAKBONNR +
                            EdiConnectorData.getInstance().FK_RFFCDN +
                            EdiConnectorData.getInstance().FK_RFFALO +
                            EdiConnectorData.getInstance().FK_RFFVN +
                            EdiConnectorData.getInstance().FK_RFFVNDAT +
                            EdiConnectorData.getInstance().FK_NAD_BY +
                            EdiConnectorData.getInstance().FK_A_EANCODE +
                            EdiConnectorData.getInstance().FK_F_EANCODE +
                            EdiConnectorData.getInstance().FK_NAD_SF +
                            EdiConnectorData.getInstance().FK_NAD_SU +
                            EdiConnectorData.getInstance().FK_NAD_UC +
                            EdiConnectorData.getInstance().FK_NAD_PE +
                            EdiConnectorData.getInstance().FK_OBNUMMER +
                            EdiConnectorData.getInstance().FK_ACT +
                            EdiConnectorData.getInstance().FK_CUX +
                            EdiConnectorData.getInstance().FK_DAGEN +
                            EdiConnectorData.getInstance().FK_KORTPERC +
                            EdiConnectorData.getInstance().FK_KORTBEDR +
                            EdiConnectorData.getInstance().FK_ONTVANGER;

                        writer.WriteLine("0" + header);

                        EdiConnectorData.getInstance().AK_SOORT = "";
                        EdiConnectorData.getInstance().AK_QUAL = "";
                        EdiConnectorData.getInstance().AK_BEDRAG = "";
                        EdiConnectorData.getInstance().AK_BTWSOORT = "";
                        EdiConnectorData.getInstance().AK_FOOTMOA = "";
                        EdiConnectorData.getInstance().AK_NOTINCALC = "";

                        if (oInvoice.DiscountPercent > 0)
                            EdiConnectorData.getInstance().AK_SOORT = FormatField("C", 1, "STRING");
                        else
                            EdiConnectorData.getInstance().AK_SOORT = FormatField("A", 1, "STRING");

                        EdiConnectorData.getInstance().AK_QUAL = FormatField(oInvoice.UserFields.Fields.Item("U_QUAL").Value.ToString(), 3, "STRING");
                        EdiConnectorData.getInstance().AK_BEDRAG = FormatField(oInvoice.TotalDiscount.ToString(), 9, "STRING");

                        EdiConnectorData.getInstance().AK_BTWSOORT = FormatField(oInvoice.UserFields.Fields.Item("U_BTWSOORT").Value.ToString(), 1, "STRING");
                        EdiConnectorData.getInstance().AK_FOOTMOA = FormatField(oInvoice.UserFields.Fields.Item("U_FOOTMOA").Value.ToString(), 1, "STRING");
                        EdiConnectorData.getInstance().AK_NOTINCALC = FormatField(oInvoice.UserFields.Fields.Item("U_NOTINCALC").Value.ToString(), 1, "STRING");

                        alcHeader = EdiConnectorData.getInstance().AK_SOORT +
                            EdiConnectorData.getInstance().AK_QUAL +
                            EdiConnectorData.getInstance().AK_BEDRAG +
                            EdiConnectorData.getInstance().AK_BTWSOORT +
                            EdiConnectorData.getInstance().AK_FOOTMOA +
                            EdiConnectorData.getInstance().AK_NOTINCALC;

                        if (EdiConnectorData.getInstance().AK_SOORT == "C")
                            writer.WriteLine("1" + alcHeader);

                        for (int i = 0; i < oInvoice.Lines.Count - 1; i++)
                        {
                            oInvoice.Lines.SetCurrentLine(i);

                            line = "";

                            EdiConnectorData.getInstance().FR_DEUAC = "";
                            EdiConnectorData.getInstance().FR_DEARTNR = "";
                            EdiConnectorData.getInstance().FR_DEARTOM = "";
                            EdiConnectorData.getInstance().FR_AANTAL = "";
                            EdiConnectorData.getInstance().FR_FAANTAL = "";
                            EdiConnectorData.getInstance().FR_ARTEENHEID = "";
                            EdiConnectorData.getInstance().FR_FEENHEID = "";
                            EdiConnectorData.getInstance().FR_NETTOBEDR = "";
                            EdiConnectorData.getInstance().FR_PRIJS = "";
                            EdiConnectorData.getInstance().FR_FREKEN = "";
                            EdiConnectorData.getInstance().FR_BTWSOORT = "";
                            EdiConnectorData.getInstance().FR_PV = "";
                            EdiConnectorData.getInstance().FR_ORDER = "";
                            EdiConnectorData.getInstance().FR_REGELID = "";
                            EdiConnectorData.getInstance().FR_INVO = "";
                            EdiConnectorData.getInstance().FR_DESA = "";
                            EdiConnectorData.getInstance().FR_PRIAAA = "";
                            EdiConnectorData.getInstance().FR_PIAPB = "";

                            EdiConnectorData.getInstance().FR_DEUAC = FormatField(oInvoice.UserFields.Fields.Item("U_DEUAC").Value.ToString(), 14, "STRING");
                            EdiConnectorData.getInstance().FR_DEARTNR = FormatField("---------", 9, "STRING");
                            EdiConnectorData.getInstance().FR_DEARTOM = FormatField(oInvoice.Lines.ItemDescription, 70, "STRING");
                            EdiConnectorData.getInstance().FR_AANTAL = FormatField(oInvoice.Lines.Quantity.ToString(), 5, "STRING");
                            EdiConnectorData.getInstance().FR_FAANTAL = FormatField(oInvoice.Lines.Quantity.ToString(), 9, "STRING");
                            EdiConnectorData.getInstance().FR_ARTEENHEID = FormatField(oInvoice.UserFields.Fields.Item("U_ARTEENHEID").Value.ToString(), 3, "STRING");
                            EdiConnectorData.getInstance().FR_FEENHEID = FormatField(oInvoice.UserFields.Fields.Item("U_FEENHEID").Value.ToString(), 3, "STRING");
                            EdiConnectorData.getInstance().FR_NETTOBEDR = FormatField(oInvoice.Lines.LineTotal.ToString(), 11, "STRING");
                            EdiConnectorData.getInstance().FR_PRIJS = FormatField(oInvoice.Lines.Price.ToString(), 10, "STRING");
                            EdiConnectorData.getInstance().FR_FREKEN = FormatField(oInvoice.UserFields.Fields.Item("U_FREKEN").Value.ToString(), 9, "STRING");

                            switch (oInvoice.Lines.VatGroup.Trim())
                            {
                                case "A0":
                                    EdiConnectorData.getInstance().FR_BTWSOORT = FormatField("0", 1, "STRING");
                                    break;
                                case "A1":
                                    EdiConnectorData.getInstance().FR_BTWSOORT = FormatField("L", 1, "STRING");
                                    break;
                                case "A2":
                                    EdiConnectorData.getInstance().FR_BTWSOORT = FormatField("H", 1, "STRING");
                                    break;
                                default:
                                    EdiConnectorData.getInstance().FR_BTWSOORT = FormatField("9", 1, "STRING");
                                    break;
                            }

                            EdiConnectorData.getInstance().FR_PV = FormatField(oInvoice.UserFields.Fields.Item("U_PV").Value.ToString(), 10, "STRING");
                            EdiConnectorData.getInstance().FR_ORDER = FormatField(oInvoice.UserFields.Fields.Item("U_ORDER").Value.ToString(), 35, "STRING");
                            EdiConnectorData.getInstance().FR_REGELID = FormatField(oInvoice.Lines.LineNum.ToString(), 6, "STRING");
                            EdiConnectorData.getInstance().FR_INVO = FormatField(oInvoice.UserFields.Fields.Item("U_INVO").Value.ToString(), 35, "STRING");
                            EdiConnectorData.getInstance().FR_DESA = FormatField(oInvoice.UserFields.Fields.Item("U_DESA").Value.ToString(), 35, "STRING");
                            EdiConnectorData.getInstance().FR_PRIAAA = FormatField(oInvoice.UserFields.Fields.Item("U_PRIAAA").Value.ToString(), 10, "STRING");
                            EdiConnectorData.getInstance().FR_PIAPB = FormatField(oInvoice.Lines.SupplierCatNum, 9, "STRING");

                            line = EdiConnectorData.getInstance().FR_DEUAC +
                            EdiConnectorData.getInstance().FR_DEARTNR +
                            EdiConnectorData.getInstance().FR_DEARTOM +
                            EdiConnectorData.getInstance().FR_AANTAL +
                            EdiConnectorData.getInstance().FR_FAANTAL +
                            EdiConnectorData.getInstance().FR_ARTEENHEID +
                            EdiConnectorData.getInstance().FR_FEENHEID +
                            EdiConnectorData.getInstance().FR_NETTOBEDR +
                            EdiConnectorData.getInstance().FR_PRIJS +
                            EdiConnectorData.getInstance().FR_FREKEN +
                            EdiConnectorData.getInstance().FR_BTWSOORT +
                            EdiConnectorData.getInstance().FR_PV +
                            EdiConnectorData.getInstance().FR_ORDER +
                            EdiConnectorData.getInstance().FR_REGELID +
                            EdiConnectorData.getInstance().FR_INVO +
                            EdiConnectorData.getInstance().FR_DESA +
                            EdiConnectorData.getInstance().FR_PRIAAA +
                            EdiConnectorData.getInstance().FR_PIAPB;

                            if (oInvoice.Lines.TreeType != BoItemTreeTypes.iIngredient)
                                writer.WriteLine("2" + line);
                        }// End for loop

                        writer.Close();

                        oInvoice.UserFields.Fields.Item("U_EDI_INV_EXP").Value = "Ja";
                        oInvoice.UserFields.Fields.Item("U_EDI_INVEXP_TIJD").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

                        oInvoice.Update();

                        if (EdiConnectorData.getInstance().iSendNotification == 1)
                            MailToDLreceiver(oInvoice.DocEntry, oInvoice.DocDate.ToString(), oInvoice.DocNum);

                        Log("V", "Invoice file created!", "CreateInvoiceFile");

                    }// End using
                }
                catch (Exception e)
                {
                    File.Delete(EdiConnectorData.getInstance().sInvoicePath + @"\" + "INV_" + oInvoice.DocNum + ".DAT");
                    Log("X", e.Message, "CreateInvoiceFile");
                }
            }// Key not found
            else
            {
                Log("X", "Invoice " + _docEntry + " not found!", "CreateInvoiceFile");
            }
        }

        private string FormatField(string _fieldValue, int _fieldLen, string _type)
        {
            string newValue;

            switch (_type)
            {
                case "STRING":
                    if (_fieldValue.Length > 0)
                        newValue = _fieldValue.Trim() + new String(' ', _fieldLen).Substring(1, _fieldLen - _fieldValue.PadLeft(_fieldLen).Trim().Length);
                    else
                        newValue = new String(' ', _fieldLen);
                    break;
                default:
                    newValue = new String(' ', _fieldLen);
                    break;
            }

            return newValue;
        }

        public bool MailToDLreceiver(int _docEntry, string _docDate, int _docNum)
        {
            try
            {
                using (MailMessage mailMsg = new MailMessage())
                {
                    SmtpClient smtpMail = new SmtpClient(EdiConnectorData.getInstance().sSmpt, EdiConnectorData.getInstance().iSmtpPort);

                    if (EdiConnectorData.getInstance().bSmtpUserSecurity == true)
                        smtpMail.Credentials = new NetworkCredential(EdiConnectorData.getInstance().sSmtpUser, EdiConnectorData.getInstance().sSmtpPassword);

                    mailMsg.From = new MailAddress(EdiConnectorData.getInstance().sSenderEmail, EdiConnectorData.getInstance().sSenderName);
                    mailMsg.To.Add(EdiConnectorData.getInstance().sDeliveryMailTo);
                    mailMsg.Subject = String.Format("Delivery {0} exported", _docEntry);

                    using (StreamReader fileReader = new StreamReader(EdiConnectorData.getInstance().sApplicationPath + @"\email_d.txt"))
                    {
                        string body;
                        body = fileReader.ReadToEnd();
                        body = body.Replace("::NAME::", EdiConnectorData.getInstance().sDeliveryMailToFullName);
                        body = body.Replace("::DOCENTRY::", _docEntry.ToString());
                        body = body.Replace("::DOCDATE::", _docDate);
                        body = body.Replace("::DOCNUM::", _docNum.ToString());
                        fileReader.Close();

                        mailMsg.Body = body;
                    }
                    smtpMail.Send(mailMsg);
                }

                Log("V", "Delivery notification sent!", "MailToDLreceiver");
                return true;
            }
            catch
            {
                Log("X", "Delivery notification was NOT sent!", "MailToDLreceiver");
                return false;
            }
        }

        public bool MailToINreceiver(int _docEntry, string _docDate, int _docNum)
        {
            try
            {
                using (MailMessage mailMsg = new MailMessage())
                {
                    SmtpClient smtpMail = new SmtpClient(EdiConnectorData.getInstance().sSmpt, EdiConnectorData.getInstance().iSmtpPort);

                    if (EdiConnectorData.getInstance().bSmtpUserSecurity == true)
                        smtpMail.Credentials = new NetworkCredential(EdiConnectorData.getInstance().sSmtpUser, EdiConnectorData.getInstance().sSmtpPassword);

                    mailMsg.From = new MailAddress(EdiConnectorData.getInstance().sSenderEmail, EdiConnectorData.getInstance().sSenderName);
                    mailMsg.To.Add(EdiConnectorData.getInstance().sInvoiceMailTo);
                    mailMsg.Subject = String.Format("Invoice {0} imported", _docEntry);

                    using (StreamReader fileReader = new StreamReader(EdiConnectorData.getInstance().sApplicationPath + @"\email_i.txt"))
                    {
                        string body;
                        body = fileReader.ReadToEnd();
                        body = body.Replace("::NAME::", EdiConnectorData.getInstance().sInvoiceMailToFullName);
                        body = body.Replace("::DOCENTRY::", _docEntry.ToString());
                        body = body.Replace("::DOCDATE::", _docDate);
                        body = body.Replace("::DOCNUM::", _docNum.ToString());
                        fileReader.Close();

                        mailMsg.Body = body;
                    }
                    smtpMail.Send(mailMsg);
                }

                Log("V", "Invoice notification sent!", "MailToINreceiver");
                return true;
            }
            catch
            {
                Log("X", "Invoice notification was NOT sent!", "MailToINreceiver");
                return false;
            }
        }
        #endregion

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
