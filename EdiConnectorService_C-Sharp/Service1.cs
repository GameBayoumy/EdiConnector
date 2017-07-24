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
using System.Net;
using SAPbobsCOM;

namespace EdiConnectorService_C_Sharp
{
    public partial class EdiConnectorService : ServiceBase
    {
        private EdiConnectorData ECD;
        private bool stopping;
        private ManualResetEvent stoppedEvent;

        public EdiConnectorService()
        {
            ECD = new EdiConnectorData();
            InitializeComponent();

            this.stopping = false;
            this.stoppedEvent = new ManualResetEvent(false);
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
            this.eventLog1.WriteEntry("EdiService in OnStart.");

            ReadSettings();
            ConnectToSAP();
            CreateUdfFiledsText();

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
                CheckAndExportDelivery();
                CheckAndExportInvoice();
                SplitOrder();
                ReadSOfile();

                if (ECD.cmp.Connected == true)
                {

                }
                else
                {
                    ConnectToSAP();
                }

                Thread.Sleep(4000);

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

        public bool ConnectToSAP()
        {
            try
            {
                if (ECD.cmp.Connected == true)
                {
                    Log("X", "SAP is already Connected", "ConnectToSAP");
                    return false;
                }

                bool bln = ConnectToDataBase();

                if (bln == false)
                {
                    return false;
                }

                ECD.cmp.DbServerType = ECD.bstDBServerType;
                ECD.cmp.Server = ECD.sServer;
                ECD.cmp.DbUserName = ECD.sDBUserName;
                ECD.cmp.DbPassword = ECD.sDBPassword;
                ECD.cmp.CompanyDB = ECD.sCompanyDB;
                ECD.cmp.UserName = ECD.sUserName;
                ECD.cmp.Password = ECD.sPassword;
                ECD.cmp.UseTrusted = false;

                if (ECD.cmp.Connect() != 0)
                {
                    int errCode;
                    string errMsg = "";
                    ECD.cmp.GetLastError(out errCode, out errMsg);
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
                ECD.cn.ConnectionString = "Data Source=" + ECD.sServer +
                    ";Initial Catalog=" + ECD.sCompanyDB +
                    ";User ID=" + ECD.sDBUserName +
                    ";Password=" + ECD.sDBPassword;

                if (ECD.sSQLVersion == "2005")
                    ECD.bstDBServerType = BoDataServerTypes.dst_MSSQL2005;
                else if (ECD.sSQLVersion == "2008")
                    ECD.bstDBServerType = BoDataServerTypes.dst_MSSQL2008;
                else if (ECD.sSQLVersion == "2012")
                    ECD.bstDBServerType = BoDataServerTypes.dst_MSSQL2012;
                else
                    ECD.bstDBServerType = BoDataServerTypes.dst_MSSQL;

                ECD.cn.Open();

                Log("V", "Connected to database " + ECD.sCompanyDB + ".", "ConnectToDatabase");
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
                dataSet.ReadXml(ECD.sApplicationPath + @"\settings.xml");

                if (dataSet.Tables["server"].Rows[0]["sqlversion"] == "2005")
                    ECD.bstDBServerType = BoDataServerTypes.dst_MSSQL2005;
                else if (dataSet.Tables["server"].Rows[0]["sqlversion"] == "2008")
                    ECD.bstDBServerType = BoDataServerTypes.dst_MSSQL2008;
                else if (dataSet.Tables["server"].Rows[0]["sqlversion"] == "2012")
                    ECD.bstDBServerType = BoDataServerTypes.dst_MSSQL2012;
                else
                    ECD.bstDBServerType = BoDataServerTypes.dst_MSSQL;

                ECD.sServer = dataSet.Tables["server"].Rows[0]["name"].ToString();
                ECD.sDBUserName = dataSet.Tables["server"].Rows[0]["userid"].ToString();
                ECD.sDBPassword = dataSet.Tables["server"].Rows[0]["password"].ToString();
                ECD.sCompanyDB = dataSet.Tables["server"].Rows[0]["catalog"].ToString();
                ECD.sUserName = dataSet.Tables["server"].Rows[0]["sapuser"].ToString();
                ECD.sPassword = dataSet.Tables["server"].Rows[0]["sappassword"].ToString();
                ECD.sSQLVersion = dataSet.Tables["server"].Rows[0]["sqlversion"].ToString();
                ECD.sDesAdvLevel = dataSet.Tables["server"].Rows[0]["desadvlevel"].ToString();

                ECD.sSOPath = dataSet.Tables["folders"].Rows[0]["so"].ToString();
                ECD.sSOTempPath = dataSet.Tables["folders"].Rows[0]["sotemp"].ToString();
                ECD.sSODonePath = dataSet.Tables["folders"].Rows[0]["sodone"].ToString();
                ECD.sSOErrorPath = dataSet.Tables["folders"].Rows[0]["soerror"].ToString();
                ECD.sInvoicePath = dataSet.Tables["folders"].Rows[0]["invoice"].ToString();
                ECD.sDeliveryPath = dataSet.Tables["folders"].Rows[0]["delivery"].ToString();

                ECD.iSendNotification = Convert.ToInt32(dataSet.Tables["email"].Rows[0]["send_notification"]);

                ECD.sSmpt = dataSet.Tables["email"].Rows[0]["smtp"].ToString();
                ECD.iSmtpPort = Convert.ToInt32(dataSet.Tables["email"].Rows[0]["port"]);

                if (dataSet.Tables["email"].Rows[0]["security"].ToString() == "0")
                    ECD.bSmtpUserSecurity = false;
                else
                    ECD.bSmtpUserSecurity = true;

                ECD.sSmtpUser = dataSet.Tables["email"].Rows[0]["user"].ToString();
                ECD.sSmtpPassword = dataSet.Tables["email"].Rows[0]["password"].ToString();

                ECD.sSenderEmail = dataSet.Tables["email"].Rows[0]["emailaddress"].ToString();
                ECD.sSenderName = dataSet.Tables["email"].Rows[0]["fullname"].ToString();

                ECD.sOrderMailTo = dataSet.Tables["email"].Rows[0]["emailaddress_order"].ToString();
                ECD.sOrderMailToFullName = dataSet.Tables["email"].Rows[0]["fullname_order"].ToString();

                ECD.sDeliveryMailTo = dataSet.Tables["email"].Rows[0]["emailaddress_delivery"].ToString();
                ECD.sDeliveryMailToFullName = dataSet.Tables["email"].Rows[0]["fullname_delivery"].ToString();

                ECD.sInvoiceMailTo = dataSet.Tables["email"].Rows[0]["emailaddress_invoice"].ToString();
                ECD.sInvoiceMailToFullName = dataSet.Tables["email"].Rows[0]["fullname_invoice"].ToString();

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
            dataSet.ReadXml(ECD.sApplicationPath + @"\orderfld.xml");

            for (int orderHead = 0; orderHead < dataSet.Tables["OK"].Rows.Count; orderHead++ )
            {
                ECD.OK_POS[orderHead] = Convert.ToInt32(dataSet.Tables["OK"].Rows[orderHead][0]);
                ECD.OK_LEN[orderHead] = Convert.ToInt32(dataSet.Tables["OK"].Rows[orderHead][1]);
            }

            for (int orderItem = 0; orderItem < dataSet.Tables["OR"].Rows.Count; orderItem++)
            {
                ECD.OK_POS[orderItem] = Convert.ToInt32(dataSet.Tables["OR"].Rows[orderItem][0]);
                ECD.OK_LEN[orderItem] = Convert.ToInt32(dataSet.Tables["OR"].Rows[orderItem][1]);
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
            if (ECD.cmp.Connected == true)
            {
                ECD.cmp.Disconnect();
                ECD.cn.Close();
                Log("V", "SAP is disconneted.", "DisconnectToSAP");
            }
            else
                Log("X", "SAP is already disconneted.", "DisconnectToSAP");
        }

        public void SplitOrder()
        {
            string fileSize = "";
            DirectoryInfo dirI = new DirectoryInfo(ECD.sSOPath);
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
                using (StreamWriter writer = new StreamWriter(ECD.sSOTempPath + @"\" + "ORDER" + "_" + BGMnumber + ".DAT"))
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

        }
    }
}
