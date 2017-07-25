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
            CreateUdfFieldsText();

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

                bool bln = ConnectToDatabase();

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
            DirectoryInfo dirI = new DirectoryInfo(ECD.sSOTempPath);
            FileInfo[] arrayFileI = dirI.GetFiles("*.DAT");

            foreach (FileInfo fileInfo in arrayFileI)
            {
                try
                {
                    Log("V", "Begin reading sales order filename " + fileInfo.Name + ".", "Read_SO_file");

                    StreamReader reader = new StreamReader(fileInfo.FullName, false);
                    ECD.SO_FILE = reader.ReadToEnd();
                    ECD.SO_FILENAME = fileInfo.Name;
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

                using (StringReader reader = new StringReader(ECD.SO_FILE))
                {
                    while (true)
                    {
                        line = reader.ReadLine();
                        switch (line.Substring(1, 1))
                        {
                            case "0":
                                ECD.OK_K_EANCODE = line.Substring(ECD.OK_POS[0], ECD.OK_LEN[0]).Trim();
                                ECD.OK_TEST = line.Substring(ECD.OK_POS[1], ECD.OK_LEN[1]).Trim();
                                ECD.OK_KNAAM = line.Substring(ECD.OK_POS[2], ECD.OK_LEN[2]).Trim();
                                ECD.OK_BGM = line.Substring(ECD.OK_POS[3], ECD.OK_LEN[3]).Trim();
                                if (line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Trim().Length == 8)
                                    ECD.OK_K_ORDDAT = Convert.ToDateTime(line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(7, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(1, 4));
                                else
                                    ECD.OK_K_ORDDAT = Convert.ToDateTime("01-01-0001");

                                if (line.Substring(ECD.OK_POS[5], ECD.OK_LEN[5]).Trim().Length == 8)
                                    ECD.OK_DTM_2 = Convert.ToDateTime(line.Substring(ECD.OK_POS[5], ECD.OK_LEN[5]).Substring(7, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(1, 4));
                                else
                                    ECD.OK_DTM_2 = Convert.ToDateTime("01-01-0001");
                                ECD.OK_TIJD_2 = line.Substring(ECD.OK_POS[6], ECD.OK_LEN[6]).Trim();

                                if (line.Substring(ECD.OK_POS[7], ECD.OK_LEN[7]).Trim().Length == 8)
                                    ECD.OK_DTM_17 = Convert.ToDateTime(line.Substring(ECD.OK_POS[7], ECD.OK_LEN[7]).Substring(7, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(1, 4));
                                else
                                    ECD.OK_DTM_17 = Convert.ToDateTime("01-01-0001");
                                ECD.OK_TIJD_17 = line.Substring(ECD.OK_POS[8], ECD.OK_LEN[8]).Trim();

                                if (line.Substring(ECD.OK_POS[9], ECD.OK_LEN[9]).Trim().Length == 8)
                                    ECD.OK_DTM_64 = Convert.ToDateTime(line.Substring(ECD.OK_POS[9], ECD.OK_LEN[9]).Substring(7, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(1, 4));
                                else
                                    ECD.OK_DTM_64 = Convert.ToDateTime("01-01-0001");
                                ECD.OK_TIJD_64 = line.Substring(ECD.OK_POS[10], ECD.OK_LEN[10]).Trim();

                                if (line.Substring(ECD.OK_POS[11], ECD.OK_LEN[11]).Trim().Length == 8)
                                    ECD.OK_DTM_63 = Convert.ToDateTime(line.Substring(ECD.OK_POS[11], ECD.OK_LEN[11]).Substring(7, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(1, 4));
                                else
                                    ECD.OK_DTM_63 = Convert.ToDateTime("01-01-0001");
                                ECD.OK_TIJD_63 = line.Substring(ECD.OK_POS[12], ECD.OK_LEN[12]).Trim();

                                ECD.OK_RFF_BO = line.Substring(ECD.OK_POS[13], ECD.OK_LEN[13]).Trim();
                                ECD.OK_RFF_CR = line.Substring(ECD.OK_POS[14], ECD.OK_LEN[14]).Trim();
                                ECD.OK_RFF_PD = line.Substring(ECD.OK_POS[15], ECD.OK_LEN[15]).Trim();
                                ECD.OK_RFFCT = line.Substring(ECD.OK_POS[16], ECD.OK_LEN[16]).Trim();
                                if (line.Substring(ECD.OK_POS[17], ECD.OK_LEN[17]).Trim().Length == 8)
                                    ECD.OK_DTMCT = Convert.ToDateTime(line.Substring(ECD.OK_POS[17], ECD.OK_LEN[17]).Substring(7, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(5, 2) + "-" + line.Substring(ECD.OK_POS[4], ECD.OK_LEN[4]).Substring(1, 4));
                                else
                                    ECD.OK_DTMCT = Convert.ToDateTime("01-01-0001");

                                ECD.OK_FLAGS[0] = line.Substring(ECD.OK_POS[18], ECD.OK_LEN[18]).Trim();
                                ECD.OK_FLAGS[1] = line.Substring(ECD.OK_POS[19], ECD.OK_LEN[19]).Trim();
                                ECD.OK_FLAGS[2] = line.Substring(ECD.OK_POS[20], ECD.OK_LEN[20]).Trim();
                                ECD.OK_FLAGS[3] = line.Substring(ECD.OK_POS[21], ECD.OK_LEN[21]).Trim();
                                ECD.OK_FLAGS[4] = line.Substring(ECD.OK_POS[22], ECD.OK_LEN[22]).Trim();
                                ECD.OK_FLAGS[5] = line.Substring(ECD.OK_POS[23], ECD.OK_LEN[23]).Trim();
                                ECD.OK_FLAGS[6] = line.Substring(ECD.OK_POS[24], ECD.OK_LEN[24]).Trim();
                                ECD.OK_FLAGS[7] = line.Substring(ECD.OK_POS[25], ECD.OK_LEN[25]).Trim();
                                ECD.OK_FLAGS[8] = line.Substring(ECD.OK_POS[26], ECD.OK_LEN[26]).Trim();
                                ECD.OK_FLAGS[9] = line.Substring(ECD.OK_POS[27], ECD.OK_LEN[27]).Trim();
                                ECD.OK_FLAGS[10] = line.Substring(ECD.OK_POS[28], ECD.OK_LEN[28]).Trim();
                                ECD.OK_FTXDSI = line.Substring(ECD.OK_POS[29], ECD.OK_LEN[29]).Trim();
                                ECD.OK_NAD_BY = line.Substring(ECD.OK_POS[30], ECD.OK_LEN[30]).Trim();
                                ECD.OK_NAD_DP = line.Substring(ECD.OK_POS[31], ECD.OK_LEN[31]).Trim();
                                ECD.OK_NAD_IV = line.Substring(ECD.OK_POS[32], ECD.OK_LEN[32]).Trim();
                                ECD.OK_NAD_SF = line.Substring(ECD.OK_POS[33], ECD.OK_LEN[33]).Trim();
                                ECD.OK_NAD_SU = line.Substring(ECD.OK_POS[34], ECD.OK_LEN[34]).Trim();
                                ECD.OK_NAD_UC = line.Substring(ECD.OK_POS[35], ECD.OK_LEN[35]).Trim();
                                ECD.OK_NAD_BCO = line.Substring(ECD.OK_POS[36], ECD.OK_LEN[36]).Trim();
                                ECD.OK_RECEIVER = line.Substring(ECD.OK_POS[37], ECD.OK_LEN[37]).Trim();
                                bool matchHead = WriteSOhead();
                                if (matchHead == false)
                                {
                                    // Move file
                                    File.Move(ECD.sSOTempPath + @"\" + ECD.SO_FILENAME, ECD.sSOErrorPath + @"\" + DateTime.Now.ToString("HHmmss") + "_" + ECD.SO_FILENAME);
                                    Log("X", ECD.SO_FILENAME + " copied to the errors folder!", "MatchSOdata");
                                    return false;
                                }
                                break;

                            case "1":
                                ECD.OR_DEUAC = line.Substring(ECD.OR_POS[0], ECD.OR_LEN[0]).Trim();
                                ECD.OR_QTY = Convert.ToDouble(line.Substring(ECD.OR_POS[1], ECD.OR_LEN[1]).Trim());
                                ECD.OR_LEVARTCODE = line.Substring(ECD.OR_POS[2], ECD.OR_LEN[2]).Trim();
                                ECD.OR_DEARTOM = line.Substring(ECD.OR_POS[3], ECD.OR_LEN[3]).Trim();
                                ECD.OR_COLOR = line.Substring(ECD.OR_POS[4], ECD.OR_LEN[4]).Trim();
                                ECD.OR_LENGTH = line.Substring(ECD.OR_POS[5], ECD.OR_LEN[5]).Trim();
                                ECD.OR_WIDTH = line.Substring(ECD.OR_POS[6], ECD.OR_LEN[6]).Trim();
                                ECD.OR_HEIGHT = line.Substring(ECD.OR_POS[7], ECD.OR_LEN[7]).Trim();
                                ECD.OR_CUX = line.Substring(ECD.OR_POS[8], ECD.OR_LEN[8]).Trim();
                                ECD.OR_PIA = line.Substring(ECD.OR_POS[9], ECD.OR_LEN[9]).Trim();
                                ECD.OR_RFFLI1 = line.Substring(ECD.OR_POS[10], ECD.OR_LEN[10]).Trim();
                                ECD.OR_RFFLI2 = line.Substring(ECD.OR_POS[11], ECD.OR_LEN[11]).Trim();
                                if (line.Substring(ECD.OR_POS[12], ECD.OR_LEN[12]).Trim().Length == 8)
                                    ECD.OR_DTM_2 = Convert.ToDateTime(line.Substring(ECD.OR_POS[12], ECD.OR_LEN[12]).Substring(7, 2) + "-" + line.Substring(ECD.OR_POS[12], ECD.OR_LEN[12]).Substring(5, 2) + "-" + line.Substring(ECD.OR_POS[12], ECD.OR_LEN[12]).Substring(1, 4));
                                else
                                    ECD.OR_DTM_2 = Convert.ToDateTime("01-01-0001");
                                ECD.OR_LINNR = line.Substring(ECD.OR_POS[13], ECD.OR_LEN[13]).Trim();
                                ECD.OR_PRI = Convert.ToDouble(line.Substring(ECD.OR_POS[14], ECD.OR_LEN[14]).Trim());
                                bool matchItem = WriteSOitems();
                                if (matchItem == false)
                                {
                                    File.Move(ECD.sSOTempPath + @"\" + ECD.SO_FILENAME, ECD.sSOErrorPath + @"\" + DateTime.Now.ToString("HHmmss") + "_" + ECD.SO_FILENAME);
                                    Log("X", ECD.SO_FILENAME + " copied to the errors folder!", "MatchSOdata");
                                    return false;
                                }
                                break;
                        } // End switch
                    } // End while
                } // End using 
                OrderSave();

                ECD.SO_FILE = "";
                ECD.SO_FILENAME = "";
                return true;
            } // End try
            catch(Exception e)
            {
                File.Move(ECD.sSOTempPath + @"\" + ECD.SO_FILENAME, ECD.sSOErrorPath + @"\" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "_" + ECD.SO_FILENAME);
                Log("X", e.Message, "MatchSOdata");
                Log("X", ECD.SO_FILENAME + " copied to the errors folder!", "MatchSOdata");
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

                ECD.oOrder = ECD.cmp.GetBusinessObject(BoObjectTypes.oOrders);

                ECD.oOrder.CardCode = ECD.CARDCODE;
                ECD.oOrder.NumAtCard = ECD.OK_BGM;
                ECD.oOrder.DocDate =  ECD.OK_K_ORDDAT;
                ECD.oOrder.TaxDate = ECD.OK_K_ORDDAT;

                // Reference Rosis DocDueDate
                if (ECD.OK_DTM_2.Year > 1999)
                    ECD.oOrder.DocDueDate = ECD.OK_DTM_2;
                else if(ECD.OK_DTM_17.Year > 1999)
                    ECD.oOrder.DocDueDate = ECD.OK_DTM_17;
                else if (ECD.OK_DTM_64.Year > 1999)
                    ECD.oOrder.DocDueDate = ECD.OK_DTM_64;
                else
                    ECD.oOrder.DocDueDate = Convert.ToDateTime("31/12/" + DateTime.Now.Year);

                ECD.oOrder.ShipToCode = ECD.OK_NAD_DP;
                ECD.oOrder.PayToCode = ECD.OK_NAD_IV;

                // Loop UDF
                foreach (Field field in ECD.oOrder.UserFields.Fields)
                {
                    switch (field.Name)
                    {
                        case "U_BGM": field.Value = ECD.OK_BGM; break;
                        case "U_RFF": field.Value = ECD.OK_BGM; break;

                        case "U_K_EANCODE": field.Value = ECD.OK_K_EANCODE; break;
                        case "U_KNAAM": field.Value = ECD.OK_KNAAM; break;
                        case "U_TEST":
                            if (ECD.OK_TEST == "1")
                                field.Value = "J";
                            else
                                field.Value = "N";
                            break;

                        case "U_DTM_2": if (ECD.OK_DTM_2.ToString() != "1-1-0001 0:00:00") field.Value = ECD.OK_DTM_2.ToString().Replace(" 0:00:00", ""); break;
                        case "U_TIJD_2": field.Value = ECD.OK_TIJD_2; break;

                        case "U_DTM_17": if (ECD.OK_DTM_17.ToString() != "1-1-0001 0:00:00") field.Value = ECD.OK_DTM_17.ToString().Replace(" 0:00:00", ""); break;
                        case "U_TIJD_17": field.Value = ECD.OK_TIJD_17; break;

                        case "U_DTM_64": if (ECD.OK_DTM_64.ToString() != "1-1-0001 0:00:00") field.Value = ECD.OK_DTM_64.ToString().Replace(" 0:00:00", ""); break;
                        case "U_TIJD_64": field.Value = ECD.OK_TIJD_64; break;

                        case "U_RFF_BO": field.Value = ECD.OK_RFF_BO.ToString(); break;
                        case "U_RFF_CR": field.Value = ECD.OK_RFF_CR.ToString(); break;
                        case "U_RFF_PD": field.Value = ECD.OK_RFF_PD.ToString(); break;
                        case "U_RFFCT": field.Value = ECD.OK_RFFCT.ToString(); break;

                        case "U_DTMCT": if (ECD.OK_DTMCT.ToString() != "1-1-0001 0:00:00") field.Value = ECD.OK_DTMCT.ToString().Replace(" 0:00:00", ""); break;

                        case "U_FLAG0": field.Value = ECD.OK_FLAGS[0].ToString(); break;
                        case "U_FLAG1": field.Value = ECD.OK_FLAGS[1].ToString(); break;
                        case "U_FLAG2": field.Value = ECD.OK_FLAGS[2].ToString(); break;
                        case "U_FLAG3": field.Value = ECD.OK_FLAGS[3].ToString(); break;
                        case "U_FLAG4": field.Value = ECD.OK_FLAGS[4].ToString(); break;
                        case "U_FLAG5": field.Value = ECD.OK_FLAGS[5].ToString(); break;
                        case "U_FLAG6": field.Value = ECD.OK_FLAGS[6].ToString(); break;
                        case "U_FLAG7": field.Value = ECD.OK_FLAGS[7].ToString(); break;
                        case "U_FLAG8": field.Value = ECD.OK_FLAGS[8].ToString(); break;
                        case "U_FLAG9": field.Value = ECD.OK_FLAGS[9].ToString(); break;
                        case "U_FLAG10": field.Value = ECD.OK_FLAGS[10].ToString(); break;

                        case "U_NAD_BY": field.Value = ECD.OK_NAD_BY.ToString(); break;
                        case "U_NAD_DP": field.Value = ECD.OK_NAD_DP.ToString(); break;
                        case "U_NAD_IV": field.Value = ECD.OK_NAD_IV.ToString(); break;
                        case "U_NAD_SF": field.Value = ECD.OK_NAD_SF.ToString(); break;
                        case "U_NAD_SU": field.Value = ECD.OK_NAD_SU.ToString(); break;
                        case "U_NAD_UC": field.Value = ECD.OK_NAD_UC.ToString(); break;
                        case "U_NAD_BCO": field.Value = ECD.OK_NAD_BCO.ToString(); break;

                        case "U_ONTVANGER": field.Value = ECD.OK_RECEIVER.ToString(); break;
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
            Recordset oRecordSet;
            oRecordSet = ECD.cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRecordSet.DoQuery("SELECT LicTradNum FROM OCRD WHERE U_K_EANCODE = '" + buyerEANCode + "'");
            if (oRecordSet.RecordCount == 1)
                buyerAddress = oRecordSet.Fields.Item(0).Value.ToString();
            else if (oRecordSet.RecordCount > 1)
                Log("X", "Error: Duplicate EANcode Buyers found!", "CheckBuyerAddress");
            else
                Log("X", "Error: Match EANcode Buyers NOT found!", "CheckBuyerAddress");

            oRecordSet = null;
            return buyerAddress;
        }

        public bool CheckSOhead()
        {
            ECD.CARDCODE = "";
            Recordset oRecordSet;
            oRecordSet = ECD.cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet.DoQuery("SELECT CardCode FROM OCRD WHERE U_K_EANCODE = '" + ECD.OK_K_EANCODE + "'");
                if (oRecordSet.RecordCount == 1)
                {
                    ECD.CARDCODE = oRecordSet.Fields.Item(0).Value.ToString();
                    return true;
                }
                else if (oRecordSet.RecordCount > 1)
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
                oRecordSet = null;
            }
        }

        public bool WriteSOitems()
        {
            try
            {
                bool matchItems = CheckSOitems();
                if (matchItems == false)
                    return false;

                ECD.oOrder.Lines.ItemCode = ECD.ITEMCODE;
                string replace = ECD.OR_QTY.ToString().Replace(",", ".");
                ECD.oOrder.Lines.Quantity = Convert.ToDouble(String.Format(ECD.SALPACKUN, replace));
                ECD.oOrder.Lines.ShipDate = ECD.OK_DTM_2;

                foreach (Field field in ECD.oOrder.Lines.UserFields.Fields)
                {
                    switch (field.Name)
                    {
                        case "U_DEUAC": field.Value = ECD.OR_DEUAC.ToString(); break;
                        case "U_LEVARTCODE": field.Value = ECD.OR_LEVARTCODE.ToString(); break;
                        case "U_DEARTOM": field.Value = ECD.OR_DEARTOM.ToString(); break;
                        case "U_KLEUR": field.Value = ECD.OR_COLOR.ToString(); break;
                        case "U_LENGTE": field.Value = ECD.OR_LENGTH.ToString(); break;
                        case "U_BREEDTE": field.Value = ECD.OR_WIDTH.ToString(); break;
                        case "U_HOOGTE": field.Value = ECD.OR_HEIGHT.ToString(); break;
                        case "U_CUX": field.Value = ECD.OR_CUX.ToString(); break;
                        case "U_PIA": field.Value = ECD.OR_PIA.ToString(); break;
                        case "U_RFFLI1": field.Value = ECD.OR_RFFLI1.ToString(); break;
                        case "U_RFFLI2": field.Value = ECD.OR_RFFLI2.ToString(); break;
                        case "U_LINNR": field.Value = ECD.OR_LINNR.ToString(); break;
                        case "U_PRI": field.Value = ECD.OR_PRI.ToString(); break;

                    }
                }
                ECD.oOrder.Lines.Add();
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
            ECD.ITEMCODE = "";
            ECD.ITEMNAME = "";
            ECD.SALPACKUN = "1";
            Recordset oRecordSet;
            oRecordSet = ECD.cmp.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet.DoQuery("SELECT ItemCode, ItemName, SalPackUn FROM OITM WHERE U_EAN_Handels_EH = '" + ECD.OR_DEUAC + "'");

                if (oRecordSet.RecordCount == 1)
                {
                    ECD.ITEMCODE = oRecordSet.Fields.Item(0).Value.ToString();
                    ECD.ITEMNAME = oRecordSet.Fields.Item(1).Value.ToString();
                    ECD.SALPACKUN = oRecordSet.Fields.Item(2).Value.ToString();
                    return true;
                }
                else if (oRecordSet.RecordCount > 1)
                {
                    Log("X", "Error: Duplicate EANcodes found! Eancode " + ECD.OR_DEUAC, "CheckSOitems");
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
                oRecordSet = null;
            }
        }

        public void OrderSave()
        {
            int errCode;
            string errMsg = "";

            if (ECD.oOrder.Add() != 0)
            {
                ECD.cmp.GetLastError(out errCode, out errMsg);

                File.Move(ECD.sSOTempPath + @"\" + ECD.SO_FILENAME, ECD.sSOErrorPath + @"\" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "_" + ECD.SO_FILENAME);
                Log("X", "Error: " + errMsg + "(" + errCode + ")", "OrderSave");
                Log("X", ECD.SO_FILENAME + " copied to the errors folder!", "OrderSave");
            }
            else
            {
                if (ECD.iSendNotification == 1)
                    MailToSOreceiver(ECD.oOrder.DocEntry, ECD.oOrder.DocDate, ECD.oOrder.DocNum);

                Log("V", "Order written in SAP BO!", "OrderSave");
                File.Move(ECD.sSOTempPath + @"\" + ECD.SO_FILENAME, ECD.sSODonePath + @"\" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "_" + ECD.SO_FILENAME);
            }   
        }

        public void CreateUdfFieldsText()
        {
            if(File.Exists(ECD.sApplicationPath + @"\udf.xml") == false)
            {
                return;
            }

            DataSet dataSet = new DataSet();
            dataSet.ReadXml(ECD.sApplicationPath + @"\udf.xml");

            try
            {
                if (dataSet.Tables["udf"].Rows.Count > 0)
                {
                    for (int i = 0; i < dataSet.Tables["udf"].Rows.Count; i++)
                        CreateUdf(dataSet.Tables["udf"].Rows[i][0].ToString(), dataSet.Tables["udf"].Rows[i][1].ToString(), dataSet.Tables["udf"].Rows[i][2].ToString(), 
                            BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, Convert.ToInt32(dataSet.Tables["udf"].Rows[i][3]), false, false, "");
                }
            }
            catch (Exception e)
            {
                dataSet.Dispose();
            }
        }

        private void CreateUdf(string _tableName, string _fieldName, string _description,
            BoFieldTypes _boType, BoFldSubTypes _boSubType, int _editSize, bool _mandatory, bool _default, string _defaultValue)
        {
            IUserFieldsMD oUDF;
            oUDF = ECD.cmp.GetBusinessObject(BoObjectTypes.oUserFields);
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
                ECD.cmp.GetLastError(out errCode, out errMsg);
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
                    SmtpClient smtpMail = new SmtpClient(ECD.sSmpt, ECD.iSmtpPort);

                    if (ECD.bSmtpUserSecurity == true)
                        smtpMail.Credentials = new NetworkCredential(ECD.sSmtpUser, ECD.sSmtpPassword);

                    mailMsg.From = new MailAddress(ECD.sSenderEmail, ECD.sSenderName);
                    mailMsg.To.Add(ECD.sOrderMailTo);
                    mailMsg.Subject = String.Format("Order {0} imported", _docEntry);

                    using(StreamReader reader = new StreamReader(ECD.sApplicationPath + @"\email_o.txt"))
                    {
                        string body = reader.ReadToEnd();
                        body = body.Replace("::NAME::", ECD.sOrderMailToFullName);
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
            catch (Exception e)
            {
                Log("X", "Order notification was not sent!", "MailToSOreceiver");
                return false;
            }
        }
    }
}
