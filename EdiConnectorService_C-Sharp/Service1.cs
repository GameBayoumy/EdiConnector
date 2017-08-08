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

            // Initialize objects
            EdiConnectorData.getInstance();
            ConnectionManager.getInstance();
            agent = new Agent();

            EdiConnectorData.getInstance().sApplicationPath = @"H:\Projecten\Sharif\GitKraken\EdiConnector\EdiConnectorService_C-Sharp\";
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
            EdiConnectorData.getInstance().sApplicationPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + @"\";
            EdiConnectorData.getInstance().sProcessedDirName = "Processed";

            // Create connections from config.xml and try to connect all servers
            agent.QueueCommand(new CreateConnectionsCommand());
            ConnectionManager.getInstance().ConnectAll();

            // Creates udf fields for every connected server
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
            //Periodically check if the service is stopping.
            do
            {
                //Perform main service function here...

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
                        //foreach (FileInfo file in Files)
                        //{
                        //    agent.QueueCommand(new ProcessMessage(connectedServer, file.Name));
                        //}
                    }
                    else
                    {
                        EventLogger.getInstance().EventError("Messages file path does not exist!: " + messagesFilePath);
                    }
                }

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

        #region
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

        public void Log(string sType, string msg, string functionSender)
        {
            switch (sType)
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
