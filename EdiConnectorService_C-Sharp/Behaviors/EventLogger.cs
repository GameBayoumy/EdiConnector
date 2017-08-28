using System.Diagnostics;

namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// This class is used to display different kind of log messages using the System.Diagnostics.EventLog class.
    /// It supports Informative, Error and Warning entry types.
    /// 
    /// This class is also used to create and update log messages of incoming EDI messages.
    /// You can create and update messages through a connected server.
    /// </summary>
    class EventLogger
    {
        // Create a static instance of the event logger
        private static EventLogger instance = null;

        private EventLog eventLog1;

        // Initialize event log with constructor
	    private EventLogger(EventLog _eventLog)
	    {
            this.eventLog1 = _eventLog;
	    }

        public static void setInstance(EventLog _eventLog)
        {
            if (instance == null) // Create only 1 instance of the event logger
                instance = new EventLogger(_eventLog);
        }

        /// <summary>
        /// Gets the instance.
        /// </summary>
        /// <returns></returns>
        public static EventLogger getInstance()
	    {
		    return instance;
	    }

        public void EventInfo(string text)
        {
            eventLog1.WriteEntry(text, EventLogEntryType.Information);
        }

        public void EventError(string text)
        {
            eventLog1.WriteEntry(text, EventLogEntryType.Error);
        }

        public void EventWarning(string text)
        {
            eventLog1.WriteEntry(text, EventLogEntryType.Warning);
        }

        /// <summary>
        /// Creates the SAP log message.
        /// </summary>
        /// <param name="_connectedServer">The connected server.</param>
        /// <param name="_fileName">Name of the message file.</param>
        /// <param name="_xDoc">The loaded XML document.</param>
        /// <param name="_logMessage">The log message.</param>
        /// <param name="_status">The display status.</param>
        /// <returns>The record reference of the created field.</returns>
        public string CreateSAPLogMessage(string _connectedServer, string _fileName, System.Xml.Linq.XDocument _xDoc, string _logMessage, string _status)
        {
            SAPbobsCOM.UserTable oUDT = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.UserTables.Item("0_SWS_EDI_LOG");
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            System.DateTime createDateTime = System.DateTime.Now;
            string recordReference;

            try
            {
                oUDT.Name = _fileName;
                oUDT.UserFields.Fields.Item("U_XML_ATTACHMENT").Value = _xDoc.ToString();
                oUDT.UserFields.Fields.Item("U_STATUS").Value = _status;
                oUDT.UserFields.Fields.Item("U_LOG_MESSAGE").Value = createDateTime.ToString("dd-MM-yy HH:mm:ss : ") + oUDT.UserFields.Fields.Item("U_STATUS").Value + " - " + _logMessage;
                oUDT.UserFields.Fields.Item("U_CREATE_DATE").Value = createDateTime.Date;
                oUDT.UserFields.Fields.Item("U_CREATE_TIME").Value = createDateTime.ToShortTimeString();

                if (oUDT.Add() == 0)
                {
                    oRs.DoQuery(@"SELECT Max(""Code"") FROM ""@0_SWS_EDI_LOG""");
                    EventInfo("Server: " + _connectedServer + ". " + "Succesfully added incoming xml message (" + _fileName + ") to UDT: " + oUDT.TableName);
                    return recordReference = oRs.Fields.Item(0).Value.ToString();
                }
                else
                {
                    EventError("Server: " + _connectedServer + ". " + "Error adding items to UDT: " + ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription());
                    return recordReference = null;
                }
            }
            catch
            {
                ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                EventError("Server: " + _connectedServer + ". " + "Error adding items to UDF: " + errMsg);
                return recordReference = null;
            }
            finally
            {
                EdiConnectorService.ClearObject(oUDT);
                EdiConnectorService.ClearObject(oRs);
            }
        }

        /// <summary>
        /// Updates the SAP log message using the record reference.
        /// </summary>
        /// <param name="_connectedServer">The connected server.</param>
        /// <param name="_recordReference">The record reference.</param>
        /// <param name="_logMessage">The log message.</param>
        /// <param name="_status">The display status.</param>
        /// <param name="_docNumber">The document number.</param>
        public void UpdateSAPLogMessage(string _connectedServer, string _recordReference, string _logMessage, string _status = "", string _docNumber = "")
        {
            SAPbobsCOM.UserTable oUDT = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.UserTables.Item("0_SWS_EDI_LOG");

            oUDT.GetByKey(_recordReference);

            if(_status != "")
                oUDT.UserFields.Fields.Item("U_STATUS").Value = _status;
            if (_docNumber != "")
                oUDT.UserFields.Fields.Item("U_DOC_NR").Value = _docNumber;
            oUDT.UserFields.Fields.Item("U_LOG_MESSAGE").Value += System.Environment.NewLine + System.DateTime.Now.ToString("dd-MM-yy HH:mm:ss : ") + oUDT.UserFields.Fields.Item("U_STATUS").Value + " - " + _logMessage;

            if (oUDT.Update() != 0)
            {
                EventError($"Server: {_connectedServer}.  Error updating items to UDT: {ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription()}");
            }

            EdiConnectorService.ClearObject(oUDT);
        }
    }
}
