using System.Diagnostics;

namespace EdiConnectorService_C_Sharp
{
    class EventLogger
    {
        private static EventLogger instance = null;
        private EventLog eventLog1;

	    private EventLogger(EventLog _eventLog)
	    {
            this.eventLog1 = _eventLog;
	    }

        public static void setInstance(EventLog _eventLog)
        {
            if (instance == null)
                instance = new EventLogger(_eventLog);
        }

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
                    EventLogger.getInstance().EventInfo("Server: " + _connectedServer + " Succesfully added incoming xml message to UDT: " + oUDT.TableName);
                    return recordReference = oRs.Fields.Item(0).Value.ToString();
                }
                else
                {
                    EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error adding items to UDT: " + ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription());
                    return recordReference = null;
                }
            }
            catch
            {
                ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error adding items to UDF: " + errMsg);
                return recordReference = null;
            }
            finally
            {
                EdiConnectorService.ClearObject(oUDT);
                EdiConnectorService.ClearObject(oRs);
            }
        }

        public void UpdateSAPLogMessage(string _connectedServer, string _recordReference, string _logMessage, string _status = "")
        {
            SAPbobsCOM.UserTable oUDT = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.UserTables.Item("0_SWS_EDI_LOG");

            oUDT.GetByKey(_recordReference);

            if(_status != "")
                oUDT.UserFields.Fields.Item("U_STATUS").Value = _status;
            oUDT.UserFields.Fields.Item("U_LOG_MESSAGE").Value += System.Environment.NewLine + System.DateTime.Now.ToString("dd-MM-yy HH:mm:ss : ") + oUDT.UserFields.Fields.Item("U_STATUS").Value + " - " + _logMessage;

            if (oUDT.Update() != 0)
            {
                EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error updating items to UDT: " + ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription());
            }

            EdiConnectorService.ClearObject(oUDT);
        }
    }
}
