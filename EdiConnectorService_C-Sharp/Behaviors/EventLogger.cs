using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            System.GC.Collect();
        }

        public void EventError(string text)
        {
            eventLog1.WriteEntry(text, EventLogEntryType.Error);
            System.GC.Collect();
        }

        public void EventWarning(string text)
        {
            eventLog1.WriteEntry(text, EventLogEntryType.Warning);
            System.GC.Collect();
        }
    }
}
