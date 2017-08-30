using System;

namespace EdiConnectorService_C_Sharp
{
    public class EdiConnectorData
    {
        private static EdiConnectorData instance = null;

        public static EdiConnectorData GetInstance()
        {
            if (instance == null)
            {
                instance = new EdiConnectorData();
            }
            return instance;
        }

        public string ApplicationPath;
        public string ProcessedDirName;
        public string RecordReference;
    }
}
