using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using SAPbobsCOM;

namespace EdiConnectorService_C_Sharp
{
    class UdfFields
    {
        /// <summary>
        /// Creates the udf fields.
        /// </summary>
        /// <param name="_connectedServer">The connected server.</param>
        public static void CreateUdfFields(string _connectedServer)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(EdiConnectorData.getInstance().sApplicationPath + @"\udf.xml");
                XmlNodeList xmlList = xmlDoc.SelectNodes("/fields/udffield");
                foreach (XmlNode xmlNode in xmlList)
                {
                    CreateUdf(_connectedServer, xmlNode["udf"].GetAttribute("table"), xmlNode["udf"].GetAttribute("name"), xmlNode["udf"].GetAttribute("description"), Convert.ToInt32(xmlNode["udf"].GetAttribute("size")),
                        BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, false, false, "");
                }
            }
            catch (Exception e)
            {
                EventLogger.getInstance().EventError("Error reading udf.xml: " + e.Message);
            }
        }

        private static void CreateUdf(string _connectedServer, string _tableName, string _fieldName, string _description, int _editSize, 
            BoFieldTypes _boType, BoFldSubTypes _boSubType, bool _mandatory, bool _default, string _defaultValue)
        {
            IUserFieldsMD oUDF;
            oUDF = (SAPbobsCOM.IUserFieldsMD)ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(BoObjectTypes.oUserFields);
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
                ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out errCode, out errMsg);
                EventLogger.getInstance().EventError("Error creating UDF: " + errMsg + " (" + errCode + ")");
            }
            else
            {
                EventLogger.getInstance().EventInfo("Udf " + _fieldName + " successfully created!");
            }

            EdiConnectorService.ClearObject(oUDF);
        }
    }
}
