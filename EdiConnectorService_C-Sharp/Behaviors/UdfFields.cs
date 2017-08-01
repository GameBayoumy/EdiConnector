using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using SAPbobsCOM;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class UdfFields
    {

        /// <summary>
        /// Creates the UDT tables from the "udf.xml".
        /// </summary>
        /// <param name="_connectedServer">The connected server.</param>
        public static void CreateUdtTables(string _connectedServer)
        {
            try
            {
                XDocument xDoc = XDocument.Load(EdiConnectorData.getInstance().sApplicationPath + @"\udf.xml");
                foreach (XElement xEle in xDoc.Element("UserDefined").Element("Tables").Descendants("Udt"))
                {
                    CreateUdt(_connectedServer, xEle.Attribute("name").Value, BoUTBTableType.bott_NoObjectAutoIncrement);
                }
            }
            catch (Exception e)
            {
                EventLogger.getInstance().EventError("Error creating UDT - Message: " + e.Message);
            }  
        }

        private static void CreateUdt(string _connectedServer, string _tableName, SAPbobsCOM.BoUTBTableType _tableType)
        {
            SAPbobsCOM.UserTablesMD oUDT;
            oUDT = (SAPbobsCOM.UserTablesMD)ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(BoObjectTypes.oUserTables);
            int errCode;
            string errMsg = "";

            if (!oUDT.GetByKey(_tableName))
            {
                oUDT.TableName = _tableName;
                oUDT.TableDescription = _tableName;
                oUDT.TableType = _tableType;

                if (oUDT.Add() != 0)
                {
                    ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out errCode, out errMsg);
                    EventLogger.getInstance().EventError("Error creating UDT: " + errMsg + " (" + errCode + ")");
                }
                else
                {
                    EventLogger.getInstance().EventInfo("UDT " + _tableName + " successfully created!");
                }
            }

            EdiConnectorService.ClearObject(oUDT);
        }

        /// <summary>
        /// Creates the UDF fields from the "udf.xml".
        /// </summary>
        /// <param name="_connectedServer">The connected server.</param>
        public static void CreateUdfFields(string _connectedServer)
        {
            try
            {
                XDocument xDoc = XDocument.Load(EdiConnectorData.getInstance().sApplicationPath + @"\udf.xml");
                foreach (XElement xEle in xDoc.Element("UserDefined").Element("Fields").Descendants("Udf"))
                {
                    CreateUdf(_connectedServer, xEle.Attribute("table").Value, xEle.Attribute("name").Value, xEle.Attribute("description").Value, 
                        Convert.ToInt32(xEle.Attribute("size").Value), GetFieldType(xEle.Attribute("type").Value), BoFldSubTypes.st_None, false, false, "");
                }
            }
            catch (Exception e)
            {
                EventLogger.getInstance().EventError("Error creating UDF: " + e.Message);
            }
        }

        private static void CreateUdf(string _connectedServer, string _tableName, string _fieldName, string _description, int _editSize, 
            BoFieldTypes _boType, BoFldSubTypes _boSubType, bool _mandatory, bool _default, string _defaultValue)
        {
            IUserFieldsMD oUDF;
            oUDF = (SAPbobsCOM.IUserFieldsMD)ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(BoObjectTypes.oUserFields);
            
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
                int errCode;
                string errMsg = "";
                ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out errCode, out errMsg);
                EventLogger.getInstance().EventError("Error creating UDF: " + errMsg + " (" + errCode + ")");
            }
            else
            {
                EventLogger.getInstance().EventInfo("UDF " + _fieldName + " successfully created!");
            }

            EdiConnectorService.ClearObject(oUDF);
        }

        private static SAPbobsCOM.BoFieldTypes GetFieldType(string _type)
        {
            if (_type == "alpha")
                return BoFieldTypes.db_Alpha;
            else if (_type == "memo")
                return BoFieldTypes.db_Memo;
            else if (_type == "date")
                return BoFieldTypes.db_Date;
            else if (_type == "float")
                return BoFieldTypes.db_Float;
            else if (_type == "numeric")
                return BoFieldTypes.db_Numeric;
            else
                return BoFieldTypes.db_Alpha;
        }
    }
}
