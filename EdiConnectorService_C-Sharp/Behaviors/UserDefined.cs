using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using SAPbobsCOM;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class UserDefined
    {

        /// <summary>
        /// Creates the UDT tables from the "udf.xml".
        /// </summary>
        /// <param name="_connectedServer">The connected server.</param>
        public static void CreateTables(string _connectedServer)
        {
            try
            {
                XDocument xDoc = XDocument.Load(EdiConnectorData.getInstance().sApplicationPath + @"udf.xml");
                foreach (XElement xEle in xDoc.Element("UserDefined").Element("Tables").Descendants("Udt"))
                {
                    CreateTable(_connectedServer, xEle.Attribute("name").Value, BoUTBTableType.bott_NoObjectAutoIncrement);
                }
            }
            catch 
            {
                //EventLogger.getInstance().EventError("Error creating UDT - Message: " + e.Message);
            }  
        }

        private static void CreateTable(string _connectedServer, string _tableName, BoUTBTableType _tableType)
        {
            UserTablesMD oUDT;
            oUDT = (UserTablesMD)ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(BoObjectTypes.oUserTables);

            if (!oUDT.GetByKey(_tableName))
            {
                oUDT.TableName = _tableName;
                oUDT.TableDescription = _tableName;
                oUDT.TableType = _tableType;

                if (oUDT.Add() != 0)
                {
                    EventLogger.getInstance().EventError("Error creating UDT: " + ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastErrorDescription());
                }
                else
                {
                    EventLogger.getInstance().EventInfo("UDT " + _tableName + " successfully created!");
                }
            }

            EdiConnectorData.getInstance().sUdtName.Add(_tableName);
            EdiConnectorService.ClearObject(oUDT);
        }

        /// <summary>
        /// Creates the UDF fields from the "udf.xml".
        /// </summary>
        /// <param name="_connectedServer">The connected server.</param>
        public static void CreateFields(string _connectedServer)
        {
            try
            {
                XDocument xDoc = XDocument.Load(EdiConnectorData.getInstance().sApplicationPath + @"udf.xml");
                foreach (XElement xEle in xDoc.Element("UserDefined").Element("Fields").Descendants("Udf"))
                {
                    var typeAttribute = (string)xEle.Attribute("type") ?? "alpha";
                    var subTypeAttribute = (string)xEle.Attribute("subtype") ?? "none";
                    var sizeAttribute = (string)xEle.Attribute("size") ?? "1";
                    CreateField(_connectedServer, xEle.Attribute("table").Value, xEle.Attribute("name").Value, xEle.Attribute("description").Value, 
                        Convert.ToInt32(sizeAttribute), GetFieldType(typeAttribute), GetFieldSubType(subTypeAttribute), false, false, "");
                }
            }
            catch
            {
                //EventLogger.getInstance().EventError("Error creating UDF: " + e.Message);
            }
        }

        private static void CreateField(string _connectedServer, string _tableName, string _fieldName, string _description, int _editSize, 
            BoFieldTypes _boType, BoFldSubTypes _boSubType, bool _mandatory, bool _default, string _defaultValue)
        {
            IUserFieldsMD oUDF = (IUserFieldsMD)ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(BoObjectTypes.oUserFields);
            UserTable oUDT = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.UserTables.Item(_tableName);

            foreach(IField existingUDF in oUDT.UserFields.Fields)
            {
                if (existingUDF.Name == "U_"+_fieldName)
                    return;
            }
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
                ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                EventLogger.getInstance().EventError("Error creating UDF: " + errMsg);
            }
            else
            {
                EventLogger.getInstance().EventInfo("UDF " + _fieldName + " successfully created!");
            }

            EdiConnectorService.ClearObject(oUDF);
            EdiConnectorService.ClearObject(oUDT);
        }

        /// <summary>
        /// Gets the type of the field.
        /// </summary>
        /// <param name="_type">The type.</param>
        /// <returns></returns>
        private static BoFieldTypes GetFieldType(string _type)
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

        private static BoFldSubTypes GetFieldSubType(string _type)
        {
            if (_type == "none")
                return BoFldSubTypes.st_None;
            else if (_type == "link")
                return BoFldSubTypes.st_Link;
            else if (_type == "time")
                return BoFldSubTypes.st_Time;
            else
                return BoFldSubTypes.st_None;
        }
    }
}
