using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class PlasticaOrderDoc : IEdiDocumentType
    {
        public string TypeName { get; set; } = "Plastica Sales Order Document";

        // Data structure
        public string MessageFormat { get; set; }
        public string MessageType { get; set; }
        public string MessageReference { get; set; }
        public string IsTestMessage { get; set; }
        public string IsCrossDock { get; set; }
        public string IsUrgent { get; set; }
        public string IsShopInstallation { get; set; }
        public string IsAcknowledgementRequested { get; set; }
        public string IsCallOffOrder { get; set; }
        public string IsBackHauling { get; set; }
        public string IsDutyFree { get; set; }
        public string IsDropShipment { get; set; }
        public string BlanketOrderNumber { get; set; }
        public string ActionNumber { get; set; }
        public string OrderNumberBuyer { get; set; }
        public DateTime OrderDate { get; set; }
        public DateTime EarliestDeliveryDate { get; set; }
        public DateTime LatestDeliveryDate { get; set; }
        public DateTime RequestedDeliveryDate { get; set; }
        public DateTime PickUpDate { get; set; }
        public string BuyerGLN { get; set; }
        public string Buyer { get; set; }
        public string BuyerID { get; set; }
        public string Supplier { get; set; }
        public string SupplierGLN { get; set; }
        public string SupplierID { get; set; }
        public string Invoicee { get; set; }
        public string InvoiceeGLN { get; set; }
        public string InvoiceeID { get; set; }
        public string ShipFromParty { get; set; }
        public string ShipFromPartyID { get; set; }
        public string ShipFromPartyGLN { get; set; }
        public string ShipFromPartyAddress { get; set; }
        public string ShipFromPartyCity { get; set; }
        public string ShipFromPartyZipcode { get; set; }
        public string ShipFromPartyCountry { get; set; }
        public string DeliveryParty { get; set; }
        public string DeliveryPartyID { get; set; }
        public string DeliveryPartyGLN { get; set; }
        public string DeliveryPartyAddress { get; set; }
        public string DeliveryPartyCity { get; set; }
        public string DeliveryPartyZipcode { get; set; }
        public string DeliveryPartyCountry { get; set; }
        public string UltimateConsignee { get; set; }
        public string UltimateConsigneeID { get; set; }
        public string UltimateConsigneeGLN { get; set; }
        public string UltimateConsigneeAddress { get; set; }
        public string UltimateConsigneeZipcode { get; set; }
        public string UltimateConsigneeCity { get; set; }
        public string UltimateConsigneeCountry { get; set; }
        public string CurrencyCode { get; set; }
        public string BuyerAddress { get; set; }
        public string BuyerZipcode { get; set; }
        public string BuyerCity { get; set; }
        public string BuyerCountry { get; set; }
        public string BuyerContactPerson { get; set; }
        public string BuyerContactPersonTelephone { get; set; }
        public string BuyerContactPersonFax { get; set; }
        public string BuyerContactPersonEmail { get; set; }
        public string SupplierAddress { get; set; }
        public string SupplierCity { get; set; }
        public string SupplierZipcode { get; set; }
        public string SupplierCountry { get; set; }
        public string InvoiceeAddress { get; set; }
        public string InvoiceeCity { get; set; }
        public string InvoiceeZipcode { get; set; }
        public string InvoiceeCountry { get; set; }
        public string ConsumerReference { get; set; }
        public string InvoiceeVATNumber { get; set; }
        public string CarrierType { get; set; }
        public string DeliveryPartyContactPersonTelephone { get; set; }
        public string DeliveryPartyContactPersonEmail { get; set; }
        public string PromotionVariantCode { get; set; }
        public string Location { get; set; }
        public string SupplierContactPerson { get; set; }
        public string SupplierContactPersonEmail { get; set; }
        public string SupplierContactPersonFax { get; set; }
        public string SupplierContactPersonTelephone { get; set; }
        public string OrderAdditionalDetails { get; set; }
        public List<Article> Articles { get; set; }

        public class Article
        {
            public string LineNumber { get; set; }
            public string ArticleDescription { get; set; }
            public string ArticleCodeSupplier { get; set; }
            public string GTIN { get; set; }
            public string ArticleCodeBuyer { get; set; }
            public string OrderedQuantity { get; set; }
            public string OrderedQuantityUnitCode { get; set; }
            public string PromotionVariantCode { get; set; }
            public string ColourCode { get; set; }
            public string ColourSystem { get; set; }
            public string Length { get; set; }
            public string LengthUnitCode { get; set; }
            public string Width { get; set; }
            public string WidthUnitCode { get; set; }
            public string Height { get; set; }
            public string HeightUnitCode { get; set; }
            public string RetailPrice { get; set; }
            public string RetailPriceCurrencyCode { get; set; }
            public string PurchasePrice { get; set; }
            public string PurchasePriceCurrencyCode { get; set; }
            public string Location { get; set; }
            public string LocationID { get; set; }
            public string LocationGLN { get; set; }
            public string GrossWeight { get; set; }
            public string GrossWeightUnitCode { get; set; }
            public DateTime RequestedDeliveryDate { get; set; }
            public DateTime LatestDeliveryDate { get; set; }
            public string NetLineAmount { get; set; }
            public string PackagingType { get; set; }
            public string ConsumerReference { get; set; }
            public string UnspecifiedText { get; set; }
        }

        /// <summary>
        /// Reads the XML data.
        /// </summary>
        /// <param name="_xMessages">The incoming XML messages.</param>
        /// <param name="_exception">The error message.</param>
        /// <returns>
        /// List of OrderDocuments
        /// </returns>
        public Object ReadXMLData(XElement _xMessages, out string _exception)
        {
            _exception = null;
            try
            {
                /// <summary>
                /// Checks if the MessageType is for a Order Response Document.
                /// Then it will create new OrderDocuments for every message
                /// 
                /// ?? "" null-coalescing operator is used to check if the xml element is existant, else it will return ""
                /// By not using this operator this will not fill the document and return an empty list.
                /// This method is used to make sure the important values are read from the xml files and optional values will be set to ""
                /// </summary>
                List<PlasticaOrderDoc> OrderMsgList = _xMessages.Elements().Where(x => x.Element("MessageType").Value == "3").Select(x =>
                    new PlasticaOrderDoc()
                    {
                        MessageFormat = x.Element("MessageFormat").Value ?? "",
                        MessageType = x.Element("MessageType").Value ?? "",
                        MessageReference = x.Element("MessageReference")?.Value ?? "",
                        IsTestMessage = x.Element("IsTestMessage")?.Value ?? "",
                        IsCrossDock = x.Element("IsCrossDock")?.Value ?? "",
                        IsUrgent = x.Element("IsUrgent")?.Value ?? "",
                        IsShopInstallation = x.Element("IsShopInstallation")?.Value ?? "",
                        IsAcknowledgementRequested = x.Element("IsAcknowledgementRequested")?.Value ?? "",
                        IsCallOffOrder = x.Element("IsCallOffOrder")?.Value ?? "",
                        IsBackHauling = x.Element("IsBackHauling")?.Value ?? "",
                        IsDutyFree = x.Element("IsDutyFree")?.Value ?? "",
                        IsDropShipment = x.Element("IsDropShipment")?.Value ?? "",
                        BlanketOrderNumber = x.Element("BlanketOrderNumber")?.Value ?? "",
                        ActionNumber = x.Element("ActionNumber")?.Value ?? "",
                        OrderNumberBuyer = x.Element("OrderNumberBuyer")?.Value ?? "",
                        OrderDate = DateTime.ParseExact(x.Element("OrderDate")?.Value ?? "000101010000", "yyyyMMddHHmm", System.Globalization.CultureInfo.InvariantCulture),
                        EarliestDeliveryDate = DateTime.ParseExact(x.Element("EarliestDeliveryDate")?.Value ?? "00010101", "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                        LatestDeliveryDate = DateTime.ParseExact(x.Element("LatestDeliveryDate")?.Value ?? "00010101", "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                        RequestedDeliveryDate = DateTime.ParseExact(x.Element("RequestedDeliveryDate")?.Value ?? "00010101", "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                        PickUpDate = DateTime.ParseExact(x.Element("PickUpDate")?.Value ?? "000101010000", "yyyyMMddHHmm", System.Globalization.CultureInfo.InvariantCulture),
                        BuyerGLN = x.Element("BuyerGLN")?.Value ?? "",
                        Buyer = x.Element("Buyer")?.Value ?? "",
                        BuyerID = x.Element("BuyerID")?.Value ?? "",
                        Supplier = x.Element("Supplier")?.Value ?? "",
                        SupplierGLN = x.Element("SupplierGLN")?.Value ?? "",
                        SupplierID = x.Element("SupplierID")?.Value ?? "",
                        Invoicee = x.Element("Invoicee")?.Value ?? "",
                        InvoiceeGLN = x.Element("InvoiceeGLN")?.Value ?? "",
                        InvoiceeID = x.Element("InvoiceeID")?.Value ?? "",
                        ShipFromParty = x.Element("ShipFromParty")?.Value ?? "",
                        ShipFromPartyID = x.Element("ShipFromPartyID")?.Value ?? "",
                        ShipFromPartyGLN = x.Element("ShipFromPartyGLN")?.Value ?? "",
                        ShipFromPartyAddress = x.Element("ShipFromPartyAddress")?.Value ?? "",
                        ShipFromPartyCity = x.Element("ShipFromPartyCity")?.Value ?? "",
                        ShipFromPartyZipcode = x.Element("ShipFromPartyZipcode")?.Value ?? "",
                        ShipFromPartyCountry = x.Element("ShipFromPartyCountry")?.Value ?? "",
                        DeliveryParty = x.Element("DeliveryParty")?.Value ?? "",
                        DeliveryPartyID = x.Element("DeliveryPartyID")?.Value ?? "",
                        DeliveryPartyGLN = x.Element("DeliveryPartyGLN")?.Value ?? "",
                        DeliveryPartyAddress = x.Element("DeliveryPartyAddress")?.Value ?? "",
                        DeliveryPartyCity = x.Element("DeliveryPartyCity")?.Value ?? "",
                        DeliveryPartyZipcode = x.Element("DeliveryPartyZipcode")?.Value ?? "",
                        DeliveryPartyCountry = x.Element("DeliveryPartyCountry")?.Value ?? "",
                        UltimateConsignee = x.Element("UltimateConsignee")?.Value ?? "",
                        UltimateConsigneeID = x.Element("UltimateConsigneeID")?.Value ?? "",
                        UltimateConsigneeGLN = x.Element("UltimateConsigneeGLN")?.Value ?? "",
                        UltimateConsigneeAddress = x.Element("UltimateConsigneeAddress")?.Value ?? "",
                        UltimateConsigneeZipcode = x.Element("UltimateConsigneeZipcode")?.Value ?? "",
                        UltimateConsigneeCity = x.Element("UltimateConsigneeCity")?.Value ?? "",
                        UltimateConsigneeCountry = x.Element("UltimateConsigneeCountry")?.Value ?? "",
                        CurrencyCode = x.Element("CurrencyCode")?.Value ?? "",
                        BuyerAddress = x.Element("BuyerAddress")?.Value ?? "",
                        BuyerZipcode = x.Element("BuyerZipcode")?.Value ?? "",
                        BuyerCity = x.Element("BuyerCity")?.Value ?? "",
                        BuyerCountry = x.Element("BuyerCountry")?.Value ?? "",
                        BuyerContactPerson = x.Element("BuyerContactPerson")?.Value ?? "",
                        BuyerContactPersonTelephone = x.Element("BuyerContactPersonTelephone")?.Value ?? "",
                        BuyerContactPersonFax = x.Element("BuyerContactPersonFax")?.Value ?? "",
                        BuyerContactPersonEmail = x.Element("BuyerContactPersonEmail")?.Value ?? "",
                        SupplierAddress = x.Element("SupplierAddress")?.Value ?? "",
                        SupplierCity = x.Element("SupplierCity")?.Value ?? "",
                        SupplierZipcode = x.Element("SupplierZipcode")?.Value ?? "",
                        SupplierCountry = x.Element("SupplierCountry")?.Value ?? "",
                        InvoiceeAddress = x.Element("InvoiceeAddress")?.Value ?? "",
                        InvoiceeCity = x.Element("InvoiceeCity")?.Value ?? "",
                        InvoiceeZipcode = x.Element("InvoiceeZipcode")?.Value ?? "",
                        InvoiceeCountry = x.Element("InvoiceeCountry")?.Value ?? "",
                        ConsumerReference = x.Element("ConsumerReference")?.Value ?? "",
                        InvoiceeVATNumber = x.Element("InvoiceeVATNumber")?.Value ?? "",
                        CarrierType = x.Element("CarrierType")?.Value ?? "",
                        DeliveryPartyContactPersonTelephone = x.Element("DeliveryPartyContactPersonTelephone")?.Value ?? "",
                        DeliveryPartyContactPersonEmail = x.Element("DeliveryPartyContactPersonEmail")?.Value ?? "",
                        PromotionVariantCode = x.Element("PromotionVariantCode")?.Value ?? "",
                        Location = x.Element("Location")?.Value ?? "",
                        SupplierContactPerson = x.Element("SupplierContactPerson")?.Value ?? "",
                        SupplierContactPersonEmail = x.Element("SupplierContactPersonEmail")?.Value ?? "",
                        SupplierContactPersonFax = x.Element("SupplierContactPersonFax")?.Value ?? "",
                        SupplierContactPersonTelephone = x.Element("SupplierContactPersonTelephone")?.Value ?? "",
                        OrderAdditionalDetails = x.Element("OrderAdditionalDetails")?.Value ?? "",
                        Articles = x.Elements("Article")?.Select(a => new Article()
                        {
                            LineNumber = a.Element("LineNumber")?.Value ?? "",
                            ArticleDescription = a.Element("ArticleDescription")?.Value ?? "",
                            ArticleCodeSupplier = a.Element("ArticleCodeSupplier")?.Value ?? "",
                            GTIN = a.Element("GTIN")?.Value ?? "",
                            ArticleCodeBuyer = a.Element("ArticleCodeBuyer")?.Value ?? "",
                            OrderedQuantity = a.Element("OrderedQuantity")?.Value ?? "",
                            OrderedQuantityUnitCode = a.Element("OrderedQuantityUnitCode")?.Value ?? "",
                            PromotionVariantCode = a.Element("PromotionVariantCode")?.Value ?? "",
                            ColourCode = a.Element("ColourCode")?.Value ?? "",
                            ColourSystem = a.Element("ColourSystem")?.Value ?? "",
                            Length = a.Element("Length")?.Value ?? "",
                            LengthUnitCode = a.Element("LengthUnitCode")?.Value ?? "",
                            Width = a.Element("Width")?.Value ?? "",
                            WidthUnitCode = a.Element("WidthUnitCode")?.Value ?? "",
                            Height = a.Element("Height")?.Value ?? "",
                            HeightUnitCode = a.Element("HeightUnitCode")?.Value ?? "",
                            RetailPrice = a.Element("RetailPrice")?.Value ?? "",
                            RetailPriceCurrencyCode = a.Element("RetailPriceCurrencyCode")?.Value ?? "",
                            PurchasePrice = a.Element("PurchasePrice")?.Value ?? "",
                            PurchasePriceCurrencyCode = a.Element("PurchasePriceCurrencyCode")?.Value ?? "",
                            Location = a.Element("Location")?.Value ?? "",
                            LocationID = a.Element("LocationID")?.Value ?? "",
                            LocationGLN = a.Element("LocationGLN")?.Value ?? "",
                            GrossWeight = a.Element("GrossWeight")?.Value ?? "",
                            GrossWeightUnitCode = a.Element("GrossWeightUnitCode")?.Value ?? "",
                            RequestedDeliveryDate = DateTime.ParseExact(a.Element("RequestedDeliveryDate")?.Value ?? "000101010000", "yyyyMMddHHmm", System.Globalization.CultureInfo.InvariantCulture),
                            LatestDeliveryDate = DateTime.ParseExact(a.Element("LatestDeliveryDate")?.Value ?? "000101010000", "yyyyMMddHHmm", System.Globalization.CultureInfo.InvariantCulture),
                            NetLineAmount = a.Element("NetLineAmount")?.Value ?? "",
                            PackagingType = a.Element("PackagingType")?.Value ?? "",
                            ConsumerReference = a.Element("ConsumerReference")?.Value ?? "",
                            UnspecifiedText = a.Element("UnspecifiedText")?.Value ?? "",
                        }).ToList()
                    }).ToList();
                return OrderMsgList;
            }
            catch (Exception e)
            {
                _exception = e.Message;
                return null;
            }
        }

        /// <summary>
        /// Saves specific document data object to SAP.
        /// </summary>
        /// <param name="_dataObject">The data object.</param>
        /// <param name="_connectedServer">The connected server.</param>
        /// <param name="exception">The exception.</param>
        public void SaveToSAP(Object _dataObject, string _connectedServer, out string exception)
        {
            // Initialize values
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            SAPbobsCOM.Documents oOrd = (SAPbobsCOM.Documents)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders));
            string buyerMailAddress = "";
            string buyerMailBody = "";
            int buyerOrderDocumentCount = 0;
            exception = null;

            // Iterate through every message in the data object
            foreach (PlasticaOrderDoc orderDocument in (List<PlasticaOrderDoc>)_dataObject)
            {
                try
                {
                    // Execute query to search for CardCode, CardName, Email in OCRD, Email in OCPR by looking up the Buyer GLN in OCRD or CRD1
                    oRs.DoQuery($@"SELECT T0.""CardCode"", T0.""CardName"", T0.""E_Mail"", T2.""E_MailL"" FROM OCRD T0 INNER JOIN CRD1 T1 ON T0.""CardCode"" = T1.""CardCode"" INNER JOIN OCPR T2 ON T1.""CardCode"" = T2.""CardCode"" WHERE T0.""U_PLA_EAN"" = '{orderDocument.BuyerGLN}' GROUP BY T0.""CardName"", T0.""CardCode"", T0.""E_Mail"", T2.""E_MailL""");
                    if(oRs.RecordCount > 0)
                    {
                        // Set email to query result "E_Mail" if its not null or empty else use "E_MailL"
                        string email = !String.IsNullOrEmpty(oRs.Fields.Item(2).Value.ToString()) ? oRs.Fields.Item(2).Value.ToString() : oRs.Fields.Item(3).Value.ToString();
                        string cardName = oRs.Fields.Item(1).Value.ToString();
                        string cardCode = oRs.Fields.Item(0).Value.ToString();
                        
                        buyerMailAddress = email;
                        oOrd.CardName = cardName;
                        oOrd.CardCode = cardCode;

                        // Execute query to search for Pay To Address by looking up the CardCode, Invoicee GLN and the Address Type "B"
                        oRs.DoQuery($@"SELECT T0.""Address"" FROM CRD1 T0 WHERE T0.""CardCode"" = '{cardCode}' AND T0.""U_AddressGLN"" = '{orderDocument.InvoiceeGLN}' AND T0.""AdresType"" = 'B'");
                        if (oRs.RecordCount > 0)
                            oOrd.PayToCode = oRs.Fields.Item(0).Value.ToString(); // Set PayToCode when query found a result
                        else
                        {
                            EventLogger.getInstance().EventWarning($"Server: {_connectedServer}. Error Pay To Address not found! With U_AddressGLN: {orderDocument.InvoiceeGLN}. Using default address.");
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.GetInstance().RecordReference, $"Error Pay To Address not found! With U_AddressGLN: {orderDocument.InvoiceeGLN}. Using default address", "Warning");
                            oRs.DoQuery($@"SELECT T0.""Address"" FROM CRD1 T0.""CardCode"" = '{cardCode}' AND T0.""AdresType"" = 'B'");
                            if (oRs.RecordCount > 0)
                                oOrd.PayToCode = oRs.Fields.Item(0).Value.ToString();
                        }

                        // For the Ship To GLN use DeliveryPartyGLN if it contains a result, otherwise use BuyerGLN
                        string shipToGLN = orderDocument.DeliveryPartyGLN != "" ? orderDocument.DeliveryPartyGLN : orderDocument.BuyerGLN;

                        // Execute query to search for Ship To Address by looking up the CardCode, shipToGLN and the Address Type "S"
                        oRs.DoQuery($@"SELECT T0.""Address"" FROM CRD1 T0 WHERE T0.""CardCode"" = '{cardCode}' AND T0.""U_AddressGLN"" = '{shipToGLN}' AND T0.""AdresType"" = 'S'");
                        if (oRs.RecordCount > 0)
                            oOrd.ShipToCode = oRs.Fields.Item(0).Value.ToString(); // Set ShipToCode when query found a result
                        else
                        {
                            EventLogger.getInstance().EventWarning($"Server: {_connectedServer}. Error Ship To Address not found! With U_AddressGLN: {shipToGLN}. Using default address");
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.GetInstance().RecordReference, $"Error Ship To Address not found! With U_AddressGLN: {shipToGLN}. Using default address", "Warning");
                            oRs.DoQuery($@"SELECT T0.""Address"" FROM CRD1 T0.""CardCode"" = '{cardCode}' AND T0.""AdresType"" = 'S'");
                            if (oRs.RecordCount > 0)
                                oOrd.ShipToCode = oRs.Fields.Item(0).Value.ToString();
                        }
                    }
                    else
                    {
                        EventLogger.getInstance().EventError($"Server: {_connectedServer}. Error U_PLA_EAN: {orderDocument.BuyerGLN} not found! Cannot find CardCode, CardName or email");
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.GetInstance().RecordReference, $"Error Pay To U_PLA_EAN: {orderDocument.BuyerGLN} not found! Cannot find CardCode, CardName or email", "Error!");
                    }

                    // Execute query to search for Sales Employee code by looking up the Sales employee name "EDI"
                    oRs.DoQuery(@"SELECT T0.""SlpCode"" FROM OSLP T0 WHERE T0.""SlpName"" = 'EDI'");
                    if(oRs.RecordCount > 0)
                        oOrd.SalesPersonCode = Convert.ToInt32(oRs.Fields.Item(0).Value); // Set Sales person code when query found a result

                    // Fill in data from order document to defined fields in SAP
                    oOrd.NumAtCard = orderDocument.OrderNumberBuyer;
                    oOrd.DocDueDate = orderDocument.RequestedDeliveryDate;
                    oOrd.UserFields.Fields.Item("U_MessageFormat").Value = orderDocument.MessageFormat;
                    oOrd.UserFields.Fields.Item("U_MessageType").Value = orderDocument.MessageType;
                    oOrd.UserFields.Fields.Item("U_IsTestMessage").Value = orderDocument.IsTestMessage;
                    oOrd.UserFields.Fields.Item("U_MessageReference").Value = orderDocument.MessageReference;
                    oOrd.UserFields.Fields.Item("U_OrderDate").Value = orderDocument.OrderDate.ToString("yyyy-MM-dd HH:mm");
                    oOrd.UserFields.Fields.Item("U_IsCrossDock").Value = orderDocument.IsCrossDock;
                    oOrd.UserFields.Fields.Item("U_IsCallOffOrder").Value = orderDocument.IsCallOffOrder;
                    oOrd.UserFields.Fields.Item("U_RequestedDeliveryDate").Value = orderDocument.RequestedDeliveryDate.ToString("yyyy-MM-dd");
                    oOrd.UserFields.Fields.Item("U_EarliestDeliveryDate").Value = orderDocument.EarliestDeliveryDate.ToString("yyyy-MM-dd");
                    oOrd.UserFields.Fields.Item("U_LatestDeliveryDate").Value = orderDocument.LatestDeliveryDate.ToString("yyyy-MM-dd");
                    oOrd.UserFields.Fields.Item("U_IsShopInstallation").Value = orderDocument.IsShopInstallation;
                    oOrd.UserFields.Fields.Item("U_BlanketOrderNumber").Value = orderDocument.BlanketOrderNumber;
                    oOrd.UserFields.Fields.Item("U_ConsumerReference").Value = orderDocument.ConsumerReference;
                    oOrd.UserFields.Fields.Item("U_ActionNumber").Value = orderDocument.ActionNumber;
                    oOrd.UserFields.Fields.Item("U_BuyerGLN").Value = orderDocument.BuyerGLN;
                    oOrd.UserFields.Fields.Item("U_BuyerID").Value = orderDocument.BuyerID;
                    oOrd.UserFields.Fields.Item("U_Buyer").Value = orderDocument.Buyer;
                    oOrd.UserFields.Fields.Item("U_SupplierGLN").Value = orderDocument.SupplierGLN;
                    oOrd.UserFields.Fields.Item("U_Supplier").Value = orderDocument.Supplier;
                    oOrd.UserFields.Fields.Item("U_Invoicee").Value = orderDocument.Invoicee;
                    oOrd.UserFields.Fields.Item("U_InvoiceeGLN").Value = orderDocument.InvoiceeGLN;
                    oOrd.UserFields.Fields.Item("U_DeliveryParty").Value = orderDocument.DeliveryParty;
                    oOrd.UserFields.Fields.Item("U_DeliveryPartyGLN").Value = orderDocument.DeliveryPartyGLN;
                    oOrd.UserFields.Fields.Item("U_DeliveryPartyAddress").Value = orderDocument.DeliveryPartyAddress;
                    oOrd.UserFields.Fields.Item("U_DeliveryPartyZipcode").Value = orderDocument.DeliveryPartyZipcode;
                    oOrd.UserFields.Fields.Item("U_DeliveryPartyCity").Value = orderDocument.DeliveryPartyCity;
                    oOrd.UserFields.Fields.Item("U_DeliveryPartyCountry").Value = orderDocument.DeliveryPartyCountry;
                    oOrd.UserFields.Fields.Item("U_ShipFromPartyGLN").Value = orderDocument.ShipFromPartyGLN;
                    oOrd.UserFields.Fields.Item("U_UltimateConsignee").Value = orderDocument.UltimateConsignee;
                    oOrd.UserFields.Fields.Item("U_UltimateConsigneeGLN").Value = orderDocument.UltimateConsigneeGLN;
                    oOrd.UserFields.Fields.Item("U_UltimateConsigneeAddress").Value = orderDocument.UltimateConsigneeAddress;
                    oOrd.UserFields.Fields.Item("U_UltimateConsigneeZipcode").Value = orderDocument.UltimateConsigneeZipcode;
                    oOrd.UserFields.Fields.Item("U_UltimateConsigneeCity").Value = orderDocument.UltimateConsigneeCity;
                    oOrd.UserFields.Fields.Item("U_UltimateConsigneeCountry").Value = orderDocument.UltimateConsigneeCountry;
                    oOrd.UserFields.Fields.Item("U_IsUrgent").Value = orderDocument.IsUrgent;
                    oOrd.UserFields.Fields.Item("U_IsDutyFree").Value = orderDocument.IsDutyFree;
                    oOrd.UserFields.Fields.Item("U_IsBackHauling").Value = orderDocument.IsBackHauling;
                    oOrd.UserFields.Fields.Item("U_IsAcknowledgedRequested").Value = orderDocument.IsAcknowledgementRequested;
                    oOrd.DocDate = orderDocument.OrderDate;
                    oOrd.TaxDate = orderDocument.OrderDate;

                    foreach (Article article in orderDocument.Articles)
                    {
                        oRs.DoQuery(@"SELECT ""ItemCode"", ""ItemName"" FROM OITM WHERE ""CodeBars"" = '" + article.GTIN + "'");
                        if (oRs.RecordCount > 0)
                        {
                            oOrd.Lines.ItemCode = oRs.Fields.Item(0).Value.ToString();
                            oOrd.Lines.ItemDescription = article.ArticleDescription != "" ? article.ArticleDescription : oRs.Fields.Item(1).Value.ToString();
                        }
                        else
                        {
                            EventLogger.getInstance().EventError($"Server: {_connectedServer}. Error CodeBars: {article.GTIN} not found!");
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.GetInstance().RecordReference, $"Error CodeBars: {article.GTIN} not found!", "Error!");
                        }
                        oOrd.Lines.UserFields.Fields.Item("U_LineNumber").Value = article.LineNumber;
                        oOrd.Lines.UserFields.Fields.Item("U_ArticleCodeSupplier").Value = article.ArticleCodeSupplier;
                        oOrd.Lines.Quantity = Convert.ToDouble(article.OrderedQuantity);
                        oOrd.Lines.UserFields.Fields.Item("U_ColourCode").Value = article.ColourCode;
                        oOrd.Lines.UserFields.Fields.Item("U_ColourSystem").Value = article.ColourSystem;
                        oOrd.Lines.UserFields.Fields.Item("U_Length").Value = article.Length;
                        oOrd.Lines.UserFields.Fields.Item("U_LengthUnitCode").Value = article.LengthUnitCode;
                        oOrd.Lines.UserFields.Fields.Item("U_Width").Value = article.Width;
                        oOrd.Lines.UserFields.Fields.Item("U_WidthUnitCode").Value = article.WidthUnitCode;
                        oOrd.Lines.UserFields.Fields.Item("U_Height").Value = article.Height;
                        oOrd.Lines.UserFields.Fields.Item("U_HeightUnitCode").Value = article.HeightUnitCode;
                        oOrd.Lines.UserFields.Fields.Item("U_RetailPrice").Value = article.RetailPrice;
                        oOrd.Lines.UserFields.Fields.Item("U_PromotionVariantCode").Value = article.PromotionVariantCode;
                        oOrd.Lines.UserFields.Fields.Item("U_ArticleCodeBuyer").Value = article.ArticleCodeBuyer;
                        oOrd.Lines.UserFields.Fields.Item("U_ArticleDescription").Value = article.ArticleDescription;
                        oOrd.Lines.UserFields.Fields.Item("U_LocationGLN").Value = article.LocationGLN;

                        oOrd.Lines.Add();
                    }

                    // Try to create a document
                    if (oOrd.Add() == 0)
                    {
                        // Get latest created document to log the document number
                        string orderDocNum = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetNewObjectKey();
                        oOrd.GetByKey(Convert.ToInt32(orderDocNum));

                        // Keep up with the created document count in the data object
                        buyerOrderDocumentCount++;
                        // Build up the mail body with the created document
                        buyerMailBody += $"{buyerOrderDocumentCount} - New Sales Order created with DocNum: {oOrd.DocNum}{Environment.NewLine}";
                        EventLogger.getInstance().EventInfo($"Server: {_connectedServer}. Succesfully created Sales Order: {oOrd.DocNum}");
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.GetInstance().RecordReference, $"Succesfully created Sales Order: {oOrd.DocNum}", "Processing..", oOrd.DocNum.ToString());
                    }
                    else
                    {
                        ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                        EventLogger.getInstance().EventError($"Server: {_connectedServer}. Error creating Sales Order: ({errCode}) {errMsg}");
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.GetInstance().RecordReference, $"Error creating Sales Order: ({errCode}) {errMsg}", "Error!");
                        exception = $"({errCode}) {errMsg}";
                    }
                }
                catch (Exception e)
                {
                    exception = e.Message;
                    EventLogger.getInstance().EventError($"Server: {_connectedServer}. Error saving to SAP: {e.Message} with order document: {orderDocument}");
                    EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.GetInstance().RecordReference, $"Error saving to SAP: {e.Message} with order document: {orderDocument}", "Error!");
                }
            }

            // Send a mail notification if a document has been created
            if(buyerOrderDocumentCount > 0)
            {
                //ConnectionManager.getInstance().GetConnection(_connectedServer).SendMailNotification("New sales order(s) created:" + buyerOrderDocumentCount, buyerMailBody, buyerMailAddress);
            }

            EdiConnectorService.ClearObject(oOrd);
            EdiConnectorService.ClearObject(oRs);
        }
    }

}
