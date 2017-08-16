using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class PlasticaOrderDoc : IEdiDocumentType
    {
        public string TypeName { get; set; } = "Sales Order Document";

        // Data structure
        public string MessageFormat { get; set; }
        public string MessageType { get; set; }
        public string MessageReference { get; set; }
        public string IsTestMessage { get; set; }
        public string IsCrossdock { get; set; }
        public string IsUrgent { get; set; }
        public string IsShopInstallation { get; set; }
        public string IsAcknowledgementRequested { get; set; }
        public string IsCallOffOrder { get; set; }
        public string IsBackhauling { get; set; }
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
        /// <param name="errMsg">The error MSG.</param>
        /// <returns>
        /// List of OrderDocuments
        /// </returns>
        public Object ReadXMLData(XElement _xMessages, out Exception errMsg)
        {
            // Checks if the MessageType is for a Order Response Document.
            // Then it will create new OrderDocuments for every message
            errMsg = null;
            try
            {
                /// <summary>
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
                        IsCrossdock = x.Element("IsCrossdock")?.Value ?? "",
                        IsUrgent = x.Element("IsUrgent")?.Value ?? "",
                        IsShopInstallation = x.Element("IsShopInstallation")?.Value ?? "",
                        IsAcknowledgementRequested = x.Element("IsAcknowledgementRequested")?.Value ?? "",
                        IsCallOffOrder = x.Element("IsCallOffOrder")?.Value ?? "",
                        IsBackhauling = x.Element("IsBackhauling")?.Value ?? "",
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
                errMsg = e;
                return null;
            }
        }

        /// <summary>
        /// Saves specific document data object to SAP.
        /// </summary>
        /// <param name="_dataObject">The data object.</param>
        /// <param name="_connectedServer">The connected server.</param>
        /// <param name="ex">The ex.</param>
        public void SaveToSAP(Object _dataObject, string _connectedServer, out Exception ex)
        {
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            SAPbobsCOM.Documents oOrd = (SAPbobsCOM.Documents)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders));
            string buyerMailAddress = "";
            string buyerMailBody = "";
            int buyerOrderDocumentCount = 0;
            ex = null;

            foreach (PlasticaOrderDoc orderDocument in (List<PlasticaOrderDoc>)_dataObject)
            {
                try
                {
                    oRs.DoQuery(@"SELECT T0.""Address"", T1.""CardCode"", T1.""CardName"" FROM CRD1 T0 INNER JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" WHERE T0.""GlblLocNum"" = '" + orderDocument.InvoiceeGLN + "'");
                    if (oRs.RecordCount > 0)
                    {
                        if (oRs.Fields.Item(0).Size > 0)
                            oOrd.PayToCode = oRs.Fields.Item(0).Value.ToString();
                        else
                        {
                            EventLogger.getInstance().EventError("Server: " + _connectedServer + ". " + "Error Pay To Address not found! With GlblLocNum: " + orderDocument.InvoiceeGLN);
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error Pay To Address not found! With GlblLocNum: " + orderDocument.InvoiceeGLN, "Error!");
                        }
                    }
                    else
                    {
                        EventLogger.getInstance().EventError("Server: " + _connectedServer + ". " + "Error Pay To GlblLocNum: " + orderDocument.InvoiceeGLN + " not found!");
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error Pay To GlblLocNum: " + orderDocument.InvoiceeGLN + " not found!", "Error!");
                    }

                    oRs.DoQuery(@"SELECT T2.""Address"", T0.""CardCode"", T0.""CardName"", T1.""E_MailL"" FROM OCRD T0 INNER JOIN OCPR T1 ON T0.""CardCode"" = T1.""CardCode"" INNER JOIN CRD1 T2 ON T0.""CardCode"" = T2.""CardCode"" WHERE T2.""GlblLocNum"" = '" + orderDocument.BuyerGLN + "'");
                    if (oRs.RecordCount > 0)
                    {
                        string fieldNotFound = "";
                        if (oRs.Fields.Item(0).Size > 0) oOrd.ShipToCode = oRs.Fields.Item(0).Value.ToString();
                        else fieldNotFound += "Error Ship To Address not found! With GlblLocNum: " + orderDocument.BuyerGLN + ". ";

                        if (oRs.Fields.Item(1).Size > 0) oOrd.CardCode = oRs.Fields.Item(1).Value.ToString();
                        else fieldNotFound += "Error Ship To CardCode not found! With GlblLocNum: " + orderDocument.BuyerGLN + ". ";

                        if (oRs.Fields.Item(2).Size > 0) oOrd.CardName = oRs.Fields.Item(2).Value.ToString();
                        else fieldNotFound += "Error Ship To CardName not found! With GlblLocNum: " + orderDocument.BuyerGLN + ". ";

                        if (oRs.Fields.Item(3).Size > 0) buyerMailAddress = oRs.Fields.Item(3).Value.ToString();
                        else fieldNotFound += "Error Ship To E_MailL not found! With GlblLocNum: " + orderDocument.BuyerGLN + ". ";

                        if (fieldNotFound.Length > 0)
                        {
                            EventLogger.getInstance().EventError("Server: " + _connectedServer + " " + fieldNotFound);
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, fieldNotFound, "Error!");
                        }
                    }
                    else
                    {
                        EventLogger.getInstance().EventError("Server: " + _connectedServer + ". " + "Error Ship To GlblLocNum: " + orderDocument.BuyerGLN + " not found!");
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error Ship To GlblLocNum: " + orderDocument.SupplierGLN + " not found!", "Error!");
                    }

                    oOrd.UserFields.Fields.Item("U_TEST").Value = orderDocument.IsTestMessage;
                    oOrd.NumAtCard = orderDocument.OrderNumberBuyer;
                    oOrd.UserFields.Fields.Item("U_DTM_2").Value = orderDocument.RequestedDeliveryDate.ToString("yyyy-MM-dd");
                    oOrd.UserFields.Fields.Item("U_TIJD_2").Value = orderDocument.RequestedDeliveryDate.ToString("HH:mm");
                    oOrd.UserFields.Fields.Item("U_DTM_17").Value = orderDocument.OrderDate.ToString("yyyy-MM-dd");
                    oOrd.UserFields.Fields.Item("U_TIJD_17").Value = orderDocument.OrderDate.ToString("HH:mm");
                    oOrd.UserFields.Fields.Item("U_FLAGS0").Value = orderDocument.IsShopInstallation; // Winkelinstallatie
                    oOrd.UserFields.Fields.Item("U_FLAGS4").Value = orderDocument.IsCrossdock; // Crossdock order
                    oOrd.UserFields.Fields.Item("U_FLAGS7").Value = orderDocument.IsDutyFree; // Accijnsvrije levering
                    oOrd.UserFields.Fields.Item("U_FLAGS8").Value = orderDocument.IsUrgent; // Spoed
                    oOrd.UserFields.Fields.Item("U_FLAGS9").Value = orderDocument.IsBackhauling; // Backhauling ophalen
                    oOrd.UserFields.Fields.Item("U_FLAGS10").Value = orderDocument.IsAcknowledgementRequested; // Bevest. met regels
                    oOrd.DocDueDate = orderDocument.RequestedDeliveryDate;
                    //oOrd.CardName = orderDocument.Sender;
                    //oOrd.UserFields.Fields.Item("U_IsTestMessage").Value = orderDocument.IsTestMessage;
                    //oOrd.UserFields.Fields.Item("U_BGM").Value = orderDocument.OrderNumberBuyer;
                    //oOrd.DocDate = orderDocument.OrderDate;
                    //oOrd.TaxDate = orderDocument.OrderDate;
                    //oOrd.UserFields.Fields.Item("U_DTM_64").Value = orderDocument.EarliestDeliveryDate.ToString("yyyy-MM-dd");
                    //oOrd.UserFields.Fields.Item("U_TIJD_64").Value = orderDocument.EarliestDeliveryDate.ToString("HH:mm");
                    //oOrd.UserFields.Fields.Item("U_DTM_63").Value = orderDocument.LatestDeliveryDate.ToString("yyyy-MM-dd");
                    //oOrd.UserFields.Fields.Item("U_TIJD_63").Value = orderDocument.LatestDeliveryDate.ToString("HH:mm");
                    //oOrd.UserFields.Fields.Item("U_RFF_BO").Value = orderDocument.BlanketOrderNumber;
                    //oOrd.UserFields.Fields.Item("U_RFF_CR").Value = orderDocument.ConsumerReference;
                    //oOrd.UserFields.Fields.Item("U_RFF_PD").Value = orderDocument.ActionNumber;
                    //oOrd.UserFields.Fields.Item("U_RFFCT").Value = orderDocument.; // Contractnummer
                    //oOrd.UserFields.Fields.Item("U_DTMCT").Value = orderDocument.; // Contractdatum
                    //oOrd.UserFields.Fields.Item("U_FLAGS1").Value = orderDocument.; // Geheellevering
                    //oOrd.UserFields.Fields.Item("U_FLAGS2").Value = orderDocument.; // Nul-order
                    //oOrd.UserFields.Fields.Item("U_FLAGS3").Value = orderDocument.; // Aperak gevraagd
                    //oOrd.UserFields.Fields.Item("U_FLAGS5").Value = orderDocument.; // Raamorder
                    //oOrd.UserFields.Fields.Item("U_FLAGS6").Value = orderDocument.; // Geimproviseerde order
                    //oOrd.UserFields.Fields.Item("U_FTXDSI").Value = orderDocument.OrderAdditionalDetails; // Text voor pakbon
                    //oOrd.UserFields.Fields.Item("U_NAD_SF").Value = orderDocument.BuyerGLN; // Eancode haaladres
                    //oOrd.UserFields.Fields.Item("U_NAD_SU").Value = orderDocument.SupplierGLN; // Eancode leverancier
                    //oOrd.UserFields.Fields.Item("U_NAD_UC").Value = orderDocument.UltimateConsigneeGLN; // Eancode eindbestemming
                    //oOrd.UserFields.Fields.Item("U_NAD_BCO").Value = orderDocument.; // Eancode inkoopcombinatie afnemer
                    //oOrd.UserFields.Fields.Item("ONTVANGER").Value = orderDocument.; // Identificatie van uzelf in het EDI-bericht

                    foreach (Article article in orderDocument.Articles)
                    {
                        oRs.DoQuery(@"SELECT ""ItemCode"" FROM OITM WHERE ""CodeBars"" = '" + article.GTIN + "'");
                        if (oRs.RecordCount > 0)
                            oOrd.Lines.ItemCode = oRs.Fields.Item(0).Value.ToString();
                        else
                        {
                            EventLogger.getInstance().EventError("Server: " + _connectedServer + ". " + "Error CodeBars: " + article.GTIN + " not found!");
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error CodeBars: " + article.GTIN + " not found!", "Error!");
                        }
                        oOrd.Lines.ItemDescription = article.ArticleDescription;
                        oOrd.Lines.Quantity = Convert.ToDouble(article.OrderedQuantity);
                        oOrd.Lines.UserFields.Fields.Item("U_LINNR").Value = article.LineNumber;
                        oOrd.Lines.UserFields.Fields.Item("U_LEVARTCODE").Value = article.ArticleCodeSupplier;
                        oOrd.Lines.UserFields.Fields.Item("U_DEUAC").Value = article.GTIN;
                        //oOrd.Lines.UserFields.Fields.Item("U_EdiLineNumber").Value = article.LineNumber;
                        //oOrd.Lines.UserFields.Fields.Item("U_DEARTOM").Value = article.ArticleDescription;
                        //oOrd.Lines.UserFields.Fields.Item("U_KLEUR").Value = article.ColourCode;
                        //oOrd.Lines.UserFields.Fields.Item("U_LENGTE").Value = article.Length;
                        //oOrd.Lines.UserFields.Fields.Item("U_BREEDTE").Value = article.Width;
                        //oOrd.Lines.UserFields.Fields.Item("U_HOOGTE").Value = article.Height;
                        //oOrd.Lines.UserFields.Fields.Item("U_CUX").Value = article.PurchasePriceCurrencyCode;
                        //oOrd.Lines.UserFields.Fields.Item("U_PIA").Value = article.PromotionVariantCode;
                        //oOrd.Lines.UserFields.Fields.Item("U_RFFLI1").Value = article.; // Ordernummer voor onderregel identificatie
                        //oOrd.Lines.UserFields.Fields.Item("U_RFFLI2").Value = article.RequestedDeliveryDate.ToString("yyyy-MM-dd"); // Onderregel identificatie
                        //oOrd.Lines.UserFields.Fields.Item("U_PRI").Value = article.RetailPrice;

                        oOrd.Lines.Add();
                    }

                    if (oOrd.Add() == 0)
                    {
                        string serviceCallID = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetNewObjectKey();
                        oOrd.GetByKey(Convert.ToInt32(serviceCallID));
                        buyerOrderDocumentCount++;
                        buyerMailBody += buyerOrderDocumentCount + " - New Sales Order created with DocNum: " + oOrd.DocNum + System.Environment.NewLine;
                        EventLogger.getInstance().EventInfo("Server: " + _connectedServer + ". " + "Succesfully created Sales Order: " + oOrd.DocNum);
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Succesfully created Sales Order: " + oOrd.DocNum, "Processing..");
                    }
                    else
                    {
                        ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                        EventLogger.getInstance().EventError("Server: " + _connectedServer + ". " + "Error creating Sales Order: (" + errCode + ") " + errMsg);
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error creating Sales Order: (" + errCode + ") " + errMsg, "Error!");
                    }
                }
                catch (Exception e)
                {
                    ex = e;
                    EventLogger.getInstance().EventError("Server: " + _connectedServer + ". " + "Error saving to SAP: " + e.Message + " with order document: " + orderDocument);
                    EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error saving to SAP: " + e.Message + " with order document: " + orderDocument, "Error!");
                }
            }
            ConnectionManager.getInstance().GetConnection(_connectedServer).SendMailNotification("New sales order(s) created:" + buyerOrderDocumentCount, buyerMailBody, buyerMailAddress);

            EdiConnectorService.ClearObject(oOrd);
            EdiConnectorService.ClearObject(oRs);
        }
    }

}
