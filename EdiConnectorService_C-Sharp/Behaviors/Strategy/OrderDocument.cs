using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class OrderDocument : IEdiDocumentType
    {
        public string TypeName { get; set; } = "Sales Order Document";

        // Data structure
        public string MessageStandard { get; set;}
        public string MessageType { get; set; }
        public string Sender { get; set; }
        public string SenderGLN { get; set; }
        public string RecipientGLN { get; set; }
        public string IsTestMessage { get; set; }
        public string OrderNumberBuyer { get; set; }
        public string RequestedDeliveryDate { get; set; }
        public string BuyerGLN { get; set; }
        public string BuyerVATNumber { get; set; }
        public string SupplierGLN { get; set; }
        public string SupplierVATNumber { get; set; }
        public List<Article> Articles { get; set; }
        
        public class Article
        {
            public string LineNumber { get; set; }
            public string ArticleDescription { get; set; }
            public string GTIN { get; set; }
            public string ArticleNetPrice { get; set; }
            public string OrderedQuantity { get; set; }
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
                /// By not using this operator this will not fill the document and crash the program.
                /// This method is used to make sure the important values are read from the xml files and optional values will be set to ""
                /// </summary>
                List<OrderDocument> OrderMsgList = _xMessages.Elements().Where(x => x.Element("MessageType").Value == "3").Select(x =>
                    new OrderDocument()
                    {
                        MessageStandard = x.Element("MessageStandard").Value ?? "",
                        MessageType = x.Element("MessageType").Value,
                        Sender = x.Element("Sender").Value,
                        SenderGLN = x.Element("SenderGLN").Value,
                        RecipientGLN = x.Element("RecipientGLN").Value,
                        IsTestMessage = x.Element("IsTestMessage").Value ?? "",
                        OrderNumberBuyer = x.Element("OrderNumberBuyer").Value,
                        RequestedDeliveryDate = x.Element("RequestedDeliveryDate").Value,
                        BuyerGLN = x.Element("BuyerGLN").Value,
                        BuyerVATNumber = x.Element("BuyerVATNumber").Value,
                        SupplierGLN = x.Element("SupplierGLN").Value,
                        SupplierVATNumber = x.Element("SupplierVATNumber").Value,
                        Articles = x.Elements("Article").Select(a => new Article()
                        {
                            LineNumber = a.Element("LineNumber").Value ?? "",
                            ArticleDescription = a.Element("ArticleDescription").Value ?? "",
                            GTIN = a.Element("GTIN").Value,
                            ArticleNetPrice = a.Element("ArticleNetPrice").Value ?? "",
                            OrderedQuantity = a.Element("OrderedQuantity").Value
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
            string buyerMailAddress ="";
            string buyerMailBody = "";
            int buyerOrderDocumentCount = 0;
            ex = null;

            foreach (OrderDocument orderDocument in (List<OrderDocument>)_dataObject)
            {
                try
                {
                    //oOrd.CardName = orderDocument.Sender;
                    oRs.DoQuery(@"SELECT T0.""Address"", T1.""CardCode"", T1.""CardName"" FROM CRD1 T0 INNER JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" WHERE T0.""GlblLocNum"" = '" + orderDocument.SenderGLN + "'");
                    if (oRs.RecordCount > 0)
                    {
                        if(oRs.Fields.Item(0).Size > 0)
                            oOrd.PayToCode = oRs.Fields.Item(0).Value.ToString();
                        else
                        {
                            EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error Pay To Address not found! With GlblLocNum: " + orderDocument.SenderGLN);
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error Pay To Address not found! With GlblLocNum: " + orderDocument.SenderGLN, "Error!");
                        }
                    }
                    else
                    {
                        EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error Pay To GlblLocNum: " + orderDocument.SenderGLN + " not found!");
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error Pay To GlblLocNum: " + orderDocument.SenderGLN + " not found!", "Error!");
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

                        if(fieldNotFound.Length > 0)
                        {
                            EventLogger.getInstance().EventError("Server: " + _connectedServer + " " + fieldNotFound);
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, fieldNotFound, "Error!");
                        }
                    }
                    else
                    {
                        EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error Ship To GlblLocNum: " + orderDocument.BuyerGLN + " not found!");
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error Ship To GlblLocNum: " + orderDocument.SenderGLN + " not found!", "Error!");
                    }
                    foreach (Article article in orderDocument.Articles)
                    {
                        oRs.DoQuery(@"SELECT ""ItemCode"" FROM OITM WHERE ""CodeBars"" = '" + article.GTIN + "'");
                        if (oRs.RecordCount > 0)
                            oOrd.Lines.ItemCode = oRs.Fields.Item(0).Value.ToString();
                        else
                        {
                            EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error CodeBars: " + article.GTIN + " not found!");
                            EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error CodeBars: " + article.GTIN + " not found!", "Error!");
                        }

                        oOrd.Lines.UserFields.Fields.Item("U_EdiLineNumber").Value = article.LineNumber;
                        oOrd.Lines.ItemDescription = article.ArticleDescription;
                        oOrd.Lines.Quantity = Convert.ToDouble(article.OrderedQuantity);

                        oOrd.Lines.Add();
                    }
                    oOrd.NumAtCard = orderDocument.OrderNumberBuyer;
                    oOrd.DocDueDate = Convert.ToDateTime(orderDocument.RequestedDeliveryDate);
                    oOrd.UserFields.Fields.Item("U_IsTestMessage").Value = orderDocument.IsTestMessage;

                    if(oOrd.Add() == 0)
                    {
                        string serviceCallID = ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetNewObjectKey();
                        oOrd.GetByKey(Convert.ToInt32(serviceCallID));
                        buyerOrderDocumentCount++;
                        buyerMailBody += buyerOrderDocumentCount + " - New Sales Order created with DocNum: " + oOrd.DocNum + System.Environment.NewLine;
                        EventLogger.getInstance().EventInfo("Server: " + _connectedServer + " Succesfully created Sales Order: " + oOrd.DocNum);
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Succesfully created Sales Order: " + oOrd.DocNum, "Processing..");
                    }
                    else
                    {
                        ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                        EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error creating Sales Order: (" + errCode + ") " + errMsg);
                        EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error creating Sales Order: (" + errCode + ") " + errMsg, "Error!");
                    }
                }
                catch (Exception e)
                {
                    ex = e;
                    EventLogger.getInstance().EventError("Server: " + _connectedServer + " Error saving to SAP: " + e.Message + " with order document: " + orderDocument);
                    EventLogger.getInstance().UpdateSAPLogMessage(_connectedServer, EdiConnectorData.getInstance().sRecordReference, "Error saving to SAP: " + e.Message + " with order document: " + orderDocument, "Error!");
                }
            }
            ConnectionManager.getInstance().GetConnection(_connectedServer).SendMailNotification("New sales order(s) created:" + buyerOrderDocumentCount, buyerMailBody, buyerMailAddress);

            EdiConnectorService.ClearObject(oOrd);
            EdiConnectorService.ClearObject(oRs);
        }
    }

}
