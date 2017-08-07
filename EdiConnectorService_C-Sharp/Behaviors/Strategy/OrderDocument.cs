using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class OrderDocument : IEdiDocumentType
    {

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
        /// <returns>List of OrderDocuments</returns>
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
        public void SaveToSAP(Object _dataObject, string _connectedServer, out Exception ex)
        {
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            SAPbobsCOM.Documents oOrd = (SAPbobsCOM.Documents)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders));
            string buyerMailAddress ="";
            ex = null;

            foreach (OrderDocument orderDocument in (List<OrderDocument>)_dataObject)
            {
                try
                {
                    //oOrd.CardName = orderDocument.Sender;
                    oRs.DoQuery(@"SELECT T0.""Address"", T1.""CardCode"", T1.""CardName"" FROM CRD1 T0 INNER JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" WHERE T0.""GlblLocNum"" = '" + orderDocument.SenderGLN + "'");
                    if (oRs.RecordCount > 0)
                    {
                        oOrd.PayToCode = oRs.Fields.Item(0).Value.ToString();
                    }
                    oRs.DoQuery(@"SELECT T2.""Address"", T0.""CardCode"", T0.""CardName"", T1.""E_MailL"" FROM OCRD T0 INNER JOIN OCPR T1 ON T0.""CardCode"" = T1.""CardCode"" INNER JOIN CRD1 T2 ON T0.""CardCode"" = T2.""CardCode"" WHERE T2.""GlblLocNum"" = '" + orderDocument.BuyerGLN + "'");
                    if (oRs.RecordCount > 0)
                    {
                        oOrd.ShipToCode = oRs.Fields.Item(0).Value.ToString();
                        oOrd.CardCode = oRs.Fields.Item(1).Value.ToString();
                        oOrd.CardName = oRs.Fields.Item(2).Value.ToString();
                        buyerMailAddress = oRs.Fields.Item(3).Value.ToString();
                    }
                    foreach (Article article in orderDocument.Articles)
                    {
                        oRs.DoQuery(@"SELECT ""ItemCode"" FROM OITM WHERE ""CodeBars"" = '" + article.GTIN + "'");
                        if (oRs.RecordCount > 0)
                            oOrd.Lines.ItemCode = oRs.Fields.Item(0).Value.ToString();
                        else
                            EventLogger.getInstance().EventError("Error CodeBars:" + article.GTIN + " not found!");

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
                        ConnectionManager.getInstance().GetConnection(_connectedServer).SendMailNotification("New sales order created:" + oOrd.DocNum, "", buyerMailAddress);
                        EventLogger.getInstance().EventInfo("Succesfully added Sales Order: " + oOrd.DocNum);
                    }
                    else
                    {
                        ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetLastError(out var errCode, out var errMsg);
                        EventLogger.getInstance().EventError("Error adding Sales Order: (" + errCode + ") " + errMsg);
                    }
                }
                catch (Exception e)
                {
                    ex = e;
                    EventLogger.getInstance().EventError("Error saving to SAP: " + e.Message + " with order document: " + orderDocument);
                }

                EdiConnectorService.ClearObject(oOrd);
                EdiConnectorService.ClearObject(oRs);
            }
        }
    }

}
