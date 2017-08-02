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
        public Object ReadXMLData(XElement _xMessages)
        {
            // Checks if the MessageType is for a Order Response Document.
            // Then it will create new OrderDocuments for every message
             List<OrderDocument>OrderMsgList = _xMessages.Elements().Where(x => x.Element("MessageType").Value == "3").Select(x =>
                new OrderDocument()
                {
                    MessageStandard = x.Element("MessageStandard").Value,
                    MessageType = x.Element("MessageType").Value,
                    Sender = x.Element("Sender").Value,
                    SenderGLN = x.Element("SenderGLN").Value,
                    RecipientGLN = x.Element("RecipientGLN").Value,
                    IsTestMessage = x.Element("IsTestMessage").Value,
                    OrderNumberBuyer = x.Element("OrderNumberBuyer").Value,
                    BuyerGLN = x.Element("BuyerGLN").Value,
                    BuyerVATNumber = x.Element("BuyerVATNumber").Value,
                    SupplierGLN = x.Element("SupplierGLN").Value,
                    SupplierVATNumber = x.Element("SupplierVATNumber").Value,
                    Articles = x.Elements("Article").Select(a => new Article()
                    {
                        LineNumber = a.Element("LineNumber").Value,
                        ArticleDescription = a.Element("ArticleDescription").Value,
                        GTIN = a.Element("GTIN").Value,
                        ArticleNetPrice = a.Element("ArticleNetPrice").Value,
                        OrderedQuantity = a.Element("OrderedQuantity").Value
                    }).ToList()
                }).ToList();
             return OrderMsgList;
        }

        /// <summary>
        /// Saves specific document data object to SAP.
        /// </summary>
        /// <param name="_dataObject">The data object.</param>
        public void SaveToSAP(Object _dataObject)
        {
            foreach (OrderDocument orderDocument in (List<OrderDocument>)_dataObject)
            {
                SAPbobsCOM.Documents oInv;
                oInv = (SAPbobsCOM.Documents)(ConnectionManager.getInstance().GetConnection(_connectedServer).Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                
                try
                {

                }
                catch (Exception e)
                {
                    EventLogger.getInstance().EventError("Error saving to SAP: " + e.Message + " with order document: " + orderDocument);
                }
            }
        }
    }

}
