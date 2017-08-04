using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class OrderResponseDocument : IEdiDocumentType
    {

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
        /// <returns>List of OrderResponseDocuments</returns>
        public Object ReadXMLData(XElement _xMessages, out Exception _ex)
        {
            // Checks if the MessageType is for a Order document.
            // Then it will create new OrderResponseDocuments for every message
            _ex = null;
            try
            {
                List<OrderResponseDocument>OrderMsgList = _xMessages.Elements().Where(x => x.Element("MessageType").Value == "5").Select(x =>
                new OrderResponseDocument()
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
                        OrderedQuantity = a.Element("OrderedQuantity").Value,
                    }).ToList()
                }).ToList();
                return OrderMsgList;
            }
            catch (Exception e)
            {
                _ex = e;
                return null;
            }
        }

        public void SaveToSAP(Object _dataObject, string _connectedServer, out Exception _ex)
        {
            _ex = null;
            try
            {

            }
            catch(Exception e)
            {
                _ex = e;
            }

        }
    }

}
