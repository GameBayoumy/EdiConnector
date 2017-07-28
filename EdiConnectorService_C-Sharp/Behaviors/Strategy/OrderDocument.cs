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

        public Object ReadXMLData(XElement _xMessages)
        {
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
                        OrderedQuantity = a.Element("OrderedQuantity").Value,
                    }).ToList()
                }).ToList();
             return OrderMsgList;
        }
    }

}
