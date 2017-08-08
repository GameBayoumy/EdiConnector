using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace EdiConnectorService_C_Sharp
{
    class InvoiceDocument : IEdiDocumentType
    {

        // Data structure
        public string MessageStandard { get; set;}
        public string MessageType { get; set; }
        public string Sender { get; set; }
        public string SenderGLN { get; set; }
        public string RecipientGLN { get; set; }
        public string IsTestMessage { get; set; }
        public string InvoiceNumber { get; set; }
        public string InvoiceDate { get; set; }
        public string InvoiceType { get; set; }
        public string CurrencyCode { get; set; }
        public string DespatchAdviceNumber { get; set; }
        public string IsDutyFree { get; set; }
        public string BuyerGLN { get; set; }
        public string BuyerVATNumber { get; set; }
        public string SupplierGLN { get; set; }
        public string SupplierVATNumber { get; set; }
        public string InvoiceeGLN { get; set; }
        public string DeliveryPartyGLN { get; set; }
        public List<Article> Articles { get; set; }
        public InvoiceTotal InvoiceTotals { get; set; }
        
        public class Article
        {
            public string LineNumber { get; set; }
            public string ArticleDescription { get; set; }
            public string GTIN { get; set; }
            public string ArticleNetPrice { get; set; }
            public string ArticleNetPriceUnitCode { get; set; }
            public string InvoicedQuantity { get; set; }
            public string DeliveredQuantity { get; set; }
        }

        public class InvoiceTotal
        {
            public string InvoiceAmount { get; set; }
            public string NetLineAmount { get; set; }
            public string VATAmount { get; set; }
            public string DiscountAmount { get; set; }
            public InvoiceVATTotal InvoiceVATTotals { get; set; }
        }

        public class InvoiceVATTotal
        {
            public string IsDutyFree { get; set; }
            public string VATPercentage { get; set; }
            public string VATAmount { get; set; }
            public string VATBaseAmount { get; set; }
        }

        /// <summary>
        /// Reads the XML data.
        /// </summary>
        /// <param name="_xMessages">The incoming XML messages.</param>
        /// <returns>List of InvoiceDocuments</returns>
        public Object ReadXMLData(XElement _xMessages, out Exception _ex)
        {
            // Checks if the MessageType is for a Invoice Document.
            // Then it will create new InvoiceDocuments for every message
            _ex = null;
            try
            {
                List<InvoiceDocument>InvoiceMsgList = _xMessages.Elements().Where(x => x.Element("MessageType").Value == "8").Select(x =>
                new InvoiceDocument()
                {
                    MessageStandard = x.Element("MessageStandard").Value,
                    MessageType = x.Element("MessageType").Value,
                    Sender = x.Element("Sender").Value,
                    SenderGLN = x.Element("SenderGLN").Value,
                    RecipientGLN = x.Element("RecipientGLN").Value,
                    IsTestMessage = x.Element("IsTestMessage").Value,
                    InvoiceNumber = x.Element("InvoiceNumber").Value,
                    InvoiceDate = x.Element("InvoiceDate").Value,
                    InvoiceType = x.Element("InvoiceType").Value,
                    CurrencyCode = x.Element("CurrencyCode").Value,
                    DespatchAdviceNumber = x.Element("DespatchAdviceNumber").Value,
                    IsDutyFree = x.Element("IsDutyFree").Value,
                    BuyerGLN = x.Element("BuyerGLN").Value,
                    BuyerVATNumber = x.Element("BuyerVATNumber").Value,
                    SupplierGLN = x.Element("SupplierGLN").Value,
                    SupplierVATNumber = x.Element("SupplierVATNumber").Value,
                    InvoiceeGLN = x.Element("InvoiceeGLN").Value,
                    DeliveryPartyGLN = x.Element("DeliveryPartyGLN").Value,
                    Articles = x.Elements("Article").Select(a => new Article()
                    {
                        LineNumber = a.Element("LineNumber").Value,
                        ArticleDescription = a.Element("ArticleDescription").Value,
                        GTIN = a.Element("GTIN").Value,
                        ArticleNetPrice = a.Element("ArticleNetPrice").Value,
                        ArticleNetPriceUnitCode = a.Element("ArticleNetPriceUnitCode").Value,
                        InvoicedQuantity = a.Element("InvoicedQuantity").Value,
                        DeliveredQuantity = a.Element("DeliveredQuantity").Value
                    }).ToList(),
                    InvoiceTotals = x.Elements("InvoiceTotals").Select(ia => new InvoiceTotal()
                    {
                        InvoiceAmount = ia.Element("InvoiceAmount").Value,
                        NetLineAmount = ia.Element("NetLineAmount").Value,
                        VATAmount = ia.Element("VATAmount").Value,
                        DiscountAmount = ia.Element("DiscountAmount").Value,
                        InvoiceVATTotals = ia.Elements("InvoiceVATTotals").Select(ivat => new InvoiceVATTotal()
                        {
                            IsDutyFree = ivat.Element("IsDutyFree").Value,
                            VATPercentage = ivat.Element("VATPercentage").Value,
                            VATAmount = ivat.Element("VATAmount").Value,
                            VATBaseAmount = ivat.Element("VATBaseAmount").Value
                        }).First()
                    }).First()
                }).ToList();
                return InvoiceMsgList;
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
            catch (Exception e)
            {
                _ex = e;
            }
        }
    }

}
