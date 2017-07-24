using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace EdiConnectorService_C_Sharp
{
    class EdiConnectorData
    {
        private string sApplicationPath;

        //Dit is de huidige versie aangepast op 15/09/2014
        private SAPbobsCOM.Documents oOrder;
        private string CARDCODE;
        private string ITEMCODE;
        private string ITEMNAME;
        private string SALPACKUN;
        private string KOPERADRES;

        //Email notifications
        private int iSendNotification;
        private string sSmpt;
        private int iSmtpPort;
        private string sSmtpUser;
        private string sSmtpPassword;

        private string sSenderEmail;
        private string sSenderName;

        private string sOrderMailTo;
        private string sOrderMailToFullName;

        private string sDeliveryMailTo;
        private string sDeliveryMailToFullName;

        private string sInvoiceMailTo;
        private string sInvoiceMailToFullName;

        //Connections
        private SqlConnection cn;
        private SAPbobsCOM.Company cmp;
        private SAPbobsCOM.BoDataServerTypes bstDBServerType;
        private string sSQL;
        private string sDBServerType;
        private string sServer;
        private string sDBUsername;
        private string sDBPassword;
        private string sUserName;
        private string sPassword;
        private string sSQLVersion;

        //Level
        private string sDesAdvLevel;

        //Files
        private string SO_FILE;
        private string SO_FILENAME;

        //Paths
        private string sSOPath;
        private string sSOTempPath;
        private string sSODonePath;
        private string sSOErrorPath;
        private string sInvoicePath;
        private string sDeliveryPath;

        //Invoice Dimension
        private string FK_K_EANCODE;
        private string FK_FAKTEST;
        private string FK_KNAAM;
        private string FK_F_SOORT;
        private string FK_FAKT_NUM;
        private string FK_FAKT_DATUM;
        private string FK_AFL_DATUM;
        private string FK_RFFIV;
        private string FK_K_ORDERNR;
        private string FK_K_ORDDAT;
        private string FK_PAKBONNR;
        private string FK_RFFCDN;
        private string FK_RFFALO;
        private string FK_RFFVN;
        private string FK_RFFVNDAT;
        private string FK_NAD_BY;
        private string FK_A_EANCODE;
        private string FK_F_EANCODE;
        private string FK_NAD_SF;
        private string FK_NAD_SU;
        private string FK_NAD_UC;
        private string FK_NAD_PE;
        private string FK_OBNUMMER;
        private string FK_ACT;
        private string FK_CUX;
        private string FK_DAGEN;
        private string FK_KORTPERC;
        private string FK_KORTBEDR;
        private string FK_ONTVANGER;

        private string AK_SOORT;
        private string AK_QUAL;
        private string AK_BEDRAG;
        private string AK_BTWSOORT;
        private string AK_FOOTMOA;
        private string AK_NOTINCALC;

        private string FR_DEUAC;
        private string FR_DEARTNR;
        private string FR_DEARTOM;
        private string FR_AANTAL;
        private string FR_FAANTAL;
        private string FR_ARTEENHEID;
        private string FR_FEENHEID;
        private string FR_NETTOBEDR;
        private string FR_PRIJS;
        private string FR_FREKEN;
        private string FR_BTWSOORT;
        private string FR_PV;
        private string FR_ORDER;
        private string FR_REGELID;
        private string FR_INVO;
        private string FR_DESA;
        private string FR_PRIAAA;
        private string FR_PIAPB;

        private string AR_SOORT;
        private string AR_QUAL;
        private string AR_BEDRAG;
        private string AR_BTWSOORT;
        private string AR_PERC;
        private string AR_LTOTAL;
        private string AR_FOOTMOA;
        private string AR_NOTINCALC;

        //Delivery
        private string DK_K_EANCODE;
        private string DK_PBTEST;
        private string DK_KNAAM;
        private string DK_BGM;
        private string DK_DTM_137;
        private string DK_DTM_2;
        private string DK_TIJD_2;
        private string DK_DTM_17;
        private string DK_TIJD_17;
        private string DK_DTM_64;
        private string DK_TIJD_64;
        private string DK_DTM_63;
        private string DK_TIJD_63;
        private string DK_BH_DAT;
        private string DK_BH_TIJD;
        private string DK_RFF;
        private string DK_RFFVN;
        private string DK_BH_EAN;
        private string DK_NAD_BY;
        private string DK_NAD_DP;
        private string DK_NAD_SU;
        private string DK_NAD_UC;
        private string DK_DESATYPE;
        private string DK_ONTVANGER;

        private string DR_DEUAC;
        private string DR_OLDUAC;
        private string DR_DEARTNR;
        private string DR_DEARTOM;
        private string DR_PIA;
        private string DR_BATCH;
        private string DR_QTY;
        private string DR_ARTEENHEID;
        private string DR_RFFONID;
        private string DR_RFFONORD;
        private string DR_DTM_23E;
        private string DR_TGTDATUM;
        private string DR_GEWICHT;
        private string DR_FEENHEID;
        private string DR_QTY_AFW;
        private string DR_REDEN;
        private string DR_GINTYPE;
        private string DR_GINID;
        private string DR_BATCHH;

        //Sales Order
        private string OK_K_EANCODE;
        private string OK_TEST;
        private string OK_KNAAM;
        private string OK_BGM;
        private DateTime OK_K_ORDDAT;
        private DateTime OK_DTM_2;
        private string OK_TIJD_2;
        private DateTime OK_DTM_17;
        private string OK_TIJD_17;
        private DateTime OK_DTM_64;
        private string OK_TIJD_64;
        private DateTime OK_DTM_63;
        private string OK_TIJD_63;
        private string OK_RFF_BO;
        private string OK_RFF_CR;
        private string OK_RFF_PD;
        private string OK_RFFCT;
        private DateTime OK_DTMCT;

        private string[] OK_FLAGS = new string[10];

        private string OK_FTXDSI;
        private string OK_NAD_BY;
        private string OK_NAD_DP;
        private string OK_NAD_IV;
        private string OK_NAD_SF;
        private string OK_NAD_SU;
        private string OK_NAD_UC;
        private string OK_NAD_BCO;
        private string OK_ONTVANGER;

        private string OR_DEUAC;
        private double OR_QTY;
        private string OR_LEVARTCODE;
        private string OR_DEARTOM;
        private string OR_KLEUR;
        private string OR_LENGTE;
        private string OR_BREEDTE;
        private string OR_HOOGTE;
        private string OR_CUX;
        private string OR_PIA;
        private string OR_RFFLI1;
        private string OR_RFFLI2;
        private DateTime OR_DTM_2;
        private string OR_LINNR;
        private double OR_PRI;

        //order file pos/len
        private int[] OK_POS = new int[37];
        private int[] OK_LEN = new int[37];

        private int[] OR_POS = new int[14];
        private int[] OR_LEN = new int[14];
    }
}
