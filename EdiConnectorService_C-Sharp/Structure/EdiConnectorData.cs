using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace EdiConnectorService_C_Sharp
{
    public class EdiConnectorData
    {
        private static EdiConnectorData instance = null;

        public static EdiConnectorData getInstance()
        {
            if (instance == null)
            {
                instance = new EdiConnectorData();
            }
            return instance;
        }

        public string sApplicationPath;
        public string sProcessedDirName;

        public SAPbobsCOM.Documents oOrder;
        public string CARDCODE;
        public string ITEMCODE;
        public string ITEMNAME;
        public string SALPACKUN;
        public string KOPERADRES;

        //Email notifications
        public int iSendNotification;
        public string sSmpt;
        public int iSmtpPort;
        public bool bSmtpUserSecurity;
        public string sSmtpUser;
        public string sSmtpPassword;

        public string sSenderEmail;
        public string sSenderName;

        public string sOrderMailTo;
        public string sOrderMailToFullName;

        public string sDeliveryMailTo;
        public string sDeliveryMailToFullName;

        public string sInvoiceMailTo;
        public string sInvoiceMailToFullName;

        //Connections
        public SqlConnection cn = new SqlConnection();
        public SAPbobsCOM.Company cmp = new SAPbobsCOM.Company();
        public SAPbobsCOM.BoDataServerTypes bstDBServerType;
        public string sSQL;
        public string sDBServerType;
        public string sServer;
        public string sDBUserName;
        public string sDBPassword;
        public string sCompanyDB;
        public string sUserName;
        public string sPassword;
        public string sSQLVersion;

        //Level
        public string sDesAdvLevel;

        //Files
        public string SO_FILE;
        public string SO_FILENAME;

        //Paths
        public string sSOPath;
        public string sSOTempPath;
        public string sSODonePath;
        public string sSOErrorPath;
        public string sInvoicePath;
        public string sDeliveryPath;

        //Invoice Dimension
        public string FK_K_EANCODE;
        public string FK_FAKTEST;
        public string FK_KNAAM;
        public string FK_F_SOORT;
        public string FK_FAKT_NUM;
        public string FK_FAKT_DATUM;
        public string FK_AFL_DATUM;
        public string FK_RFFIV;
        public string FK_K_ORDERNR;
        public string FK_K_ORDDAT;
        public string FK_PAKBONNR;
        public string FK_RFFCDN;
        public string FK_RFFALO;
        public string FK_RFFVN;
        public string FK_RFFVNDAT;
        public string FK_NAD_BY;
        public string FK_A_EANCODE;
        public string FK_F_EANCODE;
        public string FK_NAD_SF;
        public string FK_NAD_SU;
        public string FK_NAD_UC;
        public string FK_NAD_PE;
        public string FK_OBNUMMER;
        public string FK_ACT;
        public string FK_CUX;
        public string FK_DAGEN;
        public string FK_KORTPERC;
        public string FK_KORTBEDR;
        public string FK_ONTVANGER;

        public string AK_SOORT;
        public string AK_QUAL;
        public string AK_BEDRAG;
        public string AK_BTWSOORT;
        public string AK_FOOTMOA;
        public string AK_NOTINCALC;

        public string FR_DEUAC;
        public string FR_DEARTNR;
        public string FR_DEARTOM;
        public string FR_AANTAL;
        public string FR_FAANTAL;
        public string FR_ARTEENHEID;
        public string FR_FEENHEID;
        public string FR_NETTOBEDR;
        public string FR_PRIJS;
        public string FR_FREKEN;
        public string FR_BTWSOORT;
        public string FR_PV;
        public string FR_ORDER;
        public string FR_REGELID;
        public string FR_INVO;
        public string FR_DESA;
        public string FR_PRIAAA;
        public string FR_PIAPB;

        public string AR_SOORT;
        public string AR_QUAL;
        public string AR_BEDRAG;
        public string AR_BTWSOORT;
        public string AR_PERC;
        public string AR_LTOTAL;
        public string AR_FOOTMOA;
        public string AR_NOTINCALC;

        //Delivery
        public string DK_K_EANCODE;
        public string DK_PBTEST;
        public string DK_KNAAM;
        public string DK_BGM;
        public string DK_DTM_137;
        public string DK_DTM_2;
        public string DK_TIJD_2;
        public string DK_DTM_17;
        public string DK_TIJD_17;
        public string DK_DTM_64;
        public string DK_TIJD_64;
        public string DK_DTM_63;
        public string DK_TIJD_63;
        public string DK_BH_DAT;
        public string DK_BH_TIJD;
        public string DK_RFF;
        public string DK_RFFVN;
        public string DK_BH_EAN;
        public string DK_NAD_BY;
        public string DK_NAD_DP;
        public string DK_NAD_SU;
        public string DK_NAD_UC;
        public string DK_DESATYPE;
        public string DK_ONTVANGER;

        public string DR_DEUAC;
        public string DR_OLDUAC;
        public string DR_DEARTNR;
        public string DR_DEARTOM;
        public string DR_PIA;
        public string DR_BATCH;
        public string DR_QTY;
        public string DR_ARTEENHEID;
        public string DR_RFFONID;
        public string DR_RFFONORD;
        public string DR_DTM_23E;
        public string DR_TGTDATUM;
        public string DR_GEWICHT;
        public string DR_FEENHEID;
        public string DR_QTY_AFW;
        public string DR_REDEN;
        public string DR_GINTYPE;
        public string DR_GINID;
        public string DR_BATCHH;

        //Sales Order
        public string OK_K_EANCODE;
        public string OK_TEST;
        public string OK_KNAAM;
        public string OK_BGM;
        public DateTime OK_K_ORDDAT;
        public DateTime OK_DTM_2;
        public string OK_TIJD_2;
        public DateTime OK_DTM_17;
        public string OK_TIJD_17;
        public DateTime OK_DTM_64;
        public string OK_TIJD_64;
        public DateTime OK_DTM_63;
        public string OK_TIJD_63;
        public string OK_RFF_BO;
        public string OK_RFF_CR;
        public string OK_RFF_PD;
        public string OK_RFFCT;
        public DateTime OK_DTMCT;

        public string[] OK_FLAGS = new string[10];

        public string OK_FTXDSI;
        public string OK_NAD_BY;
        public string OK_NAD_DP;
        public string OK_NAD_IV;
        public string OK_NAD_SF;
        public string OK_NAD_SU;
        public string OK_NAD_UC;
        public string OK_NAD_BCO;
        public string OK_RECEIVER;

        public string OR_DEUAC;
        public double OR_QTY;
        public string OR_LEVARTCODE;
        public string OR_DEARTOM;
        public string OR_COLOR;
        public string OR_LENGTH;
        public string OR_WIDTH;
        public string OR_HEIGHT;
        public string OR_CUX;
        public string OR_PIA;
        public string OR_RFFLI1;
        public string OR_RFFLI2;
        public DateTime OR_DTM_2;
        public string OR_LINNR;
        public double OR_PRI;

        //order file pos/len
        public int[] OK_POS = new int[37];
        public int[] OK_LEN = new int[37];

        public int[] OR_POS = new int[14];
        public int[] OR_LEN = new int[14];


    }
}
