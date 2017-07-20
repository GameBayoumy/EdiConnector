Imports System.Data.SqlClient
Imports SAPbobsCOM

Module mod_var

    'connections
    Public cn As New SqlConnection
    Public cmp As New Company
    Public bstDBServerType As BoDataServerTypes
    Public strSQL As String
    Public DBServerType As String
    Public strServer As String
    Public strDbUserName As String
    Public strDbPassword As String
    Public strCompanyDB As String
    Public strUserName As String
    Public strPassword As String
    Public strSQLVersion As String

    'Level
    Public strDesAdvLevel As String

    'Files
    Public SO_FILE As String = ""
    Public SO_FILENAME As String = ""

    'Paths
    Public strSoPath As String
    Public strSoTempPath As String
    Public strSoDonePath As String
    Public strSoErrorPath As String
    Public strInvoicePath As String
    Public strDeliveryPath As String

    'Factuur Dimensio
    Public FK_K_EANCODE As String
    Public FK_FAKTEST As String
    Public FK_KNAAM As String
    Public FK_F_SOORT As String
    Public FK_FAKT_NUM As String
    Public FK_FAKT_DATUM As String
    Public FK_AFL_DATUM As String
    Public FK_RFFIV As String
    Public FK_K_ORDERNR As String
    Public FK_K_ORDDAT As String
    Public FK_PAKBONNR As String
    Public FK_RFFCDN As String
    Public FK_RFFALO As String
    Public FK_RFFVN As String
    Public FK_RFFVNDAT As String
    Public FK_NAD_BY As String
    Public FK_A_EANCODE As String
    Public FK_F_EANCODE As String
    Public FK_NAD_SF As String
    Public FK_NAD_SU As String
    Public FK_NAD_UC As String
    Public FK_NAD_PE As String
    Public FK_OBNUMMER As String
    Public FK_ACT As String
    Public FK_CUX As String
    Public FK_DAGEN As String
    Public FK_KORTPERC As String
    Public FK_KORTBEDR As String
    Public FK_ONTVANGER As String

    Public AK_SOORT As String
    Public AK_QUAL As String
    Public AK_BEDRAG As String
    Public AK_BTWSOORT As String
    Public AK_FOOTMOA As String
    Public AK_NOTINCALC As String

    Public FR_DEUAC As String
    Public FR_DEARTNR As String
    Public FR_DEARTOM As String
    Public FR_AANTAL As String
    Public FR_FAANTAL As String
    Public FR_ARTEENHEID As String
    Public FR_FEENHEID As String
    Public FR_NETTOBEDR As String
    Public FR_PRIJS As String
    Public FR_FREKEN As String
    Public FR_BTWSOORT As String
    Public FR_PV As String
    Public FR_ORDER As String
    Public FR_REGELID As String
    Public FR_INVO As String
    Public FR_DESA As String
    Public FR_PRIAAA As String
    Public FR_PIAPB As String

    Public AR_SOORT As String
    Public AR_QUAL As String
    Public AR_BEDRAG As String
    Public AR_BTWSOORT As String
    Public AR_PERC As String
    Public AR_LTOTAL As String
    Public AR_FOOTMOA As String
    Public AR_NOTINCALC As String


    'Delivery
    Public DK_K_EANCODE As String
    Public DK_PBTEST As String
    Public DK_KNAAM As String
    Public DK_BGM As String
    Public DK_DTM_137 As String
    Public DK_DTM_2 As String
    Public DK_TIJD_2 As String
    Public DK_DTM_17 As String
    Public DK_TIJD_17 As String
    Public DK_DTM_64 As String
    Public DK_TIJD_64 As String
    Public DK_DTM_63 As String
    Public DK_TIJD_63 As String
    Public DK_BH_DAT As String
    Public DK_BH_TIJD As String
    Public DK_RFF As String
    Public DK_RFFVN As String
    Public DK_BH_EAN As String
    Public DK_NAD_BY As String
    Public DK_NAD_DP As String
    Public DK_NAD_SU As String
    Public DK_NAD_UC As String
    Public DK_DESATYPE As String
    Public DK_ONTVANGER As String

    Public DR_DEUAC As String
    Public DR_OLDUAC As String
    Public DR_DEARTNR As String
    Public DR_DEARTOM As String
    Public DR_PIA As String
    Public DR_BATCH As String
    Public DR_QTY As String
    Public DR_ARTEENHEID As String
    Public DR_RFFONID As String
    Public DR_RFFONORD As String
    Public DR_DTM_23E As String
    Public DR_TGTDATUM As String
    Public DR_GEWICHT As String
    Public DR_FEENHEID As String
    Public DR_QTY_AFW As String
    Public DR_REDEN As String
    Public DR_GINTYPE As String
    Public DR_GINID As String
    Public DR_BATCHH As String


    'Sales Order
    Public OK_K_EANCODE As String
    Public OK_TEST As String
    Public OK_KNAAM As String
    Public OK_BGM As String
    Public OK_K_ORDDAT As Date
    Public OK_DTM_2 As Date
    Public OK_TIJD_2 As String
    Public OK_DTM_17 As Date
    Public OK_TIJD_17 As String
    Public OK_DTM_64 As Date
    Public OK_TIJD_64 As String
    Public OK_DTM_63 As Date
    Public OK_TIJD_63 As String
    Public OK_RFF_BO As String
    Public OK_RFF_CR As String
    Public OK_RFF_PD As String
    Public OK_RFFCT As String
    Public OK_DTMCT As Date

    Public OK_FLAGS(10) As String

    Public OK_FTXDSI As String
    Public OK_NAD_BY As String
    Public OK_NAD_DP As String
    Public OK_NAD_IV As String
    Public OK_NAD_SF As String
    Public OK_NAD_SU As String
    Public OK_NAD_UC As String
    Public OK_NAD_BCO As String
    Public OK_ONTVANGER As String

    Public OR_DEUAC As String
    Public OR_QTY As Double
    Public OR_LEVARTCODE As String
    Public OR_DEARTOM As String
    Public OR_KLEUR As String
    Public OR_LENGTE As String
    Public OR_BREEDTE As String
    Public OR_HOOGTE As String
    Public OR_CUX As String
    Public OR_PIA As String
    Public OR_RFFLI1 As String
    Public OR_RFFLI2 As String
    Public OR_DTM_2 As Date
    Public OR_LINNR As String
    Public OR_PRI As Double

    'order file pos/len
    Public OK_POS(37) As Integer
    Public OK_LEN(37) As Integer

    Public OR_POS(14) As Integer
    Public OR_LEN(14) As Integer


End Module
