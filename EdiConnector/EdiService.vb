Imports System.Data.SqlClient
Imports SAPbobsCOM
Imports System.IO
Imports System.Net
Imports System.Threading

Public Class EdiConnectorService

    Private applicationPath As String = My.Application.Info.DirectoryPath
    Private stopping As Boolean
    Private stoppedEvent As ManualResetEvent

    '' Dit is de huidige versie aangepast op 15/09/2014

    Private oOrder As SAPbobsCOM.Documents

    Private CARDCODE As String
    Private ITEMCODE As String
    Private ITEMNAME As String
    Private SALPACKUN As String
    Private KOPERADRES As String

    'email notifications.
    Private iSendNotification As Integer

    Private sSmtp As String
    Private iSmtpPort As Integer
    Private bSmtpUserSecurity As Boolean
    Private sSmtpUser As String
    Private sSmtpPassword As String

    Private sSenderEmail As String
    Private sSenderName As String

    Private sOrderMailTo As String
    Private sOrderMailToFullName As String

    Private sDeliveryMailTo As String
    Private sDeliveryMailToFullName As String

    Private sInvoiceMailTo As String
    Private sInvoiceMailToFullName As String

    Public Sub New()
        InitializeComponent()

        Me.stopping = False
        Me.stoppedEvent = New ManualResetEvent(False)
    End Sub

    ''' <summary>
    ''' The function is executed when a Start command is sent to the service
    ''' by the SCM or when the operating system starts (for a service that 
    ''' starts automatically). It specifies actions to take when the service 
    ''' starts. In this code sample, OnStart logs a service-start message to 
    ''' the Application log, and queues the main service function for 
    ''' execution in a thread pool worker thread.
    ''' </summary>
    ''' <param name="args">Command line arguments</param>
    ''' <remarks>
    ''' A service application is designed to be long running. Therefore, it 
    ''' usually polls or monitors something in the system. The monitoring is 
    ''' set up in the OnStart method. However, OnStart does not actually do 
    ''' the monitoring. The OnStart method must return to the operating 
    ''' system after the service's operation has begun. It must not loop 
    ''' forever or block. To set up a simple monitoring mechanism, one 
    ''' general solution is to create a timer in OnStart. The timer would 
    ''' then raise events in your code periodically, at which time your 
    ''' service could do its monitoring. The other solution is to spawn a 
    ''' new thread to perform the main service functions, which is 
    ''' demonstrated in this code sample.
    ''' </remarks>
    Protected Overrides Sub OnStart(ByVal args() As String)

        ' Log a service start message to the Application log.
        Me.EventLog1.WriteEntry("EdiService in OnStart.")

        ' Queue the main service function for execution in a worker thread.
        ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf ServiceWorkerThread))

        ReadSettings()
        If Connect_to_Sap() = True Then
            CreateUdfFiledsText()
            CheckAndExport_delivery()
            CheckAndExport_invoice()
            Split_Order()
            Read_SO_file()


        End If

    End Sub

    ''' <summary>
    ''' The method performs the main function of the service. It runs on a 
    ''' thread pool worker thread.
    ''' </summary>
    ''' <param name="state"></param>
    Private Sub ServiceWorkerThread(ByVal state As Object)
        ' Periodically check if the service is stopping.
        Do While Not Me.stopping
            ' Perform main service function here...
            Me.EventLog1.WriteEntry("EdiService Tick.")

            Thread.Sleep(2000)  ' Simulate some lengthy operations.
        Loop

        ' Signal the stopped event.
        Me.stoppedEvent.Set()
    End Sub

    ''' <summary>
    ''' The function is executed when a Stop command is sent to the service 
    ''' by SCM. It specifies actions to take when a service stops running. In 
    ''' this code sample, OnStop logs a service-stop message to the 
    ''' Application log, and waits for the finish of the main service 
    ''' function.
    ''' </summary>
    Protected Overrides Sub OnStop()

        ' Add code here to perform any tear-down necessary to stop your service.
        Disconnect_to_Sap()

        ' Log a service stop message to the Application log.
        Me.EventLog1.WriteEntry("EdiService in OnStop.")

        ' Indicate that the service is stopping and wait for the finish of 
        ' the main service function (ServiceWorkerThread).
        Me.stopping = True
        Me.stoppedEvent.WaitOne()

    End Sub

    Public Function Connect_to_Sap() As Boolean

        Try

            If cmp.Connected = True Then
                Call Log("X", "SAP is already Conneted.", "Connect_to_Sap")
                Exit Try
            End If

            Dim ret As Long = 0

            Dim bln As Boolean
            bln = Connect_to_database()

            If bln = False Then
                Return False
                Exit Try
            End If

            With cmp
                .DbServerType = bstDBServerType
                .Server = strServer
                .DbUserName = strDbUserName
                .DbPassword = strDbPassword
                .CompanyDB = strCompanyDB
                .UserName = strUserName
                .Password = strPassword
                .UseTrusted = False
                ret = .Connect()

            End With

            If ret <> 0 Then

                Dim nErr As Long
                Dim errMsg As String = ""
                ''MessageBox.Show(nErr.ToString())
                cmp.GetLastError(nErr, errMsg)
                If (nErr <> 0) Then
                    Call Log("X", "Error: " & errMsg & "(" & Str(nErr) & ")", "Connect_to_Sap")
                End If

            Else
                Return True

            End If


        Catch ex As Exception

            Call Log("X", ex.Message, "Connect_to_Sap")

        End Try

        Return False

    End Function

    Private Function Connect_to_database() As Boolean

        Try

            cn.ConnectionString = "Data Source=" & strServer & _
                      ";Initial Catalog=" & strCompanyDB & _
                      ";User ID=" & strDbUserName & _
                      ";Password=" & strDbPassword

            If strSQLVersion = "2005" Then
                bstDBServerType = BoDataServerTypes.dst_MSSQL2005
            ElseIf strSQLVersion = "2008" Then
                bstDBServerType = BoDataServerTypes.dst_MSSQL2008
            ElseIf strSQLVersion = "2012" Then
                bstDBServerType = BoDataServerTypes.dst_MSSQL2012
            Else
                bstDBServerType = BoDataServerTypes.dst_MSSQL
            End If

            cn.Open()

            Call Log("V", "Connected to database " & strCompanyDB & ".", "Connect_to_database")
            '' MessageBox.Show(strCompanyDB, "ok")
            '' If strCompanyDB <> "DimensioLive" And strCompanyDB <> "Demo_EDI_DimBasis" Then
            ''MessageBox.Show("De gekozen SAP Database heeft geen licentie. Neem voor een licentie contact op met SW Solutions", "OK")
            ''Environment.Exit(0)
            ''End If

            Return True

        Catch ex As Exception

            Call Log("X", ex.Message, "Connect_to_database")
            Return False


        End Try

    End Function

    Public Function ReadSettings() As Integer

        Try

            Dim ds As New DataSet
            ds.ReadXml(applicationPath & "\settings.xml")
            '' MessageBox.Show(applicationPath, "ok")
            If ds.Tables("server").Rows(0).Item("sqlversion") = "2005" Then
                bstDBServerType = BoDataServerTypes.dst_MSSQL2005
            ElseIf ds.Tables("server").Rows(0).Item("sqlversion") = "2008" Then
                bstDBServerType = BoDataServerTypes.dst_MSSQL2008
            ElseIf ds.Tables("server").Rows(0).Item("sqlversion") = "2012" Then
                bstDBServerType = BoDataServerTypes.dst_MSSQL2012
            Else
                bstDBServerType = BoDataServerTypes.dst_MSSQL
            End If

            strServer = ds.Tables("server").Rows(0).Item("name")
            strDbUserName = ds.Tables("server").Rows(0).Item("userid")
            strDbPassword = ds.Tables("server").Rows(0).Item("password")
            strCompanyDB = ds.Tables("server").Rows(0).Item("catalog")
            strUserName = ds.Tables("server").Rows(0).Item("sapuser")
            strPassword = ds.Tables("server").Rows(0).Item("sappassword")
            strSQLVersion = ds.Tables("server").Rows(0).Item("sqlversion")
            strDesAdvLevel = ds.Tables("server").Rows(0).Item("desadvlevel")

            strSoPath = ds.Tables("folders").Rows(0).Item("so")
            strSoTempPath = ds.Tables("folders").Rows(0).Item("sotemp")
            strSoDonePath = ds.Tables("folders").Rows(0).Item("sodone")
            strSoErrorPath = ds.Tables("folders").Rows(0).Item("soerror")
            strInvoicePath = ds.Tables("folders").Rows(0).Item("invoice")
            strDeliveryPath = ds.Tables("folders").Rows(0).Item("delivery")

            iSendNotification = ds.Tables("email").Rows(0).Item("send_notification")

            sSmtp = ds.Tables("email").Rows(0).Item("smtp")
            iSmtpPort = ds.Tables("email").Rows(0).Item("port")

            If ds.Tables("email").Rows(0).Item("security") = "0" Then
                bSmtpUserSecurity = False
            Else
                bSmtpUserSecurity = True
            End If

            sSmtpUser = ds.Tables("email").Rows(0).Item("user")
            sSmtpPassword = ds.Tables("email").Rows(0).Item("password")

            sSenderEmail = ds.Tables("email").Rows(0).Item("emailaddress")
            sSenderName = ds.Tables("email").Rows(0).Item("fullname")

            sOrderMailTo = ds.Tables("email").Rows(0).Item("emailaddress_order")
            sOrderMailToFullName = ds.Tables("email").Rows(0).Item("fullname_order")

            sDeliveryMailTo = ds.Tables("email").Rows(0).Item("emailaddress_delivery")
            sDeliveryMailToFullName = ds.Tables("email").Rows(0).Item("fullname_delivery")

            sInvoiceMailTo = ds.Tables("email").Rows(0).Item("emailaddress_invoice")
            sInvoiceMailToFullName = ds.Tables("email").Rows(0).Item("fullname_invoice")

            ds.Dispose()

            ReadInterface()

        Catch ex As Exception

            Call Log("X", "Settings are not in the right format! - " & ex.Message, "ReadSettings")

        End Try

        Return 0

    End Function

    Public Function ReadInterface() As Integer

        Dim ds As New DataSet
        ds.ReadXml(applicationPath & "\orderfld.xml")

        Dim o As Integer
        For o = 0 To ds.Tables("OK").Rows.Count - 1
            OK_POS(o) = ds.Tables("OK").Rows(o).Item(0)
            OK_LEN(o) = ds.Tables("OK").Rows(o).Item(1)
        Next

        Dim r As Integer
        For r = 0 To ds.Tables("OR").Rows.Count - 1
            OR_POS(r) = ds.Tables("OR").Rows(r).Item(0)
            OR_LEN(r) = ds.Tables("OR").Rows(r).Item(1)
        Next

        ds.Dispose()

        Return 0

    End Function

    Public Function Log(ByVal sType As String, ByVal msg As String, ByVal FunctionSender As String) As Integer

        Me.EventLog1.WriteEntry(sType & " - " & Format(Date.Now, "dd/MM/yyyy HH:mm:ss") & " - " & Replace(FunctionSender, "_", " ") & " - " & msg)

        Return 0

    End Function

    Public Function Disconnect_to_Sap() As Integer

        If cmp.Connected = True Then
            cmp.Disconnect()
            cn.Close()
            Call Log("V", "SAP is disconneted.", "Disconnect_to_Sap")
        Else
            Call Log("X", "SAP is already disconneted.", "Disconnect_to_Sap")
        End If

        Return 0

    End Function

    Public Function Split_Order() As Integer

        Dim strFileSize As String = ""
        Dim di As New IO.DirectoryInfo(strSoPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.DAT")
        Dim fi As IO.FileInfo

        Dim strNewFile As String = ""
        Dim strBGMnummer As String = ""

        For Each fi In aryFi

            strFileSize = (Format(fi.Length / 1024, "##0.00")).ToString()
            Call Log("V", "Begin splitting sales order filename " & fi.Name & " - file size: " & strFileSize & " kb", "Request_SO_file")

            Dim line As String

            Using reader As StreamReader = New StreamReader(fi.FullName)

                While True

                    line = reader.ReadLine
                    If line Is Nothing Then Exit While

                    If Mid(line, 1, 1) = "0" Then

                        If Len(strNewFile) > 0 Then
                            WriteNewFile(strNewFile, strBGMnummer)
                            strNewFile = ""
                        End If

                        strBGMnummer = Trim(Mid(line, 30, 35))

                    End If

                    strNewFile &= line & vbCrLf

                End While

                If Len(strNewFile) > 0 Then WriteNewFile(strNewFile, strBGMnummer)

            End Using

            Try
                File.Delete(fi.FullName)
                Call Log("V", "Splitted Sales order filename " & fi.Name & " deleted.", "Split_Order")
            Catch ex As Exception
                Call Log("X", ex.Message, "Split_Order")
            End Try

        Next

        Return 0

    End Function

    Private Function WriteNewFile(NEW_FILE As String, BGM_title As String) As Integer

        Try

            Using writer As New StreamWriter(strSoTempPath & "\" & "ORDER" & "_" & BGM_title & ".DAT")

                writer.Write(NEW_FILE)
                writer.Close()
                Call Log("V", "New splitted order created " & BGM_title, "WriteNewFile")

            End Using

        Catch ex As Exception

            Call Log("X", ex.Message, "WriteNewFile")

        End Try

        Return 0

    End Function

    Public Function Read_SO_file() As Integer

        Dim di As New IO.DirectoryInfo(strSoTempPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.DAT")
        Dim fi As IO.FileInfo

        For Each fi In aryFi

            Try
                Call Log("V", "Begin reading sales order filename " & fi.Name & ".", "Read_SO_file")

                Dim fs As New StreamReader(fi.FullName, False)
                SO_FILE = fs.ReadToEnd()
                SO_FILENAME = fi.Name
                fs.Close()

            Catch ex As Exception

                Call Log("X", ex.Message, "Read_SO_file")

            End Try

            Match_SO_data()

        Next

        Return 0

    End Function

    Public Function Match_SO_data() As Boolean

        Try
            Dim line = ""

            Using reader As New StringReader(SO_FILE)
                While True
                    line = reader.ReadLine
                    If line Is Nothing Then Exit While
                    Select Case Mid(line, 1, 1)
                        Case Is = 0
                            OK_K_EANCODE = Trim(Mid(line, OK_POS(0), OK_LEN(0)))
                            OK_TEST = Trim(Mid(line, OK_POS(1), OK_LEN(1)))
                            OK_KNAAM = Trim(Mid(line, OK_POS(2), OK_LEN(2)))
                            OK_BGM = Trim(Mid(line, OK_POS(3), OK_LEN(3)))
                            If Len(Trim(Mid(line, OK_POS(4), OK_LEN(4)))) = 8 Then
                                OK_K_ORDDAT = Mid(Mid(line, OK_POS(4), OK_LEN(4)), 7, 2) & "-" & Mid(Mid(line, OK_POS(4), OK_LEN(4)), 5, 2) & "-" & Mid(Mid(line, OK_POS(4), OK_LEN(4)), 1, 4)
                            Else
                                OK_K_ORDDAT = "01-01-0001"
                            End If
                            If Len(Trim(Mid(line, OK_POS(5), OK_LEN(5)))) = 8 Then
                                OK_DTM_2 = Mid(Mid(line, OK_POS(5), OK_LEN(5)), 7, 2) & "-" & Mid(Mid(line, OK_POS(5), OK_LEN(5)), 5, 2) & "-" & Mid(Mid(line, OK_POS(5), OK_LEN(5)), 1, 4)
                            Else
                                OK_DTM_2 = "01-01-0001"
                            End If
                            OK_TIJD_2 = Trim(Mid(line, OK_POS(6), OK_LEN(6)))
                            If Len(Trim(Mid(line, OK_POS(7), OK_LEN(7)))) = 8 Then
                                OK_DTM_17 = Mid(Mid(line, OK_POS(7), OK_LEN(7)), 7, 2) & "-" & Mid(Mid(line, OK_POS(7), OK_LEN(7)), 5, 2) & "-" & Mid(Mid(line, OK_POS(7), OK_LEN(7)), 1, 4)
                            Else
                                OK_DTM_17 = "01-01-0001"
                            End If
                            OK_TIJD_17 = Trim(Mid(line, OK_POS(8), OK_LEN(8)))
                            If Len(Trim(Mid(line, OK_POS(9), OK_LEN(9)))) = 8 Then
                                OK_DTM_64 = Mid(Mid(line, OK_POS(9), OK_LEN(9)), 7, 2) & "-" & Mid(Mid(line, OK_POS(9), OK_LEN(9)), 5, 2) & "-" & Mid(Mid(line, OK_POS(9), OK_LEN(9)), 1, 4)
                            Else
                                OK_DTM_64 = "01-01-0001"
                            End If
                            OK_TIJD_64 = Trim(Mid(line, OK_POS(10), OK_LEN(10)))
                            If Len(Trim(Mid(line, OK_POS(11), OK_LEN(11)))) = 8 Then
                                OK_DTM_63 = Mid(Mid(line, OK_POS(11), OK_LEN(11)), 7, 2) & "-" & Mid(Mid(line, OK_POS(11), OK_LEN(11)), 5, 2) & "-" & Mid(Mid(line, OK_POS(11), OK_LEN(11)), 1, 4)
                            Else
                                OK_DTM_63 = "01-01-0001"
                            End If
                            OK_TIJD_63 = Trim(Mid(line, OK_POS(12), OK_LEN(12)))
                            OK_RFF_BO = Trim(Mid(line, OK_POS(13), OK_LEN(13)))
                            OK_RFF_CR = Trim(Mid(line, OK_POS(14), OK_LEN(14)))
                            OK_RFF_PD = Trim(Mid(line, OK_POS(15), OK_LEN(15)))
                            OK_RFFCT = Trim(Mid(line, OK_POS(16), OK_LEN(16)))
                            If Len(Trim(Mid(line, OK_POS(17), OK_LEN(17)))) = 8 Then
                                OK_DTMCT = Mid(Mid(line, OK_POS(17), OK_LEN(17)), 7, 2) & "-" & Mid(Mid(line, OK_POS(17), OK_LEN(17)), 5, 2) & "-" & Mid(Mid(line, OK_POS(17), OK_LEN(17)), 1, 4)
                            Else
                                OK_DTMCT = "01-01-0001"
                            End If
                            OK_FLAGS(0) = Trim(Mid(line, OK_POS(18), OK_LEN(18)))
                            OK_FLAGS(1) = Trim(Mid(line, OK_POS(19), OK_LEN(19)))
                            OK_FLAGS(2) = Trim(Mid(line, OK_POS(20), OK_LEN(20)))
                            OK_FLAGS(3) = Trim(Mid(line, OK_POS(21), OK_LEN(21)))
                            OK_FLAGS(4) = Trim(Mid(line, OK_POS(22), OK_LEN(22)))
                            OK_FLAGS(5) = Trim(Mid(line, OK_POS(23), OK_LEN(23)))
                            OK_FLAGS(6) = Trim(Mid(line, OK_POS(24), OK_LEN(24)))
                            OK_FLAGS(7) = Trim(Mid(line, OK_POS(25), OK_LEN(25)))
                            OK_FLAGS(8) = Trim(Mid(line, OK_POS(26), OK_LEN(26)))
                            OK_FLAGS(9) = Trim(Mid(line, OK_POS(27), OK_LEN(27)))
                            OK_FLAGS(10) = Trim(Mid(line, OK_POS(28), OK_LEN(28)))
                            OK_FTXDSI = Trim(Mid(line, OK_POS(29), OK_LEN(29)))
                            OK_NAD_BY = Trim(Mid(line, OK_POS(30), OK_LEN(30)))
                            OK_NAD_DP = Trim(Mid(line, OK_POS(31), OK_LEN(31)))
                            OK_NAD_IV = Trim(Mid(line, OK_POS(32), OK_LEN(32)))
                            OK_NAD_SF = Trim(Mid(line, OK_POS(33), OK_LEN(33)))
                            OK_NAD_SU = Trim(Mid(line, OK_POS(34), OK_LEN(34)))
                            OK_NAD_UC = Trim(Mid(line, OK_POS(35), OK_LEN(35)))
                            OK_NAD_BCO = Trim(Mid(line, OK_POS(36), OK_LEN(36)))
                            OK_ONTVANGER = Trim(Mid(line, OK_POS(37), OK_LEN(37)))
                            Dim blnMatch As Boolean
                            blnMatch = WriteSOHead()
                            If blnMatch = False Then
                                'move file
                                File.Move(strSoTempPath & "\" & SO_FILENAME, strSoErrorPath & "\" & Format(Date.Now, "yyyyMMdd") & "_" & Format(Date.Now, "HHmmss") & "_" & SO_FILENAME)
                                Log("X", SO_FILENAME & " copied to the errors folder!", "Match_SO_data")
                                Return False
                                Exit Function
                            End If
                        Case Is = 1
                            OR_DEUAC = Trim(Mid(line, OR_POS(0), OR_LEN(0)))
                            OR_QTY = Trim(Mid(line, OR_POS(1), OR_LEN(1)))
                            OR_LEVARTCODE = Trim(Mid(line, OR_POS(2), OR_LEN(2)))
                            OR_DEARTOM = Trim(Mid(line, OR_POS(3), OR_LEN(3)))
                            OR_KLEUR = Trim(Mid(line, OR_POS(4), OR_LEN(4)))
                            OR_LENGTE = Trim(Mid(line, OR_POS(5), OR_LEN(5)))
                            OR_BREEDTE = Trim(Mid(line, OR_POS(6), OR_LEN(6)))
                            OR_HOOGTE = Trim(Mid(line, OR_POS(7), OR_LEN(7)))
                            OR_CUX = Trim(Mid(line, OR_POS(8), OR_LEN(8)))
                            OR_PIA = Trim(Mid(line, OR_POS(9), OR_LEN(9)))
                            OR_RFFLI1 = Trim(Mid(line, OR_POS(10), OR_LEN(10)))
                            OR_RFFLI2 = Trim(Mid(line, OR_POS(11), OR_LEN(11)))
                            If Len(Trim(Mid(line, OR_POS(12), OR_LEN(12)))) = 8 Then
                                OR_DTM_2 = Mid(Mid(line, OR_POS(12), OR_LEN(12)), 7, 2) & "-" & Mid(Mid(line, OR_POS(12), OR_LEN(12)), 5, 2) & "-" & Mid(Mid(line, OR_POS(12), OR_LEN(12)), 1, 4)
                            Else
                                OR_DTM_2 = "01-01-0001"
                            End If
                            OR_LINNR = Trim(Mid(line, OR_POS(13), OR_LEN(13)))
                            OR_PRI = Trim(Mid(line, OR_POS(14), OR_LEN(14)))
                            Dim blnMatchItems As Boolean
                            blnMatchItems = WriteSOItems()
                            If blnMatchItems = False Then
                                'move file
                                Return False
                                Exit Function
                            End If
                    End Select
                End While
            End Using

            OrderSave()

            SO_FILE = ""
            SO_FILENAME = ""

            Return True

        Catch ex As Exception

            File.Move(strSoTempPath & "\" & SO_FILENAME, strSoErrorPath & "\" & Format(Date.Now, "yyyyMMdd") & "_" & Format(Date.Now, "HHmmss") & "_" & SO_FILENAME)
            Call Log("X", ex.Message, "Match_SO_data")
            Log("X", SO_FILENAME & " copied to the errors folder!", "Match_SO_data")

            Return False

        End Try

        Return 0

    End Function

    Public Function WriteSOHead() As Boolean

        Try

            Dim blnMatchHead As Boolean
            blnMatchHead = CheckSOHead()
            If blnMatchHead = False Then
                Return False
                Exit Try
            End If

            oOrder = cmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

            With oOrder
                .CardCode = CARDCODE
                .NumAtCard = OK_BGM
                .DocDate = CDate(OK_K_ORDDAT)
                .TaxDate = CDate(OK_K_ORDDAT)

                'Referentie Rosis DocDueDate
                If OK_DTM_2.Year > 1999 Then
                    .DocDueDate = CDate(OK_DTM_2)
                Else
                    If OK_DTM_17.Year > 1999 Then
                        .DocDueDate = CDate(OK_DTM_17)
                    Else
                        If OK_DTM_64.Year > 1999 Then
                            .DocDueDate = CDate(OK_DTM_64)
                        Else
                            .DocDueDate = CDate(Trim("31/12/" & CStr(Date.Now.Year)))
                        End If
                    End If
                End If


                .ShipToCode = OK_NAD_DP
                .PayToCode = OK_NAD_IV

            End With

            'LOOP UDF

            Dim fld As SAPbobsCOM.Field

            For Each fld In oOrder.UserFields.Fields
                Select Case fld.Name
                    Case Is = "U_BGM" : fld.Value = OK_BGM
                    Case Is = "U_RFF" : fld.Value = OK_BGM

                    Case Is = "U_K_EANCODE" : fld.Value = OK_K_EANCODE
                    Case Is = "U_KNAAM" : fld.Value = OK_KNAAM
                    Case Is = "U_TEST"

                        If OK_TEST = "1" Then
                            fld.Value = "J"
                        Else
                            fld.Value = "N"
                        End If

                    Case Is = "U_DTM_2" : If OK_DTM_2.ToString <> "1-1-0001 0:00:00" Then fld.Value = Replace(OK_DTM_2.ToString, " 0:00:00", "")
                    Case Is = "U_TIJD_2" : fld.Value = OK_TIJD_2.ToString

                    Case Is = "U_DTM_17" : If OK_DTM_17.ToString <> "1-1-0001 0:00:00" Then fld.Value = Replace(OK_DTM_17.ToString, " 0:00:00", "")
                    Case Is = "U_TIJD_17" : fld.Value = OK_TIJD_17.ToString

                    Case Is = "U_DTM_64" : If OK_DTM_64.ToString <> "1-1-0001 0:00:00" Then fld.Value = Replace(OK_DTM_64.ToString, " 0:00:00", "")
                    Case Is = "U_TIJD_64" : fld.Value = OK_TIJD_64.ToString

                    Case Is = "U_DTM_63" : If OK_DTM_63.ToString <> "1-1-0001 0:00:00" Then fld.Value = Replace(OK_DTM_63.ToString, " 0:00:00", "")
                    Case Is = "U_TIJD_63" : fld.Value = OK_TIJD_63.ToString

                    Case Is = "U_RFF_BO" : fld.Value = OK_RFF_BO.ToString
                    Case Is = "U_RFF_CR" : fld.Value = OK_RFF_CR.ToString
                    Case Is = "U_RFF_PD" : fld.Value = OK_RFF_PD.ToString
                    Case Is = "U_RFFCT" : fld.Value = OK_RFFCT.ToString

                    Case Is = "U_DTMCT" : If OK_DTMCT.ToString <> "1-1-0001 0:00:00" Then fld.Value = Replace(OK_DTMCT.ToString, " 0:00:00", "")

                    Case Is = "U_FLAG0" : fld.Value = OK_FLAGS(0).ToString
                    Case Is = "U_FLAG1" : fld.Value = OK_FLAGS(1).ToString
                    Case Is = "U_FLAG2" : fld.Value = OK_FLAGS(2).ToString
                    Case Is = "U_FLAG3" : fld.Value = OK_FLAGS(3).ToString
                    Case Is = "U_FLAG4" : fld.Value = OK_FLAGS(4).ToString
                    Case Is = "U_FLAG5" : fld.Value = OK_FLAGS(5).ToString
                    Case Is = "U_FLAG6" : fld.Value = OK_FLAGS(6).ToString
                    Case Is = "U_FLAG7" : fld.Value = OK_FLAGS(7).ToString
                    Case Is = "U_FLAG8" : fld.Value = OK_FLAGS(8).ToString
                    Case Is = "U_FLAG9" : fld.Value = OK_FLAGS(9).ToString
                    Case Is = "U_FLAG10" : fld.Value = OK_FLAGS(10).ToString

                    Case Is = "U_NAD_BY" : fld.Value = OK_NAD_BY.ToString
                    Case Is = "U_NAD_DP" : fld.Value = OK_NAD_DP.ToString
                    Case Is = "U_NAD_IV" : fld.Value = OK_NAD_IV.ToString
                    Case Is = "U_NAD_SF" : fld.Value = OK_NAD_SF.ToString
                    Case Is = "U_NAD_SU" : fld.Value = OK_NAD_SU.ToString
                    Case Is = "U_NAD_UC" : fld.Value = OK_NAD_UC.ToString
                    Case Is = "U_NAD_BCO" : fld.Value = OK_NAD_BCO.ToString

                    Case Is = "U_ONTVANGER" : fld.Value = OK_ONTVANGER.ToString
                    Case Is = "U_EDI_BERICHT" : fld.Value = "Ja"
                    Case Is = "U_EDI_IMP_TIJD" : fld.Value = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")
                    Case Is = "U_EDI_EXPORT" : fld.Value = "Nee"

                End Select
            Next

            'END UDF

            Return True

        Catch ex As Exception

            Call Log("X", ex.Message, "WriteSOHead")
            Return False


        End Try

    End Function

    Public Function CheckKoperAdres(ByVal strEancodeKoper) As String

        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = cmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        KOPERADRES = ""

        oRecordSet.DoQuery("SELECT LicTradNum FROM OCRD WHERE U_K_EANCODE = '" & strEancodeKoper & "'")

        If oRecordSet.RecordCount = 1 Then
            KOPERADRES = oRecordSet.Fields(0).Value
        ElseIf oRecordSet.RecordCount > 1 Then
            Call Log("X", "Error: Duplicate EANcode Buyers found!", "CheckSOHead")
            KOPERADRES = ""
        Else
            Call Log("X", "Error: Match EANcode Buyers NOT found!", "CheckSOHead")
            KOPERADRES = ""
        End If

        oRecordSet = Nothing

        Return KOPERADRES

    End Function

    Public Function CheckSOHead() As Boolean

        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = cmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        CARDCODE = ""

        oRecordSet.DoQuery("SELECT CardCode FROM OCRD WHERE U_K_EANCODE = '" & OK_K_EANCODE & "'")

        If oRecordSet.RecordCount = 1 Then
            CARDCODE = oRecordSet.Fields(0).Value
            Return True
        ElseIf oRecordSet.RecordCount > 1 Then
            Call Log("X", "Error: Duplicate customers found!", "CheckSOHead")
            Return False
        Else
            Call Log("X", "Error: Match customer NOT found!", "CheckSOHead")
            Return False
        End If

        oRecordSet = Nothing

    End Function

    Public Function WriteSOItems() As Boolean

        Try

            Dim blnMatchItems As Boolean
            blnMatchItems = CheckSOItems()

            If blnMatchItems = False Then
                Return False
                Exit Try
            End If

            'items
            With oOrder.Lines
                .ItemCode = ITEMCODE
                .Quantity = Replace(OR_QTY, ",", ".") * SALPACKUN '3 decimalen

                If OR_DTM_2 = "01-01-0001" Then
                    .ShipDate = OK_DTM_2
                Else
                    .ShipDate = OR_DTM_2
                End If

                Dim fld As SAPbobsCOM.Field

                For Each fld In oOrder.Lines.UserFields.Fields
                    Select Case fld.Name
                        Case Is = "U_DEUAC" : fld.Value = OR_DEUAC.ToString
                        Case Is = "U_LEVARTCODE" : fld.Value = OR_LEVARTCODE.ToString
                        Case Is = "U_DEARTOM" : fld.Value = OR_DEARTOM.ToString
                        Case Is = "U_KLEUR" : fld.Value = OR_KLEUR.ToString
                        Case Is = "U_LENGTE" : fld.Value = OR_LENGTE.ToString
                        Case Is = "U_BREEDTE" : fld.Value = OR_BREEDTE.ToString
                        Case Is = "U_HOOGTE" : fld.Value = OR_HOOGTE.ToString
                        Case Is = "U_CUX" : fld.Value = OR_CUX.ToString
                        Case Is = "U_PIA" : fld.Value = OR_PIA.ToString
                        Case Is = "U_RFFLI1" : fld.Value = OR_RFFLI1.ToString
                        Case Is = "U_RFFLI2" : fld.Value = OR_RFFLI2.ToString
                        Case Is = "U_LINNR" : fld.Value = OR_LINNR.ToString
                        Case Is = "U_PRI" : fld.Value = OR_PRI.ToString
                    End Select
                Next

                .Add()
            End With

            Return True

        Catch ex As Exception

            Call Log("X", ex.Message, "WriteSOItems")
            Return False

        End Try

    End Function

    Public Function CheckSOItems() As Boolean

        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = cmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        ITEMCODE = ""
        ITEMNAME = ""
        SALPACKUN = 1

        oRecordSet.DoQuery("SELECT ItemCode, ItemName, SalPackUn FROM OITM WHERE U_EAN_Handels_EH = '" & OR_DEUAC & "'")

        If oRecordSet.RecordCount = 1 Then
            ITEMCODE = oRecordSet.Fields(0).Value
            ITEMNAME = oRecordSet.Fields(1).Value
            SALPACKUN = oRecordSet.Fields(2).Value
            Return True
        ElseIf oRecordSet.RecordCount > 1 Then
            Call Log("X", "Error: Duplicate Eancodes found! Eancode " & OR_DEUAC, "CheckSOItems")
            Return False
        Else
            Call Log("X", "Error: Match Eancode NOT found!", "CheckSOItems")
            Return False
        End If

        oRecordSet = Nothing

    End Function

    Public Function OrderSave() As Integer

        Dim ErrMsg As String = ""
        Dim nErr As Integer
        Dim lRetCode As Integer = 0

        lRetCode = oOrder.Add()

        If lRetCode <> 0 Then
            cmp.GetLastError(nErr, ErrMsg)

            File.Move(strSoTempPath & "\" & SO_FILENAME, strSoErrorPath & "\" & Format(Date.Now, "yyyyMMdd") & "_" & Format(Date.Now, "HHmmss") & "_" & SO_FILENAME)
            Call Log("X", "Error: " & ErrMsg & "(" & Str(nErr) & ")", "OrderSave")
            Log("X", SO_FILENAME & " copied to the errors folder!", "Match_SO_data")

        Else
            If iSendNotification = 1 Then Mail_to_SO_receiver(oOrder.DocEntry, oOrder.DocDate, oOrder.DocNum)

            Call Log("V", "Order written in SAP BO!", "OrderSave")
            File.Move(strSoTempPath & "\" & SO_FILENAME, strSoDonePath & "\" & Format(Date.Now, "yyyyMMdd") & "_" & Format(Date.Now, "HHmmss") & "_" & SO_FILENAME)
        End If

        Return 0


    End Function

    Public Function CreateUdfFiledsText() As Integer

        ''MessageBox.Show(applicationPath & "\udf.xml")
        If File.Exists(applicationPath & "\udf.xml") = False Then
            Return 0
            Exit Function
        End If

        Dim ds As New DataSet
        ds.ReadXml(applicationPath & "\udf.xml")

        Try
            If ds.Tables("udf").Rows.Count > 0 Then
                Dim i As Integer
                For i = 0 To ds.Tables("udf").Rows.Count - 1
                    CreateUdf(ds.Tables("udf").Rows(i).Item(0), ds.Tables("udf").Rows(i).Item(1), ds.Tables("udf").Rows(i).Item(2), SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, ds.Tables("udf").Rows(i).Item(3), False, False, "")
                Next
            End If

        Catch ex As Exception

            Call Log("X", ex.Message, "CreateUdfFiledsText")

        End Try

        ds.Dispose()

        Return 0

    End Function

    Private Function CreateUdf(ByVal sTablename As String, ByVal sFiledName As String, ByVal sDescription As String, _
                     ByVal boType As BoFieldTypes, ByVal boSubType As BoFldSubTypes, ByVal iEditSize As Integer, _
                       ByVal bMandatory As Boolean, ByVal bDefault As Boolean, ByVal sDefault As String) As Long

        Dim oUDF As SAPbobsCOM.UserFieldsMD
        oUDF = cmp.GetBusinessObject(BoObjectTypes.oUserFields)
        GC.Collect()

        Dim lRetCode As Long
        Dim sStr As String = ""

        With oUDF
            .TableName = sTablename
            .Name = sFiledName
            .Description = sDescription
            .Type = boType
            .SubType = boSubType
            .EditSize = iEditSize
            .Size = iEditSize
            If bDefault = True Then .DefaultValue = sDefault

            If bMandatory = True Then
                .Mandatory = BoYesNoEnum.tYES
            Else
                .Mandatory = BoYesNoEnum.tNO
            End If

        End With

        lRetCode = oUDF.Add

        If lRetCode <> 0 Then
            cmp.GetLastError(lRetCode, sStr)
            Call Log("X", "Error: " & sStr & "(" & Str(lRetCode) & ")", "CreateUdf")
        Else
            Call Log("V", "Udf " & sFiledName & " successfully created!", "CreateUdf")
        End If

        Return lRetCode

        oUDF = Nothing

    End Function

    Public Function Mail_to_SO_receiver(sDocEntry As Integer, sDocDate As String, iDocNumber As Integer) As Boolean

        Try
            Using mailMsg As New Mail.MailMessage()

                Dim SmtpMail As New Mail.SmtpClient() With {.Host = sSmtp, .Port = iSmtpPort}

                If bSmtpUserSecurity = True Then
                    SmtpMail.Credentials = New System.Net.NetworkCredential(sSmtpUser, sSmtpPassword)
                End If

                With mailMsg
                    .From = New System.Net.Mail.MailAddress(sSenderEmail, sSenderName)
                    .To.Add(sOrderMailTo)
                    .Subject = String.Format("Order {0} imported", sDocEntry)

                    Using rFile As New StreamReader(applicationPath & "\email_o.txt")
                        Dim sBody As String
                        sBody = rFile.ReadToEnd()
                        sBody = Replace(sBody, "::NAME::", sOrderMailToFullName)
                        sBody = Replace(sBody, "::DOCENTRY::", sDocEntry)
                        sBody = Replace(sBody, "::DOCDATE::", sDocDate)
                        sBody = Replace(sBody, "::DOCNUM::", CStr(iDocNumber))
                        .Body = sBody
                        rFile.Close()
                    End Using

                End With

                SmtpMail.Send(mailMsg)

            End Using

            Call Log("V", "Order notification sent!", "Mail_to_SO_receiver")

            Return True

        Catch ex As Exception

            Call Log("X", "Order notification was not sent!", "Mail_to_SO_receiver")

            Return False

        End Try

    End Function

    Public Function CheckAndExport_delivery() As Integer

        Try

            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = cmp.GetBusinessObject(BoObjectTypes.BoRecordset)

            oRecordSet.DoQuery("SELECT DocEntry  FROM ODLN WHERE Canceled='N' and U_EDI_BERICHT = 'Ja' AND U_EDI_EXPORT = 'Ja' AND (U_EDI_DEL_EXP is NUll OR U_EDI_DEL_EXP = 'Nee')")

            If oRecordSet.RecordCount < 1 Then

                Call Log("V", "No Delivery notes found to export!", "CheckAndExport_delivery")

            ElseIf oRecordSet.RecordCount > 0 Then

                Dim i As Integer
                oRecordSet.MoveFirst()
                For i = 1 To oRecordSet.RecordCount
                    Create_delivery_file(oRecordSet.Fields.Item(0).Value)
                    oRecordSet.MoveNext()
                Next

            End If

            oRecordSet = Nothing

        Catch ex As Exception

            Call Log("X", ex.Message, "CheckAndExport_delivery")

        End Try

        Return 0

    End Function

    Public Function Create_delivery_file(ByVal DocEntry As Integer) As Integer

        Dim oDelivery As SAPbobsCOM.Documents
        oDelivery = cmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)

        Dim blnFind As Boolean = oDelivery.GetByKey(DocEntry)

        If oDelivery.UserFields.Fields.Item("U_DESADV").Value.ToString() = "1" Then
            '' DESADV =1
            If blnFind = True Then

                Try
                    Using wf As StreamWriter = New StreamWriter(strDeliveryPath & "\" & "DEL_" & oDelivery.DocNum & ".DAT")

                        Dim strHeader As String = ""
                        Dim strLine As String

                        With oDelivery.UserFields.Fields

                            DK_K_EANCODE = ""
                            DK_PBTEST = ""
                            DK_KNAAM = ""
                            DK_BGM = ""
                            DK_DTM_137 = ""
                            DK_DTM_2 = ""
                            DK_TIJD_2 = ""
                            DK_DTM_17 = ""
                            DK_TIJD_17 = ""
                            DK_DTM_64 = ""
                            DK_TIJD_64 = ""
                            DK_DTM_63 = ""
                            DK_TIJD_63 = ""
                            DK_BH_DAT = ""
                            DK_BH_TIJD = ""
                            DK_RFF = ""
                            DK_RFFVN = ""
                            DK_BH_EAN = ""
                            DK_NAD_BY = ""
                            DK_NAD_DP = ""
                            DK_NAD_SU = ""
                            DK_NAD_UC = ""
                            DK_DESATYPE = ""
                            DK_ONTVANGER = ""

                            DK_K_EANCODE = Format_field(.Item("U_K_EANCODE").Value, 13, "STRING")

                            If .Item("U_TEST").Value = "J" Then
                                DK_PBTEST = Format_field("1", 1, "STRING")
                            Else
                                DK_PBTEST = Format_field(" ", 1, "STRING")
                            End If

                            DK_KNAAM = Format_field(.Item("U_KNAAM").Value, 14, "STRING")
                            DK_BGM = Format_field(oDelivery.DocNum, 35, "STRING")
                            DK_DTM_137 = Format_field(Format(oDelivery.DocDate, "yyyyMMdd"), 8, "STRING")

                            If Len(Trim(.Item("U_DTM_2").Value)) > 0 Then
                                DK_DTM_2 = Format_field(Format(CDate(.Item("U_DTM_2").Value), "yyyyMMdd"), 8, "STRING")
                            Else
                                DK_DTM_2 = Format_field(" ", 8, "STRING")
                            End If

                            DK_TIJD_2 = Format_field(.Item("U_TIJD_2").Value, 5, "STRING")

                            If Len(Trim(.Item("U_DTM_17").Value)) > 0 Then
                                DK_DTM_17 = Format_field(Format(CDate(.Item("U_DTM_17").Value), "yyyyMMdd"), 8, "STRING")
                            Else
                                DK_DTM_17 = Format_field(" ", 8, "STRING")
                            End If

                            DK_TIJD_17 = Format_field(.Item("U_TIJD_17").Value, 5, "STRING")

                            If Len(Trim(.Item("U_DTM_64").Value)) > 0 Then
                                DK_DTM_64 = Format_field(Format(CDate(.Item("U_DTM_64").Value), "yyyyMMdd"), 8, "STRING")
                            Else
                                DK_DTM_64 = Format_field(" ", 8, "STRING")
                            End If

                            DK_TIJD_64 = Format_field(.Item("U_TIJD_64").Value, 5, "STRING")

                            If Len(Trim(.Item("U_DTM_63").Value)) > 0 Then
                                DK_DTM_63 = Format_field(Format(CDate(.Item("U_DTM_63").Value), "yyyyMMdd"), 8, "STRING")
                            Else
                                DK_DTM_63 = Format_field(" ", 8, "STRING")
                            End If

                            DK_TIJD_63 = Format_field(.Item("U_TIJD_63").Value, 5, "STRING")
                            DK_BH_DAT = Format_field(.Item("U_BH_DAT").Value, 8, "STRING")
                            DK_BH_TIJD = Format_field(.Item("U_BH_TIJD").Value, 5, "STRING")
                            DK_RFF = Format_field(.Item("U_RFF").Value, 35, "STRING")
                            DK_RFFVN = Format_field(.Item("U_RFFVN").Value, 35, "STRING")
                            DK_BH_EAN = Format_field(.Item("U_BH_EAN").Value, 13, "STRING")
                            DK_NAD_BY = Format_field(.Item("U_NAD_BY").Value, 13, "STRING")
                            DK_NAD_DP = Format_field(.Item("U_NAD_DP").Value, 13, "STRING")
                            DK_NAD_SU = Format_field(.Item("U_NAD_SU").Value, 13, "STRING")
                            DK_NAD_UC = Format_field(.Item("U_NAD_UC").Value, 13, "STRING")
                            ''DK_DESATYPE = strDesAdvLevel 'see desadvlevel in settings file
                            DK_DESATYPE = oDelivery.UserFields.Fields.Item("U_DESADV").Value.ToString()
                            DK_ONTVANGER = Format_field(.Item("U_ONTVANGER").Value, 13, "STRING")

                        End With

                        strHeader = DK_K_EANCODE & _
                        DK_PBTEST & _
                        DK_KNAAM & _
                        DK_BGM & _
                        DK_DTM_137 & _
                        DK_DTM_2 & _
                        DK_TIJD_2 & _
                        DK_DTM_17 & _
                        DK_TIJD_17 & _
                        DK_DTM_64 & _
                        DK_TIJD_64 & _
                        DK_DTM_63 & _
                        DK_TIJD_63 & _
                        DK_BH_DAT & _
                        DK_BH_TIJD & _
                        DK_RFF & _
                        DK_RFFVN & _
                        DK_BH_EAN & _
                        DK_NAD_BY & _
                        DK_NAD_DP & _
                        DK_NAD_SU & _
                        DK_NAD_UC & _
                        DK_DESATYPE & _
                        DK_ONTVANGER

                        wf.WriteLine("0" & strHeader)

                        Dim i As Integer = 0

                        For i = 0 To oDelivery.Lines.Count - 1

                            strLine = ""

                            DR_DEUAC = ""
                            DR_OLDUAC = ""
                            DR_DEARTNR = ""
                            DR_DEARTOM = ""
                            DR_PIA = ""
                            DR_BATCH = ""
                            DR_QTY = ""
                            DR_ARTEENHEID = ""
                            DR_RFFONID = ""
                            DR_RFFONORD = ""
                            DR_DTM_23E = ""
                            DR_TGTDATUM = ""
                            DR_GEWICHT = ""
                            DR_FEENHEID = ""
                            DR_QTY_AFW = ""
                            DR_REDEN = ""
                            DR_GINTYPE = ""
                            DR_GINID = ""
                            DR_BATCHH = ""

                            oDelivery.Lines.SetCurrentLine(i)

                            With oDelivery.Lines.UserFields.Fields

                                DR_DEUAC = Format_field(ItemEAN(oDelivery.Lines.ItemCode), 14, "STRING")
                                DR_OLDUAC = Format_field(.Item("U_OLDUAC").Value, 14, "STRING")
                                DR_DEARTNR = Format_field(.Item("U_DEARTNR").Value, 9, "STRING")
                                DR_DEARTOM = Format_field(.Item("U_DEARTOM").Value, 35, "STRING")
                                DR_PIA = Format_field(.Item("U_PIA").Value, 10, "STRING")
                                DR_BATCH = Format_field(.Item("U_BATCH").Value, 35, "STRING")

                                '' DR_QTY = Format_field(SSCC_qty(DocEntry, i), 17, "STRING")
                                DR_QTY = Format_field(oDelivery.Lines.Quantity, 17, "STRING")

                                DR_ARTEENHEID = Format_field(.Item("U_ARTEENHEID").Value, 3, "STRING")
                                DR_RFFONID = Format_field(.Item("U_RFFONID").Value, 6, "STRING")
                                DR_RFFONORD = Format_field(.Item("U_RFFONORD").Value, 35, "STRING")
                                DR_DTM_23E = Format_field(.Item("U_DTM_23E").Value, 8, "STRING")
                                DR_TGTDATUM = Format_field(.Item("U_TGTDATUM").Value, 1, "STRING")
                                DR_GEWICHT = Format_field(.Item("U_GEWICHT").Value, 6, "STRING")
                                DR_FEENHEID = Format_field(.Item("U_FEENHEID").Value, 3, "STRING")
                                DR_QTY_AFW = Format_field(.Item("U_QTY_AFW").Value, 17, "STRING")
                                DR_REDEN = Format_field(.Item("U_REDEN").Value, 3, "STRING")
                                DR_GINTYPE = Format_field(" ", 3, "STRING")
                                DR_GINID = Format_field(" ", 19, "STRING")
                                DR_BATCHH = Format_field(.Item("U_BATCH").Value, 35, "STRING")

                            End With

                            strLine = DR_DEUAC & _
                            DR_OLDUAC & _
                            DR_DEARTNR & _
                            DR_DEARTOM & _
                            DR_PIA & _
                            DR_BATCH & _
                            DR_QTY & _
                            DR_ARTEENHEID & _
                            DR_RFFONID & _
                            DR_RFFONORD & _
                            DR_DTM_23E & _
                            DR_TGTDATUM & _
                            DR_GEWICHT & _
                            DR_FEENHEID & _
                            DR_QTY_AFW & _
                            DR_REDEN & _
                            DR_GINTYPE & _
                            DR_GINID & _
                            DR_BATCHH

                            If oDelivery.Lines.TreeType <> BoItemTreeTypes.iIngredient Then wf.WriteLine("1" & strLine)

                        Next

                        wf.Close()

                        oDelivery.UserFields.Fields.Item("U_EDI_DEL_EXP").Value = "Ja"
                        oDelivery.UserFields.Fields.Item("U_EDI_DELEXP_TIJD").Value = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")
                        oDelivery.Update()

                        If iSendNotification = 1 Then Mail_to_DL_receiver(oDelivery.DocEntry, oDelivery.DocDate, oDelivery.DocNum)

                        Call Log("V", "Delivery note file created!", "Create_delivery_file")

                    End Using

                Catch ex As Exception
                    File.Delete(strDeliveryPath & "\" & "DEL_" & oDelivery.DocNum & ".DAT")
                    Call Log("X", ex.Message, "Create_delivery_file")
                End Try

            Else
                Call Log("X", "Delivery note " & DocEntry & " not found!", "Create_delivery_file")
            End If
        Else
            '' MessageBox.Show("desadv4", DocEntry.ToString())

            If blnFind = True Then
                Dim rs As SAPbobsCOM.Recordset = cmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                rs.DoQuery("select DocEntry,SSCC_Code,ItemCode,RegelNr,Quantity from sscc where DocEntry=" + DocEntry.ToString())
                If rs.RecordCount > 0 Then
                    '' Call Log("D", "desadv4", DocEntry.ToString())
                    '' er zijn pakketten gevonden
                    Using wf As StreamWriter = New StreamWriter(strDeliveryPath & "\" & "DEL_" & oDelivery.DocNum & ".DAT")
                        For PakketTeller = 1 To rs.RecordCount
                            Try
                                Dim strHeader As String = ""
                                Dim strLine As String

                                With oDelivery.UserFields.Fields

                                    DK_K_EANCODE = ""
                                    DK_PBTEST = ""
                                    DK_KNAAM = ""
                                    DK_BGM = ""
                                    DK_DTM_137 = ""
                                    DK_DTM_2 = ""
                                    DK_TIJD_2 = ""
                                    DK_DTM_17 = ""
                                    DK_TIJD_17 = ""
                                    DK_DTM_64 = ""
                                    DK_TIJD_64 = ""
                                    DK_DTM_63 = ""
                                    DK_TIJD_63 = ""
                                    DK_BH_DAT = ""
                                    DK_BH_TIJD = ""
                                    DK_RFF = ""
                                    DK_RFFVN = ""
                                    DK_BH_EAN = ""
                                    DK_NAD_BY = ""
                                    DK_NAD_DP = ""
                                    DK_NAD_SU = ""
                                    DK_NAD_UC = ""
                                    DK_DESATYPE = ""
                                    DK_ONTVANGER = ""

                                    DK_K_EANCODE = Format_field(.Item("U_K_EANCODE").Value, 13, "STRING")

                                    If .Item("U_TEST").Value = "J" Then
                                        DK_PBTEST = Format_field("1", 1, "STRING")
                                    Else
                                        DK_PBTEST = Format_field(" ", 1, "STRING")
                                    End If

                                    DK_KNAAM = Format_field(.Item("U_KNAAM").Value, 14, "STRING")
                                    DK_BGM = Format_field(oDelivery.DocNum, 35, "STRING")
                                    DK_DTM_137 = Format_field(Format(oDelivery.DocDate, "yyyyMMdd"), 8, "STRING")

                                    If Len(Trim(.Item("U_DTM_2").Value)) > 0 Then
                                        DK_DTM_2 = Format_field(Format(CDate(.Item("U_DTM_2").Value), "yyyyMMdd"), 8, "STRING")
                                    Else
                                        DK_DTM_2 = Format_field(" ", 8, "STRING")
                                    End If

                                    DK_TIJD_2 = Format_field(.Item("U_TIJD_2").Value, 5, "STRING")

                                    If Len(Trim(.Item("U_DTM_17").Value)) > 0 Then
                                        DK_DTM_17 = Format_field(Format(CDate(.Item("U_DTM_17").Value), "yyyyMMdd"), 8, "STRING")
                                    Else
                                        DK_DTM_17 = Format_field(" ", 8, "STRING")
                                    End If

                                    DK_TIJD_17 = Format_field(.Item("U_TIJD_17").Value, 5, "STRING")

                                    If Len(Trim(.Item("U_DTM_64").Value)) > 0 Then
                                        DK_DTM_64 = Format_field(Format(CDate(.Item("U_DTM_64").Value), "yyyyMMdd"), 8, "STRING")
                                    Else
                                        DK_DTM_64 = Format_field(" ", 8, "STRING")
                                    End If

                                    DK_TIJD_64 = Format_field(.Item("U_TIJD_64").Value, 5, "STRING")

                                    If Len(Trim(.Item("U_DTM_63").Value)) > 0 Then
                                        DK_DTM_63 = Format_field(Format(CDate(.Item("U_DTM_63").Value), "yyyyMMdd"), 8, "STRING")
                                    Else
                                        DK_DTM_63 = Format_field(" ", 8, "STRING")
                                    End If

                                    DK_TIJD_63 = Format_field(.Item("U_TIJD_63").Value, 5, "STRING")
                                    DK_BH_DAT = Format_field(.Item("U_BH_DAT").Value, 8, "STRING")
                                    DK_BH_TIJD = Format_field(.Item("U_BH_TIJD").Value, 5, "STRING")
                                    DK_RFF = Format_field(.Item("U_RFF").Value, 35, "STRING")
                                    DK_RFFVN = Format_field(.Item("U_RFFVN").Value, 35, "STRING")
                                    DK_BH_EAN = Format_field(.Item("U_BH_EAN").Value, 13, "STRING")
                                    DK_NAD_BY = Format_field(.Item("U_NAD_BY").Value, 13, "STRING")
                                    DK_NAD_DP = Format_field(.Item("U_NAD_DP").Value, 13, "STRING")
                                    DK_NAD_SU = Format_field(.Item("U_NAD_SU").Value, 13, "STRING")
                                    DK_NAD_UC = Format_field(.Item("U_NAD_UC").Value, 13, "STRING")
                                    ''DK_DESATYPE = strDesAdvLevel 'see desadvlevel in settings file
                                    DK_DESATYPE = oDelivery.UserFields.Fields.Item("U_DESADV").Value.ToString()
                                    DK_ONTVANGER = Format_field(.Item("U_ONTVANGER").Value, 13, "STRING")

                                End With

                                strHeader = DK_K_EANCODE & _
                                DK_PBTEST & _
                                DK_KNAAM & _
                                DK_BGM & _
                                DK_DTM_137 & _
                                DK_DTM_2 & _
                                DK_TIJD_2 & _
                                DK_DTM_17 & _
                                DK_TIJD_17 & _
                                DK_DTM_64 & _
                                DK_TIJD_64 & _
                                DK_DTM_63 & _
                                DK_TIJD_63 & _
                                DK_BH_DAT & _
                                DK_BH_TIJD & _
                                DK_RFF & _
                                DK_RFFVN & _
                                DK_BH_EAN & _
                                DK_NAD_BY & _
                                DK_NAD_DP & _
                                DK_NAD_SU & _
                                DK_NAD_UC & _
                                DK_DESATYPE & _
                                DK_ONTVANGER

                                If PakketTeller = 1 Then wf.WriteLine("0" & strHeader)

                                Dim i As Integer = 0

                                ''For i = 0 To oDelivery.Lines.Count - 1

                                strLine = ""

                                DR_DEUAC = ""
                                DR_OLDUAC = ""
                                DR_DEARTNR = ""
                                DR_DEARTOM = ""
                                DR_PIA = ""
                                DR_BATCH = ""
                                DR_QTY = ""
                                DR_ARTEENHEID = ""
                                DR_RFFONID = ""
                                DR_RFFONORD = ""
                                DR_DTM_23E = ""
                                DR_TGTDATUM = ""
                                DR_GEWICHT = ""
                                DR_FEENHEID = ""
                                DR_QTY_AFW = ""
                                DR_REDEN = ""
                                DR_GINTYPE = ""
                                DR_GINID = ""
                                DR_BATCHH = ""

                                ''oDelivery.Lines.SetCurrentLine(i)

                                If strDesAdvLevel = "4" Then

                                    Dim sSSCC As String
                                    sSSCC = SSCC(DocEntry, rs.Fields.Item("RegelNr").Value.ToString(), rs.Fields.Item("ItemCode").Value.ToString())
                                    sSSCC = rs.Fields.Item(1).Value.ToString()
                                    ''Call Log("1331", "a", "b")
                                    '' deze regel nog aanpassen, verwijst naar udf op regelniveau
                                    ''wf.WriteLine(String.Format("1{0}1{1}{2}{3}", Space(117), Space(98), Format_field(oDelivery.Lines.UserFields.Fields.Item("U_GINTYPE").Value, 3, "STRING"), Format_field(sSSCC, 19, "STRING")))
                                    wf.WriteLine(String.Format("1{0}1{1}{2}{3}", Space(117), Space(98), Format_field("201", 3, "STRING"), Format_field(sSSCC, 19, "STRING")))

                                    If sSSCC = "" Then
                                        'move file
                                        Call Log("1331 sscc leeg", "a", "b")
                                        '' delivery file staat nu nog open --> exeption raised
                                        File.Move(strSoTempPath & "\" & SO_FILENAME, strSoErrorPath & "\" & Format(Date.Now, "yyyyMMdd") & "_" & Format(Date.Now, "HHmmss") & "_" & SO_FILENAME)
                                        Return 0
                                        Exit Function
                                    End If

                                End If

                                With oDelivery.Lines.UserFields.Fields

                                    DR_DEUAC = Format_field(ItemEAN(rs.Fields.Item("ItemCode").Value.ToString()), 14, "STRING")
                                    DR_OLDUAC = Format_field(.Item("U_OLDUAC").Value, 14, "STRING")
                                    DR_DEARTNR = Format_field(.Item("U_DEARTNR").Value, 9, "STRING")
                                    DR_DEARTOM = Format_field(.Item("U_DEARTOM").Value, 35, "STRING")
                                    DR_PIA = Format_field(.Item("U_PIA").Value, 10, "STRING")
                                    DR_BATCH = Format_field(.Item("U_BATCH").Value, 35, "STRING")


                                    DR_QTY = Format_field(SSCC_qty(DocEntry, Convert.ToInt32(rs.Fields.Item("RegelNr").Value)), 17, "STRING")

                                    DR_ARTEENHEID = Format_field(.Item("U_ARTEENHEID").Value, 3, "STRING")

                                    DR_RFFONID = Format_field(.Item("U_RFFONID").Value, 6, "STRING")
                                    DR_RFFONORD = Format_field(.Item("U_RFFONORD").Value, 35, "STRING")
                                    DR_DTM_23E = Format_field(.Item("U_DTM_23E").Value, 8, "STRING")
                                    DR_TGTDATUM = Format_field(.Item("U_TGTDATUM").Value, 1, "STRING")
                                    DR_GEWICHT = Format_field(.Item("U_GEWICHT").Value, 6, "STRING")
                                    DR_FEENHEID = Format_field(.Item("U_FEENHEID").Value, 3, "STRING")
                                    DR_QTY_AFW = Format_field(.Item("U_QTY_AFW").Value, 17, "STRING")
                                    DR_REDEN = Format_field(.Item("U_REDEN").Value, 3, "STRING")

                                    DR_GINTYPE = Format_field(.Item("U_GINTYPE").Value, 3, "STRING")

                                    DR_GINID = Format_field(" ", 19, "STRING")

                                    DR_BATCHH = Format_field(.Item("U_BATCH").Value, 35, "STRING")

                                End With

                                strLine = DR_DEUAC & DR_OLDUAC & DR_DEARTNR & DR_DEARTOM & DR_PIA & DR_BATCH & DR_QTY & DR_ARTEENHEID & _
                                DR_RFFONID & DR_RFFONORD & DR_DTM_23E & DR_TGTDATUM & DR_GEWICHT & _
                                DR_FEENHEID & _
                                DR_QTY_AFW & _
                                DR_REDEN & _
                                DR_GINTYPE & _
                                DR_GINID & _
                                DR_BATCHH

                                If oDelivery.Lines.TreeType <> BoItemTreeTypes.iIngredient Then wf.WriteLine("1" & strLine)

                                ''Next



                                '' tijdelijk uitgezet
                                oDelivery.UserFields.Fields.Item("U_EDI_DEL_EXP").Value = "Ja"
                                oDelivery.UserFields.Fields.Item("U_EDI_DELEXP_TIJD").Value = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")
                                oDelivery.Update()

                                If iSendNotification = 1 Then Mail_to_DL_receiver(oDelivery.DocEntry, oDelivery.DocDate, oDelivery.DocNum)

                                Call Log("V", "Delivery note file created!", "Create_delivery_file")


                            Catch ex As Exception
                                File.Delete(strDeliveryPath & "\" & "DEL_" & oDelivery.DocNum & ".DAT")
                                Call Log("X", ex.Message, "Create_delivery_file")
                            End Try
                            rs.MoveNext()
                        Next
                        wf.Close()
                    End Using

                Else
                    Call Log("X", "Delivery note deasdv4 : " & DocEntry & " not found!", "Create_delivery_file")
                End If

            End If

        End If


        Return 0

    End Function

    Public Function SSCC(DocEntry As Integer, RegelNr As Integer, sItemCode As String) As String
        Call Log("sscc", "sscc", "b")
        Dim new_sscc As String

        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = cmp.GetBusinessObject(BoObjectTypes.BoRecordset)

        oRecordSet.DoQuery("SELECT SSCC_Code FROM SSCC WHERE DocEntry = " & DocEntry & " AND RegelNr = " & RegelNr & " AND ItemCode = '" & sItemCode & "'")
        If oRecordSet.RecordCount > 0 Then
            new_sscc = oRecordSet.Fields.Item(0).Value
            Call Log("V", "SSCC code found!", "SSCC")
        Else
            new_sscc = ""
            Call Log("X", "SSCC code for document " & DocEntry & " not found!", "SSCC")

        End If

        Return new_sscc

    End Function

    Public Function SSCC_qty(DocEntry As Integer, RegelNr As Integer) As String

        Dim new_sscc_qty As String

        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = cmp.GetBusinessObject(BoObjectTypes.BoRecordset)

        oRecordSet.DoQuery("SELECT Quantity FROM SSCC WHERE DocEntry = " & DocEntry & " AND RegelNr = " & RegelNr)

        If oRecordSet.RecordCount > 0 Then
            new_sscc_qty = oRecordSet.Fields.Item(0).Value
            Call Log("V", "SSCC code (qty) found!", "SSCC")
        Else
            new_sscc_qty = ""
            Call Log("X", "SSCC code (QRY) for document " & DocEntry & " not found!", "SSCC")
        End If

        Return new_sscc_qty

    End Function

    Public Function ItemEAN(ICode As String) As String

        Dim i_ean As String

        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = cmp.GetBusinessObject(BoObjectTypes.BoRecordset)

        oRecordSet.DoQuery("SELECT U_EAN_Handels_EH FROM OITM WHERE ItemCode = '" & ICode & "'")

        If oRecordSet.RecordCount > 0 Then
            i_ean = oRecordSet.Fields.Item(0).Value
            Call Log("V", "ItemEAN code found!", "ItemEAN")
        Else
            i_ean = ""
            Call Log("X", "ItemEAN not found!", "ItemEAN")
        End If

        Return i_ean

    End Function

    Public Function CheckAndExport_invoice() As Integer

        Try

            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = cmp.GetBusinessObject(BoObjectTypes.BoRecordset)

            oRecordSet.DoQuery("SELECT DocEntry FROM OINV WHERE U_EDI_BERICHT = 'Ja' AND U_EDI_EXPORT = 'Ja' AND (U_EDI_INV_EXP is NUll OR U_EDI_INV_EXP = 'Nee')")

            If oRecordSet.RecordCount < 1 Then

                Call Log("V", "No Invoices note found to export!", "CheckAndExport_invoice")

            ElseIf oRecordSet.RecordCount > 0 Then

                Dim i As Integer
                oRecordSet.MoveFirst()
                For i = 1 To oRecordSet.RecordCount
                    Create_invoice_file(oRecordSet.Fields.Item(0).Value)
                    oRecordSet.MoveNext()
                Next

            End If

            oRecordSet = Nothing

        Catch ex As Exception

            Call Log("X", ex.Message, "CheckAndExport_invoice")

        End Try

        Return 0

    End Function

    Public Function Create_invoice_file(ByVal DocEntry As String) As Integer

        Dim oInvoice As SAPbobsCOM.Documents
        oInvoice = cmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

        Dim blnFind As Boolean = oInvoice.GetByKey(DocEntry)

        If blnFind = True Then

            Try

                Using wf As StreamWriter = New StreamWriter(strInvoicePath & "\" & "INV_" & oInvoice.DocNum & ".DAT")

                    Dim strHeader As String
                    Dim strLine As String

                    Dim strAlcHeader As String
                    Dim strAlcLine As String = ""

                    strHeader = ""
                    strAlcHeader = ""

                    With oInvoice.UserFields.Fields

                        FK_K_EANCODE = ""
                        FK_FAKTEST = ""
                        FK_KNAAM = ""
                        FK_F_SOORT = ""
                        FK_FAKT_NUM = ""
                        FK_FAKT_DATUM = ""
                        FK_AFL_DATUM = ""
                        FK_RFFIV = ""
                        FK_K_ORDERNR = ""
                        FK_K_ORDDAT = ""
                        FK_PAKBONNR = ""
                        FK_RFFCDN = ""
                        FK_RFFALO = ""
                        FK_RFFVN = ""
                        FK_RFFVNDAT = ""
                        FK_NAD_BY = ""
                        FK_A_EANCODE = ""
                        FK_F_EANCODE = ""
                        FK_NAD_SF = ""
                        FK_NAD_SU = ""
                        FK_NAD_UC = ""
                        FK_NAD_PE = ""
                        FK_OBNUMMER = ""
                        FK_ACT = ""
                        FK_CUX = ""
                        FK_DAGEN = ""
                        FK_KORTPERC = ""
                        FK_KORTBEDR = ""
                        FK_ONTVANGER = ""

                        'code
                        FK_K_EANCODE = Format_field(.Item("U_K_EANCODE").Value, 13, "STRING")

                        If .Item("U_TEST").Value = "J" Then
                            FK_FAKTEST = Format_field("1", 1, "STRING")
                        Else
                            FK_FAKTEST = Format_field(" ", 1, "STRING")
                        End If

                        FK_KNAAM = Format_field(.Item("U_KNAAM").Value, 14, "STRING")

                        If Len(Trim(.Item("U_F_SOORT").Value)) > 0 Then
                            FK_F_SOORT = Format_field(.Item("U_F_SOORT").Value, 3, "STRING")
                        Else
                            FK_F_SOORT = Format_field("380", 3, "STRING")
                        End If

                        FK_FAKT_NUM = Format_field(oInvoice.DocNum, 12, "STRING")

                        FK_FAKT_DATUM = Format_field(Format(Convert.ToDateTime(oInvoice.DocDate), "yyyyMMdd"), 8, "STRING")
                        Call Log("D", .Item("U_AFL_DATUM").Value, "Create_invoice_file")

                        '' FK_AFL_DATUM = Format_field(Format(CDate(.Item("U_AFL_DATUM").Value), "yyyyMMdd").ToString, 8, "STRING")
                        FK_AFL_DATUM = Format_field(Format(Convert.ToDateTime(.Item("U_AFL_DATUM").Value), "yyyyMMdd").ToString, 8, "STRING")
                        FK_RFFIV = Format_field(.Item("U_RFFIV").Value, 35, "STRING")
                        FK_K_ORDERNR = Format_field(.Item("U_K_ORDERNR").Value, 35, "STRING")
                        FK_K_ORDDAT = Format_field(.Item("U_K_ORDDAT").Value, 8, "STRING")
                        FK_PAKBONNR = Format_field(.Item("U_PAKBONNR").Value, 35, "STRING")
                        FK_RFFCDN = Format_field(.Item("U_RFFCDN").Value, 35, "STRING")
                        FK_RFFALO = Format_field(.Item("U_RFFALO").Value, 35, "STRING")
                        FK_RFFVN = Format_field(.Item("U_RFFVN").Value, 35, "STRING")
                        FK_RFFVNDAT = Format_field(.Item("U_RFFVNDAT").Value, 8, "STRING")
                        FK_NAD_BY = Format_field(.Item("U_NAD_BY").Value, 13, "STRING")
                        FK_A_EANCODE = Format_field(oInvoice.ShipToCode, 13, "STRING") ''Format_field(.Item("U_A_EANCODE").Value, 13, "STRING")
                        FK_F_EANCODE = Format_field(oInvoice.PayToCode, 13, "STRING") ''Format_field(.Item("U_F_EANCODE").Value, 13, "STRING")
                        FK_NAD_SF = Format_field(.Item("U_NAD_SF").Value, 13, "STRING")
                        FK_NAD_SU = Format_field(.Item("U_NAD_SU").Value, 13, "STRING")
                        FK_NAD_UC = Format_field(.Item("U_NAD_UC").Value, 13, "STRING")
                        FK_NAD_PE = Format_field(.Item("U_NAD_PE").Value, 13, "STRING")
                        FK_OBNUMMER = Format_field(CheckKoperAdres(.Item("U_K_EANCODE").Value), 15, "STRING")

                        FK_ACT = Format_field(.Item("U_ACT").Value, 1, "STRING")
                        FK_CUX = Format_field(oInvoice.DocCurrency, 3, "STRING")
                        FK_DAGEN = Format_field(.Item("U_DAGEN").Value, 3, "STRING")
                        FK_KORTPERC = Format_field(.Item("U_KORTPERC").Value, 8, "STRING")
                        FK_KORTBEDR = Format_field(.Item("U_KORTBEDR").Value, 9, "STRING")
                        FK_ONTVANGER = Format_field(.Item("U_ONTVANGER").Value, 13, "STRING")

                        strHeader = FK_K_EANCODE & _
                        FK_FAKTEST & _
                        FK_KNAAM & _
                        FK_F_SOORT & _
                        FK_FAKT_NUM & _
                        FK_FAKT_DATUM & _
                        FK_AFL_DATUM & _
                        FK_RFFIV & _
                        FK_K_ORDERNR & _
                        FK_K_ORDDAT & _
                        FK_PAKBONNR & _
                        FK_RFFCDN & _
                        FK_RFFALO & _
                        FK_RFFVN & _
                        FK_RFFVNDAT & _
                        FK_NAD_BY & _
                        FK_A_EANCODE & _
                        FK_F_EANCODE & _
                        FK_NAD_SF & _
                        FK_NAD_SU & _
                        FK_NAD_UC & _
                        FK_NAD_PE & _
                        FK_OBNUMMER & _
                        FK_ACT & _
                        FK_CUX & _
                        FK_DAGEN & _
                        FK_KORTPERC & _
                        FK_KORTBEDR & _
                        FK_ONTVANGER

                        wf.WriteLine("0" & strHeader)

                        AK_SOORT = ""
                        AK_QUAL = ""
                        AK_BEDRAG = ""
                        AK_BTWSOORT = ""
                        AK_FOOTMOA = ""
                        AK_NOTINCALC = ""

                        If oInvoice.DiscountPercent > 0 Then
                            AK_SOORT = Format_field("C", 1, "STRING")
                        Else
                            AK_SOORT = Format_field("A", 1, "STRING")
                        End If

                        AK_QUAL = Format_field(.Item("U_QUAL").Value, 3, "STRING")
                        AK_BEDRAG = Format_field(oInvoice.TotalDiscount, 9, "STRING")

                        AK_BTWSOORT = Format_field(.Item("U_BTWSOORT").Value, 1, "STRING")
                        AK_FOOTMOA = Format_field(.Item("U_FOOTMOA").Value, 1, "STRING")
                        AK_NOTINCALC = Format_field(.Item("U_NOTINCALC").Value, 1, "STRING")

                    End With

                    strAlcHeader = AK_SOORT & _
                    AK_QUAL & _
                    AK_BEDRAG & _
                    AK_BTWSOORT & _
                    AK_FOOTMOA & _
                    AK_NOTINCALC

                    If AK_SOORT = "C" Then wf.WriteLine("1" & strAlcHeader)

                    Dim i As Integer

                    For i = 0 To oInvoice.Lines.Count - 1
                        oInvoice.Lines.SetCurrentLine(i)

                        strLine = ""

                        With oInvoice.Lines.UserFields.Fields

                            FR_DEUAC = ""
                            FR_DEARTNR = ""
                            FR_DEARTOM = ""
                            FR_AANTAL = ""
                            FR_FAANTAL = ""
                            FR_ARTEENHEID = ""
                            FR_FEENHEID = ""
                            FR_NETTOBEDR = ""
                            FR_PRIJS = ""
                            FR_FREKEN = ""
                            FR_BTWSOORT = ""
                            FR_PV = ""
                            FR_ORDER = ""
                            FR_REGELID = ""
                            FR_INVO = ""
                            FR_DESA = ""
                            FR_PRIAAA = ""
                            FR_PIAPB = ""

                            FR_DEUAC = Format_field(.Item("U_DEUAC").Value, 14, "STRING")
                            FR_DEARTNR = Format_field("---------", 9, "STRING")
                            FR_DEARTOM = Format_field(oInvoice.Lines.ItemDescription, 70, "STRING")
                            FR_AANTAL = Format_field(oInvoice.Lines.Quantity, 5, "STRING")
                            FR_FAANTAL = Format_field(oInvoice.Lines.Quantity, 9, "STRING")
                            FR_ARTEENHEID = Format_field(.Item("U_ARTEENHEID").Value, 3, "STRING")
                            FR_FEENHEID = Format_field(.Item("U_FEENHEID").Value, 3, "STRING")
                            FR_NETTOBEDR = Format_field(oInvoice.Lines.LineTotal, 11, "STRING")
                            FR_PRIJS = Format_field(oInvoice.Lines.Price, 10, "STRING")
                            FR_FREKEN = Format_field(.Item("U_FREKEN").Value, 9, "STRING")

                            Select Case Trim(oInvoice.Lines.VatGroup)
                                Case Is = "A0"
                                    FR_BTWSOORT = Format_field("0", 1, "STRING")
                                Case Is = "A1"
                                    FR_BTWSOORT = Format_field("L", 1, "STRING")
                                Case Is = "A2"
                                    FR_BTWSOORT = Format_field("H", 1, "STRING")
                                Case Else
                                    FR_BTWSOORT = Format_field("9", 1, "STRING")
                            End Select

                            FR_PV = Format_field(.Item("U_PV").Value, 10, "STRING")
                            FR_ORDER = Format_field(.Item("U_ORDER").Value, 35, "STRING")
                            FR_REGELID = Format_field(oInvoice.Lines.LineNum, 6, "STRING")
                            FR_INVO = Format_field(.Item("U_INVO").Value, 35, "STRING")
                            FR_DESA = Format_field(.Item("U_DESA").Value, 35, "STRING")
                            FR_PRIAAA = Format_field(.Item("U_PRIAAA").Value, 10, "STRING")
                            FR_PIAPB = Format_field(oInvoice.Lines.SupplierCatNum, 9, "STRING") ''Format_field(.Item("U_PIAPB").Value, 20, "STRING")

                            strLine = FR_DEUAC & _
                            FR_DEARTNR & _
                            FR_DEARTOM & _
                            FR_AANTAL & _
                            FR_FAANTAL & _
                            FR_ARTEENHEID & _
                            FR_FEENHEID & _
                            FR_NETTOBEDR & _
                            FR_PRIJS & _
                            FR_FREKEN & _
                            FR_BTWSOORT & _
                            FR_PV & _
                            FR_ORDER & _
                            FR_REGELID & _
                            FR_INVO & _
                            FR_DESA & _
                            FR_PRIAAA & _
                            FR_PIAPB

                            If oInvoice.Lines.TreeType <> BoItemTreeTypes.iIngredient Then wf.WriteLine("2" & strLine)



                        End With



                    Next

                    wf.Close()

                    oInvoice.UserFields.Fields.Item("U_EDI_INV_EXP").Value = "Ja"
                    oInvoice.UserFields.Fields.Item("U_EDI_INVEXP_TIJD").Value = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")

                    oInvoice.Update()

                    If iSendNotification = 1 Then Mail_to_DL_receiver(oInvoice.DocEntry, oInvoice.DocDate, oInvoice.DocNum)

                    Call Log("V", "Invoice file created!", "Create_invoice_file")

                End Using

            Catch ex As Exception

                File.Delete(strInvoicePath & "\" & "INV_" & oInvoice.DocNum & ".DAT")
                Call Log("X", ex.Message, "Create_invoice_file")
            End Try

        Else
            Call Log("X", "Invoice " & DocEntry & " not found!", "Create_invoice_file")
        End If

        Return 0

    End Function

    Private Function Format_field(ByVal FieldValue As String, ByVal FieldLen As Integer, ByVal Type As String) As String
        'Call Log("D", FieldValue, FieldLen.ToString())
        Dim NewValue As String

        Select Case Type
            Case "STRING"

                If Len(FieldValue) > 0 Then
                    NewValue = Trim(FieldValue) & Mid(Space(FieldLen), 1, FieldLen - Len(Trim(Left(FieldValue, FieldLen))))
                Else
                    NewValue = Space(FieldLen)
                End If

            Case Else
                NewValue = Space(FieldLen)
        End Select


        Return NewValue

    End Function

    Public Function Mail_to_DL_receiver(sDocEntry As Integer, sDocDate As String, iDocNumber As Integer) As Boolean

        Try
            Using mailMsg As New Mail.MailMessage()

                Dim SmtpMail As New Mail.SmtpClient() With {.Host = sSmtp, .Port = iSmtpPort}

                If bSmtpUserSecurity = True Then
                    SmtpMail.Credentials = New System.Net.NetworkCredential(sSmtpUser, sSmtpPassword)
                End If

                With mailMsg
                    .From = New System.Net.Mail.MailAddress(sSenderEmail, sSenderName)
                    .To.Add(sDeliveryMailTo)
                    .Subject = String.Format("Delivery {0} exported", sDocEntry)

                    Using rFile As New StreamReader(applicationPath & "\email_d.txt")
                        Dim sBody As String
                        sBody = rFile.ReadToEnd()
                        sBody = Replace(sBody, "::NAME::", sDeliveryMailToFullName)
                        sBody = Replace(sBody, "::DOCENTRY::", sDocEntry)
                        sBody = Replace(sBody, "::DOCDATE::", sDocDate)
                        sBody = Replace(sBody, "::DOCNUM::", CStr(iDocNumber))
                        .Body = sBody
                        rFile.Close()
                    End Using

                End With

                SmtpMail.Send(mailMsg)

            End Using

            Call Log("V", "Delivery notification sent!", "Mail_to_DL_receiver")

            Return True

        Catch ex As Exception

            Call Log("X", "Delivery notification was not sent!", "Mail_to_DL_receiver")

            Return False

        End Try

    End Function

    Public Function Mail_to_IN_receiver(sDocEntry As Integer, sDocDate As String, iDocNumber As Integer) As Boolean

        Try
            Using mailMsg As New Mail.MailMessage()

                Dim SmtpMail As New Mail.SmtpClient() With {.Host = sSmtp, .Port = iSmtpPort}

                If bSmtpUserSecurity = True Then
                    SmtpMail.Credentials = New System.Net.NetworkCredential(sSmtpUser, sSmtpPassword)
                End If

                With mailMsg
                    .From = New System.Net.Mail.MailAddress(sSenderEmail, sSenderName)
                    .To.Add(sInvoiceMailTo)
                    .Subject = String.Format("Invoice {0} imported", sDocEntry)

                    Using rFile As New StreamReader(applicationPath & "\email_i.txt")
                        Dim sBody As String
                        sBody = rFile.ReadToEnd()
                        sBody = Replace(sBody, "::NAME::", sInvoiceMailToFullName)
                        sBody = Replace(sBody, "::DOCENTRY::", sDocEntry)
                        sBody = Replace(sBody, "::DOCDATE::", sDocDate)
                        sBody = Replace(sBody, "::DOCNUM::", CStr(iDocNumber))
                        .Body = sBody
                        rFile.Close()
                    End Using

                End With

                SmtpMail.Send(mailMsg)

            End Using

            Call Log("V", "Invoice notification sent!", "Mail_to_IN_receiver")

            Return True

        Catch ex As Exception

            Call Log("X", "Invoice notification was not sent!", "Mail_to_IN_receiver")

            Return False

        End Try

    End Function

End Class
