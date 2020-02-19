Imports System.Net
Imports System.Threading
Imports System.Web
Imports System.Net.Mail
Imports System.IO
Imports System.Xml
Imports System.Text
Imports Newtonsoft.Json
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Schema
Imports Newtonsoft.Json.Linq



Public Class frmMain

    '
    ' Added to reset textbox to prevent out of memory
    '
    Private myAppendLineCountMax As Integer = 100
    Private myAppendLineCount As Integer = 0

    Private myTimerInterval As Integer = 500
    'myDatabaseConnectionString replaced by gDatabaseConnectionString
    'Private myDatabaseConnectionString As String = "DATA SOURCE=DEVMSG;INITIAL CATALOG=IPIMSGSVR;USER ID=sa;PASSWORD=b4ker1"       'Current Dev Default - use this normally
    'Private myDatabaseConnectionString As String = "DATA SOURCE=DEVMAIL;INITIAL CATALOG=IPIMSGSVR;USER ID=sa;PASSWORD=b4ker1"       'Current Dev Default - use this normally
    'Private myDatabaseConnectionString As String = "DATA SOURCE=DEVSQL2008;INITIAL CATALOG=IPIMSGSVR;USER ID=sa;PASSWORD=b4ker1"   'Old Dev Default
    'Private myDatabaseConnectionString As String = "DATA SOURCE=AMSMSG1;INITIAL CATALOG=IPIMSGSVR;USER ID=sa;PASSWORD=b4ker1"      'Production Default

    Private gErrorMessage As String = ""
    Private gEnvironment As String = ""
    Public gDatabaseConnectionString As String = ""
    Public gEmailSmtpHost1 As String = ""
    Public gEmailSmtpHost2 As String = ""
    Public gEmailFromAddress As String = ""
    Public gDBErrorEmails As New DataTable
    Public gTimelastErrorEmail As Date
    Public gWaitTimeBetweenErrorEmails As Integer = 5 'minutes
    Public gIMSGUsersLastUpdate As DateTime = Nothing
    Public gIntelliMessageWIFIMSGUCreated As Boolean = False
    Public gAbacusWIFIMSGUCleared As Boolean = False
    Public gReadWIFIMSGUCreatedTime As DateTime = Nothing
    Public gReadWIFIMSGUClearedTime As DateTime = Nothing
    Public gRUN_HHMISS As String = ""
    Public gRUN_YYYYMMDD As String = ""
    Public gCureatrAlert As String = ""
    Private gCureatrAlertCount As Integer = 0
    Private gCureatrAlertSuccess As Boolean = 1
    Private gCureatrAlertLastEmailSent As DateTime = Nothing
    Private gURIBase As String = ""
    Private gURIBase1 As String = ""
    Private gURIBase2 As String = ""
    Private gHeartBeatURIBase As String = ""
    Private gPasses As Integer = 0

    Private myDS As DataSet = Nothing
    Private myDS1 As DataSet = Nothing
    Private myWIFIMSGU As DataSet = Nothing
    Private myHarkDS As DataSet = Nothing
    Private myImsgDiscoDS As DataSet = Nothing
    Private myImsgPSChangeDS As DataSet = Nothing
    Private mySqlConn As SqlClient.SqlConnection = Nothing
    Private myStop As Boolean = False
    Private myTimeOut As Boolean = False
    Private myExceptionText As String = ""
    Private myResponseReceived As Boolean = False
    Private mySendEmail As Boolean = False
    Private myEmailSentTime As DateTime = Now
    Private myHeartbeatCheckTime As DateTime = Now
    Private myHeartbeatcheckTime1 As DateTime = Now.AddMinutes(2)
    Private myHeartbeatcheckTime2 As DateTime = Now.AddMinutes(4)
    Private myActivity As Boolean = False
    Private myAccounts As Boolean = False
    Private myIMSGUsers As Boolean = False
    Private gSSOInterfaceFailing As String = ""
    Private gLoop As Int16 = 0

    '  Private gApiKey As String = ""
    Private gSSO_URL As String = ""
    Private gSSO_IM_Token As String = ""
    Private gSSO_AM_Token As String = ""
    Private gSSO_IM_APIKey As String = ""
    Private gSSO_AM_APIKey As String = ""
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        'Set everything to start on load!
        'Me.tmrMain.Interval = myTimerInterval
        'Me.tmrMain.Enabled = True

        If Me.tmrMain.Enabled = False Then
            Me.btnStart.Enabled = True
            Me.btnStop.Enabled = False
        Else
            Me.btnStart.Enabled = False
            Me.btnStop.Enabled = True
        End If

        txtEnvironment.Focus()

        'set up error message data table
        Dim auto As DataColumn = New DataColumn
        auto.DataType = System.Type.GetType("System.Int32")
        With auto
            .AutoIncrement = True
            .AutoIncrementSeed = 1
            .AutoIncrementStep = 1
        End With
        gDBErrorEmails.Columns.Add("auto")
        gDBErrorEmails.Columns.Add("ErrorMessage", Type.GetType("System.String"))

    End Sub

    Private Sub btnStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStart.Click

        gEnvironment = UCase(txtEnvironment.Text)
        If gEnvironment <> "P" And gEnvironment <> "T" And gEnvironment <> "D" Then
            MsgBox("Prod, Dev, Test must be P or T or D", MsgBoxStyle.Information)
            txtEnvironment.Focus()
            Exit Sub
        End If

        'Dim myURIBase As String = "http://63.97.58.99/xml.tbred?xmlrequest=METHOD+WIMTC820+"       'Dev outward facing IP - Use normally
        'Dim myURIBase As String = "http://10.214.34.75/xml.tbred?xmlrequest=METHOD+WIMTC820+"      'dev inside IP
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=METHOD+WIMTC820+"      'BamBam\test IP

        'Dim myURIBase As String = "http://63.97.58.98/xml.tbred?xmlrequest=METHOD+WIMTC820+"        'prod PCAXML outward facing IP - Use normally
        'Dim myURIBase As String = "http://10.214.34.76/xml.tbred?xmlrequest=METHOD+WIMTC820+"      'dev inside IP
        'Dim myURIBase As String = "http://10.200.2.20/xml.tbred?xmlrequest=METHOD+WIMTC820+"      'atprod2 IP - USE after Phred2 conversion

        gURIBase = ""
        gDatabaseConnectionString = ""
        If gEnvironment = "D" Then 'per Z 9/7/18
            MsgBox(" Not valid selection")
            Exit Sub
            'gUriBase = "http://10.200.2.70/devmxml.tbred?"  'BamBam\devm
            gURIBase = "http://10.200.2.70/devaxml.tbred?"  'BamBam\devm
            gDatabaseConnectionString = "DATA SOURCE=DEVMSG;INITIAL CATALOG=IPIMSGSVR;USER ID=sa;PASSWORD=b4ker1"
            'Setup for failover
            gEmailSmtpHost1 = "DEVMSG"
            gEmailSmtpHost2 = "POSTMAN"
            gEmailFromAddress = "administrator@dev.intellimsg.net"
            gCureatrAlert = My.Settings.CureatrAlertTEST
        End If
        If gEnvironment = "T" Then
            gURIBase = "http://10.200.2.70/testaxml.tbred?"  'BamBam\test IP Used by Abacus Interface
            gURIBase1 = "http://10.200.2.70/testxml.tbred?"  'BamBam\test IP Used by Java
            gURIBase2 = "http://10.200.2.70/testixml.tbred?" 'BamBam\test IP Used by IMSG System Interface
            gDatabaseConnectionString = "DATA SOURCE=QA-AMSMSG;INITIAL CATALOG=IPIMSGSVR;USER ID=sa;PASSWORD=b4ker1"
            'gDatabaseConnectionString = "DATA SOURCE=IPIMSG;INITIAL CATALOG=IPIMSGSVR;USER ID=sa;PASSWORD=b4ker1"
            'Setup for failover
            gEmailSmtpHost1 = "QA-AMSMSG"
            gEmailSmtpHost2 = "QA-AMSMSG"
            gEmailFromAddress = "administrator@test.intellimsg.net"
            gCureatrAlert = My.Settings.CureatrAlertTEST
        End If
        If gEnvironment = "P" Then
            'AMSMSG1 = LV-MMDatabase = Alisa
            gURIBase = "http://10.200.2.20/axml.tbred?"      'atprod2 IP Used by Abacus Interface
            gURIBase1 = "http://10.200.2.20/xml.tbred?"      'atprod2 IP Used by Java
            gURIBase2 = "http://10.200.2.20/ixml.tbred?"     'atprod2 IP Used by IMSG System Interface
            'gDatabaseConnectionString = "DATA SOURCE=AMSMsg.amsipi.alert.am;INITIAL CATALOG=IPIMSGSVR;USER ID=sa;PASSWORD=b4ker1"
            gDatabaseConnectionString = "DATA SOURCE=AMSMSG;INITIAL CATALOG=IPIMSGSVR;USER ID=mmapps;PASSWORD=Bb4ker1"
            'Setup for failover
            gEmailSmtpHost1 = "AMSMSG1"
            gEmailSmtpHost2 = "AMSMSG2"
            gEmailFromAddress = "administrator@intellimsg.net"
            gCureatrAlert = My.Settings.CureatrAlertPROD
        End If

        'SendEmail("test email from ABACUS Transaction.  TEST ONLY")

        If Me.tmrMain.Enabled = False Then
            AppendText("Process Started")
            Me.tmrMain.Interval = myTimerInterval
            Me.tmrMain.Enabled = True
            Me.btnStart.Enabled = False
            Me.btnStop.Enabled = True
        End If

        gDBErrorEmails.Clear() 'clear old error emails.  In desktop log if needed


    End Sub

    Private Sub tmrMain_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrMain.Tick

        Me.tmrMain.Enabled = False

        DoTimer()
        If myStop = True Then
            Me.tmrMain.Enabled = False
        Else
            Me.tmrMain.Interval = myTimerInterval
            Me.tmrMain.Enabled = True
        End If

    End Sub

    Private Sub DoTimer()
        Dim Debug As Integer = 0    'Turn subroutine messages off/on
        Dim terminate As Boolean = True
        Dim records As Integer = 0
        Dim drecords As Integer = 0
        Dim mySql As String = ""
        Dim myFlag As Boolean = False
        Dim myCount As Integer = 0
        Dim myError As String = ""
        Dim begReturnMsg As String = ""
        Dim processWaitFor As Integer = 0
        Dim processShouldTerminate As Boolean = False
        Dim updReturnMsg As String = ""
        Dim myBeginDS As New DataSet
        Dim myUpdDS As New DataSet
        Dim myEndDS As New DataSet
        Dim myHarkNumbers As Boolean = False
        Dim myIntelliMsgDisconnects As Boolean = False
        Dim myIntelliMsgPSChanges As Boolean = False
        Dim i As Integer = 0
        Dim intExecute As Integer = 0

        'WriteToErrorEmail("This is test1 DCB")
        'WriteToErrorEmail("This is test2 DCB")
        'WriteToErrorEmail("This is test3 DCB")
        'Do While 1 = 1
        '    'AppendText("ReadAppUserPasswordChange")
        '    'ReadAppUserPasswordChange()
        '    AppendText("Create_Update_SSO_Accounts")
        ' Create_Update_SSO_Accounts()

        '    Application.DoEvents()
        '    AppendText("ProcessErrorMessages")
        '    ProcessErrorMessages()
        '    i = 0
        '    For i = 1 To 6
        '        System.Threading.Thread.Sleep(1000)
        '        Application.DoEvents()
        '    Next
        'Loop
        'test ===============================================================
        'CureatrAlert()
        'AppendText("Process Cureatr Alert " & Now)
        'processWaitFor = 5000 * processWaitFor
        'If myActivity = False Then
        '    'processWaitFor = 10000
        '    processWaitFor = 15000
        'End If
        'myTimerInterval = processWaitFor
        'gPasses = gPasses + 1
        'If (gPasses Mod 1000) = 0 Then
        '    Me.txtText.Text = ""
        '    Me.txtText.AppendText("Passes " & gPasses.ToString & vbCrLf)
        'End If
        'Exit Sub
        'test ===============================================================

        mySql = "exec ProcessBegin 'AbacusInterface'"
        myFlag = FillSqlDataSet(mySql, myBeginDS, myError)
        If myFlag = True Then
            myFlag = False
            terminate = True
            If myBeginDS.Tables.Count = 1 Then
                begReturnMsg = myBeginDS.Tables(0).Rows(0).Item("ReturnMsg").ToString.Trim
                If begReturnMsg = "." Then
                    myFlag = True
                    terminate = False
                Else
                    myFlag = False
                    gErrorMessage = begReturnMsg
                    terminate = True
                End If
            End If
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("AbacusInterface ProcessBegin" & vbCrLf)
        End If

        mySql = "exec ProcessBegin 'XMLHeartBeatAbacusInterface'"
        myFlag = FillSqlDataSet(mySql, myBeginDS, myError)
        If myFlag = True Then
            myFlag = False
            terminate = True
            If myBeginDS.Tables.Count = 1 Then
                begReturnMsg = myBeginDS.Tables(0).Rows(0).Item("ReturnMsg").ToString.Trim
                If begReturnMsg = "." Then
                    myFlag = True
                    terminate = False
                Else
                    myFlag = False
                    gErrorMessage = begReturnMsg
                    terminate = True
                End If
            End If
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("XMLHeartBeatAbacusInterface ProcessBegin" & vbCrLf)
        End If

        mySql = "exec ProcessBegin 'XMLHeartBeatJava'"
        myFlag = FillSqlDataSet(mySql, myBeginDS, myError)
        If myFlag = True Then
            myFlag = False
            terminate = True
            If myBeginDS.Tables.Count = 1 Then
                begReturnMsg = myBeginDS.Tables(0).Rows(0).Item("ReturnMsg").ToString.Trim
                If begReturnMsg = "." Then
                    myFlag = True
                    terminate = False
                Else
                    myFlag = False
                    gErrorMessage = begReturnMsg
                    terminate = True
                End If
            End If
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("XMLHeartBeatJava ProcessBegin" & vbCrLf)
        End If

        mySql = "exec ProcessBegin 'XMLHeartBeatIMSGSystemInterface'"
        myFlag = FillSqlDataSet(mySql, myBeginDS, myError)
        If myFlag = True Then
            myFlag = False
            terminate = True
            If myBeginDS.Tables.Count = 1 Then
                begReturnMsg = myBeginDS.Tables(0).Rows(0).Item("ReturnMsg").ToString.Trim
                If begReturnMsg = "." Then
                    myFlag = True
                    terminate = False
                Else
                    myFlag = False
                    gErrorMessage = begReturnMsg
                    terminate = True
                End If
            End If
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("XMLHeartBeatIMSGSystemInterface ProcessBegin" & vbCrLf)
        End If

        Me.txtText.AppendText("Begin Cycle" & vbCrLf)

        myActivity = False

        ReadAccountsToAdd(myDS)

        If myAccounts = True Then
            SendAccountsToAdd(myDS, records)
            myActivity = True
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("ReadAccountsToAdd" & vbCrLf)
        End If

        ReadAccountsToDisconnect(myDS)

        If myAccounts = True Then
            SendAccountsToDisconnect(myDS, records)
            myActivity = True
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("ReadAccountsToDisconnect" & vbCrLf)
        End If

        ReadHarkNumbers(myHarkDS)

        If Not myHarkDS Is Nothing Then
            SendHarkNumbersToAdd(myHarkDS, records)
            myActivity = True
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("ReadHarkNumbers" & vbCrLf)
        End If

        CheckForIntelliMsgDisconnects(myIntelliMsgDisconnects)

        If myIntelliMsgDisconnects = True Then
            ReadIntelliMsgDisconnects(myImsgDiscoDS)

            If Not myImsgDiscoDS Is Nothing Then
                SendIntelliMsgDisconnects(myImsgDiscoDS, drecords)
                myActivity = True
            End If
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("CheckForIntelliMsgDisconnects" & vbCrLf)
        End If

        CheckForIntelliMsgPSChange(myIntelliMsgPSChanges)

        If myIntelliMsgPSChanges = True Then
            ReadIntelliMsgPSChange(myImsgPSChangeDS)

            If Not myImsgPSChangeDS Is Nothing Then
                SendIntelliMsgPSChanges(myImsgPSChangeDS, drecords)
                myActivity = True
            End If
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("ReadIntelliMsgPSChange" & vbCrLf)
        End If

        ReadCapcodeGroup()
        If Debug = 1 Then
            Me.txtText.AppendText("ReadCapcodeGroup" & vbCrLf)
        End If

        ReadCapcodeGroupAccounts()
        If Debug = 1 Then
            Me.txtText.AppendText("ReadCapcodeGroupAccounts" & vbCrLf)
        End If

        ReadChannelCode()
        If Debug = 1 Then
            Me.txtText.AppendText("ReadChannelCode" & vbCrLf)
        End If

        ReadCueConfiguration()
        If Debug = 1 Then
            Me.txtText.AppendText("ReadCueConfiguration" & vbCrLf)
        End If

        AppendText("ReadCueEquipment")
        ReadCueEquipment()
        If Debug = 1 Then
            Me.txtText.AppendText("ReadCueEquipment" & vbCrLf)
        End If

        AppendText("PurgeOldCueLogFiles")
        PurgeOldCueLogFiles()

        AppendText("ReadPagerGCTransfer")
        ReadPagerGCTransfer()

        AppendText("ReadPagerConfigurationTransfer")
        ReadPagerConfigurationTransfer()
        If Debug = 1 Then
            Me.txtText.AppendText("ReadPagerConfigurationTransfer" & vbCrLf)
        End If
        AppendText("ReadAppUserPasswordChange")
        ReadAppUserPasswordChange()

        AppendText("ReadWIFCLFWD")
        ReadWIFCLFWD()

        AppendText("ReadWIFMCDOO")
        ReadWIFMCDOO()


        myWIFIMSGU = Nothing

        If Now.Date.ToString > gIMSGUsersLastUpdate.Date.ToString Then
            If Now.TimeOfDay.ToString > #3:00:00 AM# Then
                If gIntelliMessageWIFIMSGUCreated = False Then
                    If gReadWIFIMSGUCreatedTime < Now Then
                        AppendText("ReadWIFIMSGUCreated")
                        ReadWIFIMSGUCreated()
                        If gIntelliMessageWIFIMSGUCreated = False Then
                            gReadWIFIMSGUCreatedTime = Now.AddMinutes(1)
                        End If
                    End If
                End If
                If gAbacusWIFIMSGUCleared = False Then
                    If gReadWIFIMSGUClearedTime < Now Then
                        AppendText("ReadWIFIMSGUCleared")
                        ReadWIFIMSGUCleared()
                        If gAbacusWIFIMSGUCleared = False Then
                            gReadWIFIMSGUClearedTime = Now.AddMinutes(1)
                        End If
                    End If
                End If

                If gIntelliMessageWIFIMSGUCreated = True And gAbacusWIFIMSGUCleared = True Then
                    AppendText("ReadWIFIMSGUToCopy")
                    ReadWIFIMSGUToCopy(myWIFIMSGU)
                    If myIMSGUsers = True Then
                        AppendText("SendWIFIMSGUToAdd")
                        SendWIFIMSGUToAdd(myWIFIMSGU, records)
                        myActivity = True
                    Else
                        AppendText("UpdateWIFIMSGUCreated")
                        UpdateWIFIMSGUCreated()
                    End If
                End If
            End If
        End If

        AppendText("Create_Update_SSO_Accounts")
        Create_Update_SSO_Accounts()

        AppendText("ProcessErrorMessages")
        ProcessErrorMessages()

        'no more than 1 per minute
        If gLoop = 0 Or gLoop > 14 Then
            gLoop = 0
            AppendText("Process Cureatr Alert ")
            CureatrAlert()
        Else
            AppendText("Process Cureatr Alert - SKIP ")
        End If
        gLoop = gLoop + 1

        ''timing test
        'For i = 1 To 30
        '    System.Threading.Thread.Sleep(1000)
        '    Application.DoEvents()
        'Next


        If gEnvironment = "T" Or gEnvironment = "P" Then
            'If Now.TimeOfDay.ToString < #11:28:00 PM# Or Now.TimeOfDay.ToString > #11:50:00 PM# Then
            If Now.TimeOfDay.ToString < #11:20:00 PM# And Now.TimeOfDay.ToString > #12:10:00 AM# Then

                If myHeartbeatCheckTime < Now Then
                    gHeartBeatURIBase = gURIBase
                    CheckXMLHeartbeat(gHeartBeatURIBase)
                    processShouldTerminate = True
                    processWaitFor = 1
                    '
                    mySql = "exec ProcessUpdate 'XMLHeartBeatAbacusInterface', " + records.ToString()
                    myFlag = FillSqlDataSet(mySql, myUpdDS, myError)
                    If myFlag = True Then
                        myFlag = False
                        If myUpdDS.Tables.Count = 1 Then
                            updReturnMsg = myUpdDS.Tables(0).Rows(0).Item("ReturnMsg").ToString.Trim
                            If updReturnMsg = "." Then
                                processWaitFor = CInt(myUpdDS.Tables(0).Rows(0).Item(("ProcessWaitFor")))
                                processShouldTerminate = CBool(myUpdDS.Tables(0).Rows(0).Item("ProcessShouldTerminate"))
                            End If
                        End If
                    End If
                    If Debug = 1 Then
                        Me.txtText.AppendText("XMLHeartBeatAbacusInterface ProcessUpdate" & vbCrLf)
                    End If
                    myHeartbeatCheckTime = Now.AddMinutes(6)
                End If

                If myHeartbeatcheckTime1 < Now Then
                    gHeartBeatURIBase = gURIBase1
                    CheckXMLHeartbeat(gHeartBeatURIBase)
                    processShouldTerminate = True
                    processWaitFor = 1
                    '
                    mySql = "exec ProcessUpdate 'XMLHeartBeatJava', " + records.ToString()
                    myFlag = FillSqlDataSet(mySql, myUpdDS, myError)
                    If myFlag = True Then
                        myFlag = False
                        If myUpdDS.Tables.Count = 1 Then
                            updReturnMsg = myUpdDS.Tables(0).Rows(0).Item("ReturnMsg").ToString.Trim
                            If updReturnMsg = "." Then
                                processWaitFor = CInt(myUpdDS.Tables(0).Rows(0).Item(("ProcessWaitFor")))
                                processShouldTerminate = CBool(myUpdDS.Tables(0).Rows(0).Item("ProcessShouldTerminate"))
                            End If
                        End If
                    End If
                    If Debug = 1 Then
                        Me.txtText.AppendText("XMLHeartBeatJava ProcessUpdate" & vbCrLf)
                    End If
                    myHeartbeatcheckTime1 = Now.AddMinutes(6)
                End If

                If myHeartbeatcheckTime2 < Now Then
                    gHeartBeatURIBase = gURIBase2
                    CheckXMLHeartbeat(gHeartBeatURIBase)
                    processShouldTerminate = True
                    processWaitFor = 1
                    '
                    mySql = "exec ProcessUpdate 'XMLHeartBeatIMSGSystemInterface', " + records.ToString()
                    myFlag = FillSqlDataSet(mySql, myUpdDS, myError)
                    If myFlag = True Then
                        myFlag = False
                        If myUpdDS.Tables.Count = 1 Then
                            updReturnMsg = myUpdDS.Tables(0).Rows(0).Item("ReturnMsg").ToString.Trim
                            If updReturnMsg = "." Then
                                processWaitFor = CInt(myUpdDS.Tables(0).Rows(0).Item(("ProcessWaitFor")))
                                processShouldTerminate = CBool(myUpdDS.Tables(0).Rows(0).Item("ProcessShouldTerminate"))
                            End If
                        End If
                    End If
                    If Debug = 1 Then
                        Me.txtText.AppendText("XMLHeartBeatIMSGSystemInterface ProcessUpdate" & vbCrLf)
                    End If
                    myHeartbeatcheckTime2 = Now.AddMinutes(6)
                End If

            End If
        End If

        ''
        processShouldTerminate = True
        processWaitFor = 1
        '
        mySql = "exec ProcessUpdate 'AbacusInterface', " + records.ToString()
        myFlag = FillSqlDataSet(mySql, myUpdDS, myError)
        If myFlag = True Then
            myFlag = False
            If myUpdDS.Tables.Count = 1 Then
                updReturnMsg = myUpdDS.Tables(0).Rows(0).Item("ReturnMsg").ToString.Trim
                If updReturnMsg = "." Then
                    processWaitFor = CInt(myUpdDS.Tables(0).Rows(0).Item(("ProcessWaitFor")))
                    processShouldTerminate = CBool(myUpdDS.Tables(0).Rows(0).Item("ProcessShouldTerminate"))
                End If
            End If
        End If
        If Debug = 1 Then
            Me.txtText.AppendText("ProcessUpdate" & vbCrLf)
        End If

        records = 0

        If (processShouldTerminate) Then
            terminate = True
            myStop = True
        Else
            'processWaitFor = 1000 * processWaitFor
            processWaitFor = 5000 * processWaitFor
            If myActivity = False Then
                'processWaitFor = 10000
                processWaitFor = 15000
            End If
            myTimerInterval = processWaitFor
            gPasses = gPasses + 1
            If (gPasses Mod 1000) = 0 Then
                Me.txtText.Text = ""
                Me.txtText.AppendText("Passes " & gPasses.ToString & vbCrLf)
            End If
        End If

        If myStop = True Then
            mySql = "exec ProcessEnd 'AbacusInterface'"
            myFlag = FillSqlDataSet(mySql, myEndDS, myError)
            mySql = "exec ProcessEnd 'XMLHeartBeatAbacusInterface'"
            myFlag = FillSqlDataSet(mySql, myEndDS, myError)
            mySql = "exec ProcessEnd 'XMLHeartBeatJava'"
            myFlag = FillSqlDataSet(mySql, myEndDS, myError)
            mySql = "exec ProcessEnd 'XMLHeartBeatIMSGSystemInterface'"
            myFlag = FillSqlDataSet(mySql, myEndDS, myError)
        End If

    End Sub

    ' Function to convert passed XML data to dataset
    Public Function ConvertXMLToDataSet(ByVal xmlData As String) As DataSet
        Dim stream As StringReader = Nothing
        Dim reader As XmlTextReader = Nothing
        Try
            Dim xmlDS As New DataSet()
            stream = New StringReader(xmlData)
            ' Load the XmlTextReader from the stream
            reader = New XmlTextReader(stream)
            xmlDS.ReadXml(reader)
            Return xmlDS
        Catch
            Return Nothing
        Finally
            If reader IsNot Nothing Then
                reader.Close()
            End If
        End Try
    End Function
    ' Use this function to get XML string from a dataset
    ' Function to convert passed dataset to XML data
    Public Function ConvertDataSetToXML(ByVal xmlDS As DataSet) As String
        Dim stream As MemoryStream = Nothing
        Dim writer As XmlTextWriter = Nothing
        Try
            stream = New MemoryStream()
            ' Load the XmlTextReader from the stream
            writer = New XmlTextWriter(stream, Encoding.Unicode)
            ' Write to the file with the WriteXml method.
            xmlDS.WriteXml(writer)
            Dim count As Integer = CInt(stream.Length)
            Dim arr As Byte() = New Byte(count - 1) {}
            stream.Seek(0, SeekOrigin.Begin)
            stream.Read(arr, 0, count)
            Dim utf As New UnicodeEncoding()
            Return utf.GetString(arr).Trim()
        Catch
            Return [String].Empty
        Finally
            If writer IsNot Nothing Then
                writer.Close()
            End If
        End Try
    End Function

    Private Sub btnStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStop.Click
        If Me.tmrMain.Enabled = True Then
            AppendText("Stop Requested")
            myStop = True
            'Me.tmrMain.Enabled = False
            Me.btnStart.Enabled = True
            Me.btnStop.Enabled = False

        End If
    End Sub

    Friend Function FillSqlDataSet(ByVal pSql As String, _
                                    ByRef pDataSet As Data.DataSet, _
                                    ByRef pError As String, _
                                    Optional ByVal pConnectionString As String = "") As Boolean
        Dim myFlag As Boolean = False
        Dim mySqlAdapter As System.Data.SqlClient.SqlDataAdapter
        Dim myConnection As New System.Data.SqlClient.SqlConnection

        Try
            If pConnectionString = "" Then
                myConnection.ConnectionString = gDatabaseConnectionString
            Else
                myConnection.ConnectionString = pConnectionString
            End If
            If myConnection.State <> System.Data.ConnectionState.Open Then
                myConnection.Open()
            End If

            pError = ""
            If Not pDataSet Is Nothing Then
                pDataSet.Dispose()
            End If
            pDataSet = New System.Data.DataSet
            mySqlAdapter = Nothing
            mySqlAdapter = New System.Data.SqlClient.SqlDataAdapter(pSql, myConnection)
            mySqlAdapter.Fill(pDataSet)
            myFlag = True

        Catch SqlEx As SqlClient.SqlException
            myFlag = False
            pError = "SqlEx: " & SqlEx.Message

        Catch ex As Exception
            myFlag = False
            pError = "PgmEx: " & ex.Message

        End Try

        If Not myConnection Is Nothing Then
            If myConnection.State = System.Data.ConnectionState.Open Then
                myConnection.Close()
            End If
            myConnection.Dispose()
        End If

        FillSqlDataSet = myFlag

    End Function

    Private Sub ReadAccountsToAdd(ByRef pDS As DataSet)
        Dim mySql As String = ""
        Dim myFlag As Boolean = False

        mySql = "exec GetAccountsToAdd"
        myFlag = FillSqlDataSet(mySql, pDS, gErrorMessage)
        myAccounts = myFlag
    End Sub

    Private Sub SendAccountsToAdd(ByRef pDS As DataSet, ByRef pRecords As Integer)
        Dim myDataTable As DataTable = Nothing
        Dim mySendSuccess As Boolean = False
        Dim coCode As String = ""
        Dim acctNo As String = ""
        Dim billRateCode As String = ""
        Dim pagerNumber As String = ""
        Dim appUserId As String = ""
        Dim enhancedServOption As String = ""
        Dim administrator As String = ""
        Dim usertype As String = ""
        Dim billedthroughdate As String = ""
        Dim myURIData As String = ""

        'Dim myURIBase As String = "http://63.97.58.99/xml.tbred?xmlrequest=METHOD+WIMTC820+"       'Dev outward facing IP - Use normally
        'Dim myURIBase As String = "http://10.214.34.75/xml.tbred?xmlrequest=METHOD+WIMTC820+"      'dev inside IP
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=METHOD+WIMTC820+"      'BamBam\test IP

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMTC820+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""
        Dim holdURIResponse As String = ""
        Dim mySQL As String = ""
        Dim myData As DataTable = Nothing
        Dim myPhoneNumber As String = ""

        If pDS.Tables.Count > 0 Then
            myDataTable = pDS.Tables(0)
            For Each myRow As DataRow In myDataTable.Rows
                coCode = myRow.Item("CoCode").ToString
                acctNo = myRow.Item("AcctNo").ToString
                billRateCode = myRow.Item("BillRateCode").ToString
                pagerNumber = myRow.Item("PagerNumber").ToString
                appUserId = myRow.Item("AppUserId").ToString
                usertype = myRow.Item("UserType").ToString
                billedthroughdate = myRow.Item("Identifier6").ToString
                administrator = myRow.Item("Administrator").ToString
                enhancedServOption = "A"

                If pagerNumber.Trim <> "" Then
                    myPhoneNumber = ""
                    For ch = 0 To Len(pagerNumber) - 1
                        If IsNumeric(pagerNumber.Substring(ch, 1)) Then
                            myPhoneNumber = myPhoneNumber & pagerNumber.Substring(ch, 1)
                        End If
                    Next
                    If myPhoneNumber.Trim <> "" Then
                        pagerNumber = myPhoneNumber
                    End If
                End If
                ''If coCode.Trim = "" Or acctNo.Trim = "" Or billRateCode.Trim = "" Or pagerNumber.Trim = "" Then
                'If billRateCode.Trim = "" Or pagerNumber.Trim = "" Then
                '    myURIResponse = "Invalid Data - Update not performed."
                '    SendErrorEmail(myDS1, myURIResponse, appUserId, coCode, acctNo, pagerNumber, billRateCode, enhancedServOption)

                '    pRecords = pRecords + 1

                'Else
                If pagerNumber.Trim <> "" Then
                    myURIData = "<msg>" & coCode & "/" & acctNo & "/" & billRateCode & "/" & pagerNumber & "/" & appUserId & "/" & enhancedServOption & "/" & usertype & "/" & administrator & "/" & billedthroughdate & "</msg>"
                    myURIRequest = myURIBase & myURIData

                    mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)
                    If mySendSuccess = False Then
                        myURIResponse = myExceptionText
                    End If

                    If myURIResponse.Contains("COMPLETION-CODE:.") Then
                        If myURIResponse.IndexOf("/") <> 0 Then
                            holdURIResponse = "COMPLETION-CODE:."
                        Else
                            holdURIResponse = ""
                        End If
                    Else
                        holdURIResponse = ""
                    End If

                    mySQL = "execute dbo.SaveAbacusTransaction @AppUserId, @CoCode, @AcctNo, @BillRateCode, @PagerNumber, @EnhancedServOption, @AbacusReturnMsg"
                    Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                    With myAdapter
                        Try
                            .SelectCommand.Parameters.Add("@AppUserId", SqlDbType.VarChar).Value = appUserId
                            .SelectCommand.Parameters.Add("@CoCode", SqlDbType.Char).Value = coCode
                            .SelectCommand.Parameters.Add("@AcctNo", SqlDbType.Char).Value = acctNo
                            .SelectCommand.Parameters.Add("@BillRateCode", SqlDbType.Char).Value = billRateCode
                            .SelectCommand.Parameters.Add("@PagerNumber", SqlDbType.VarChar).Value = pagerNumber
                            .SelectCommand.Parameters.Add("@EnhancedServOption", SqlDbType.Char).Value = enhancedServOption
                            .SelectCommand.Parameters.Add("@AbacusReturnMsg", SqlDbType.VarChar).Value = myURIResponse

                            myData = New DataTable("Response")
                            Using myAdapter
                                .Fill(myData)
                            End Using
                        Catch ex As Exception
                            ' MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                            WriteToErrorEmail(ex.Message.ToString)
                        End Try
                    End With

                    If holdURIResponse <> "" Then
                        myURIResponse = holdURIResponse
                    End If

                    If myURIResponse <> "COMPLETION-CODE:." Or myTimeOut = True Then
                        SendErrorEmail(myDS1, myURIResponse, appUserId, coCode, acctNo, pagerNumber, billRateCode, enhancedServOption)
                    End If

                    pRecords = pRecords + 1
                End If


                'End If
            Next
        End If
    End Sub

    Private Sub ReadAccountsToDisconnect(ByRef pDS As DataSet)
        Dim mySql As String = ""
        Dim myFlag As Boolean = False

        mySql = "exec GetAccountsToDisconnect"
        myFlag = FillSqlDataSet(mySql, pDS, gErrorMessage)
        myAccounts = myFlag
    End Sub

    Private Sub SendAccountsToDisconnect(ByRef pDS As DataSet, ByRef pRecords As Integer)
        Dim myDataTable As DataTable = Nothing
        Dim mySendSuccess As Boolean = False
        Dim pagerNumber As String = ""
        Dim newAppUserId As String = ""
        Dim oldAppUserId As String = ""
        Dim myURIData As String = ""
        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMTC838+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""
        Dim holdURIResponse As String = ""
        Dim mySQL As String = ""
        Dim myData As DataTable = Nothing
        Dim myPhoneNumber As String = ""

        If pDS.Tables.Count > 0 Then
            myDataTable = pDS.Tables(0)
            For Each myRow As DataRow In myDataTable.Rows
                pagerNumber = myRow.Item("PagerPhoneNo").ToString
                oldAppUserId = myRow.Item("oldAppUserId").ToString
                newAppUserId = myRow.Item("newAppUserId").ToString

                If pagerNumber.Trim <> "" Then
                    myPhoneNumber = ""
                    For ch = 0 To Len(pagerNumber) - 1
                        If IsNumeric(pagerNumber.Substring(ch, 1)) Then
                            myPhoneNumber = myPhoneNumber & pagerNumber.Substring(ch, 1)
                        End If
                    Next
                    If myPhoneNumber.Trim <> "" Then
                        pagerNumber = myPhoneNumber
                    End If
                End If

                If pagerNumber.Trim <> "" Then
                    myURIData = "<msg>" & pagerNumber & "/" & oldAppUserId & "</msg>"
                    myURIRequest = myURIBase & myURIData

                    mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)
                    If mySendSuccess = False Then
                        myURIResponse = myExceptionText
                    End If

                    If myURIResponse.Contains("COMPLETION-CODE:.") Then
                        If myURIResponse.IndexOf("/") <> 0 Then
                            holdURIResponse = "COMPLETION-CODE:."
                        Else
                            holdURIResponse = ""
                        End If
                    Else
                        holdURIResponse = ""
                    End If

                    If myURIResponse.Contains("COMPLETION-CODE:.") Then
                        myURIResponse = "."
                    End If

                    mySQL = "execute dbo.UpdateAppUserRename @PagerPhoneNo, @OldAppUserID, @NewAppUserID, @ProcessingProblems"
                    Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                    With myAdapter
                        Try
                            .SelectCommand.Parameters.Add("@PagerPhoneNo", SqlDbType.VarChar).Value = pagerNumber
                            .SelectCommand.Parameters.Add("@OldAppUserID", SqlDbType.Char).Value = oldAppUserId
                            .SelectCommand.Parameters.Add("@NewAppUserID", SqlDbType.Char).Value = newAppUserId
                            .SelectCommand.Parameters.Add("@ProcessingProblems", SqlDbType.VarChar).Value = myURIResponse

                            myData = New DataTable("Response")
                            Using myAdapter
                                .Fill(myData)
                            End Using
                        Catch ex As Exception
                            ' MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                            WriteToErrorEmail(ex.Message.ToString)
                        End Try
                    End With

                    If holdURIResponse <> "" Then
                        myURIResponse = holdURIResponse
                    End If

                    If myURIResponse <> "COMPLETION-CODE:." Or myTimeOut = True Then
                        SendIMDiscoErrorEmail(myDS1, myURIResponse, oldAppUserId, newAppUserId, pagerNumber)
                    End If

                    pRecords = pRecords + 1
                End If


                'End If
            Next
        End If
    End Sub

    Private Function SendHTTPRequest(ByVal pURIRequest As String, ByRef pURIResponse As String) As Boolean

        Dim fr As System.Net.HttpWebRequest
        Dim mySuccess As Boolean = False
        Dim p As Integer = 0
        Dim myHarkNumber As Boolean = False
        myExceptionText = ""

        Try
            Dim myDataset As New DataSet("XMLData")
            Dim bp As Integer = 0
            Dim ep As Integer = 0
            Dim myWorkString As String = ""
            Dim myresponse As HttpWebResponse = Nothing

            fr = DirectCast(HttpWebRequest.Create(pURIRequest), System.Net.HttpWebRequest)
            myresponse = fr.GetResponse

            If (myresponse.ContentLength >= -1) Then 'will wait for response
                'If (fr.GetResponse().ContentLength > 0) Then 'will wait for response

                Dim Reader As New System.IO.StreamReader(myresponse.GetResponseStream())

                'Dim Reader As New System.IO.StreamReader(fr.GetResponse().GetResponseStream())
                'Dim myresponse As HttpWebResponse
                'myresponse = fr.GetResponse
                Dim myXML As String = Reader.ReadToEnd 'Holds XML formated data

                'check for bad response  Need to add other tests maybe
                p = myXML.IndexOf("The XML Server is down for maintenance")
                If p > 0 Then ' BAD OR NO RECORD                    
                    myXML = ""
                End If

                myWorkString = ""
                If myXML = "" Then
                    myResponseReceived = False
                    myTimeOut = True
                    pURIResponse = "No Reply"
                Else
                    myResponseReceived = True
                    myTimeOut = False
                    bp = myXML.IndexOf("<ReturnMsg>")
                    If bp > 0 Then
                        bp = bp + 11
                        ep = myXML.IndexOf("</ReturnMsg>")
                        If ep > 0 Then
                            myWorkString = myXML.Substring(bp, ep - bp)
                        End If
                    Else
                        myHarkNumber = True
                    End If
                End If


                If pURIRequest.IndexOf("xmlrequest=METHOD+WIMCLFWD+") > 0 Or pURIRequest.IndexOf("xmlrequest=METHOD+WIUIMSGU+") > 0 Or pURIRequest.IndexOf("xmlrequest=METHOD+WIMXMLHB+") > 0 Then
                    pURIResponse = myWorkString ' METHOD+WIMCLFWD RETURNS . or error message
                Else
                    'If myHarkNumber = False And pURIRequest.IndexOf("xmlrequest=METHOD+WIMCLFWD+") = 0 Then
                    If myHarkNumber = False Then
                        If myWorkString <> "" Then
                            bp = myWorkString.IndexOf(" ")
                            bp = bp + 1
                            pURIResponse = myWorkString.Substring(bp)
                        Else
                            If myResponseReceived = True Then
                                If myXML.Length > 255 Then
                                    pURIResponse = myXML.Substring(0, 254)
                                Else
                                    If myXML = "" Then
                                        pURIResponse = "No Valid Response"
                                    Else
                                        pURIResponse = myXML
                                    End If
                                End If
                            End If
                        End If
                    Else
                        pURIResponse = myXML
                    End If
                End If


                Reader.Close()
                Reader = Nothing

                myresponse.Close()
                myresponse = Nothing
                'txtText.Text = myXML


                If InStr(myXML, "WILCAPCG") = 0 And InStr(myXML, "WILCAPCH") = 0 And _
                InStr(myXML, "WILCHANL") = 0 And InStr(myXML, "WILCUECF") = 0 And InStr(myXML, "WILCUEEQ") = 0 And InStr(myXML, "WILMCDOO") = 0 Then
                    Me.txtText.AppendText(myXML & vbCrLf)
                End If
                If InStr(myXML, "WILCAPCG") > 0 Then
                    Me.txtText.AppendText("WILCAPCG" & vbCrLf)
                End If
                If InStr(myXML, "WILCAPCH") > 0 Then
                    Me.txtText.AppendText("WILCAPCH" & vbCrLf)
                End If
                If InStr(myXML, "WILCHANL") > 0 Then
                    Me.txtText.AppendText("WILCHANL" & vbCrLf)
                End If
                If InStr(myXML, "WILCUECF") > 0 Then
                    Me.txtText.AppendText("WILCUECF" & vbCrLf)
                End If
                If InStr(myXML, "WILCUEEQ") > 0 Then
                    Me.txtText.AppendText("WILCUEEQ" & vbCrLf)
                End If
                If InStr(myXML, "WIMCLFWD") > 0 Then
                    Me.txtText.AppendText("WIMCLFWD" & vbCrLf)
                End If
                If InStr(myXML, "WIMMCDOO") > 0 Then
                    Me.txtText.AppendText("WIMCLFWD" & vbCrLf)
                End If
            End If

            mySuccess = True
            fr.GetResponse().Close()
            fr = Nothing

        Catch e As WebException
            AppendText(e.Message.ToString)
            WriteToErrorEmail(e.Message.ToString)
            fr = Nothing

            myExceptionText = e.Message

        Catch ex As Exception

            WriteToErrorEmail(ex.Message.ToString)
            fr = Nothing

            myExceptionText = ex.Message

        End Try

        Return mySuccess

    End Function

    Private Sub SendErrorEmail(ByRef pDS1 As DataSet, ByVal pURIResponse As String, ByVal pAppUserId As String, ByVal pCoCode As String, ByVal pAcctNo As String, ByVal pPagerNumber As String, ByVal pBillRateCode As String, ByVal pEnhancedServOption As String)
        Dim myDataTable As DataTable = Nothing
        Dim mySendSuccess As Boolean = False
        Dim emailAddress As String = ""
        Dim myEmailStatement As String = pURIResponse
        Dim myFirstEmail As Integer = 1
        Dim mySQL As String = ""
        Dim myData As DataTable = Nothing
        Dim myFromAddress As String = ""
        Dim mySubject As String = ""
        Dim myOOXMLSRVSubject As String = ""
        Dim myOOXMLSRVBody As String = ""

        If myTimeOut = True Then
            If myEmailSentTime > Now Then
                Return
            End If
            myEmailStatement = "No Response from Server"
        End If

        If gEnvironment = "P" Then
            'myFromAddress = "administrator@intellimsg.net"
            mySubject = "ABACUS Interface failure"
            myOOXMLSRVSubject = "OOXMLSRV"
            myOOXMLSRVBody = "The OOXMLSRV process has stopped on atprod2"
        End If

        If gEnvironment = "T" Then
            'myFromAddress = "administrator@dev.intellimsg.net"      ' use Dev for now
            mySubject = "DEV: ABACUS Interface failure"             ' use Dev for now
            myOOXMLSRVSubject = "DEV Email: OOXMLSRV"               ' use Dev for now
            myOOXMLSRVBody = "The OOXMLSRV process has stopped on devm" ' use Dev for now
        End If

        If gEnvironment = "D" Then
            'myFromAddress = "administrator@dev.intellimsg.net"
            mySubject = "DEV: ABACUS Interface failure"
            myOOXMLSRVSubject = "DEV Email: OOXMLSRV"
            myOOXMLSRVBody = "The OOXMLSRV process has stopped on devm"
        End If

        myEmailStatement = myEmailStatement & vbCrLf & vbCrLf
        myEmailStatement = myEmailStatement & "AppUserID: " & pAppUserId & vbCrLf
        myEmailStatement = myEmailStatement & "CoCode: " & pCoCode & vbCrLf
        myEmailStatement = myEmailStatement & "AcctNo: " & pAcctNo & vbCrLf
        myEmailStatement = myEmailStatement & "PagerNumber: " & pPagerNumber & vbCrLf
        myEmailStatement = myEmailStatement & "BillRateCode: " & pBillRateCode & vbCrLf
        myEmailStatement = myEmailStatement & "EnhancedServiceOption: " & pEnhancedServOption & vbCrLf

        myDataTable = New DataTable
        myDataTable.Columns.Add("EmailAddress", Type.GetType("System.String"))

        Dim myNewRow As DataRow = myDataTable.NewRow
        myDataTable.Rows.Add("larryr@abw.com")
        myDataTable.Rows.Add("amy.williams@americanmessaging.net")
        myDataTable.Rows.Add("zerrialb@abw.com")
        myDataTable.Rows.Add("2147071687@txt.att.net")

        For Each myRow As DataRow In myDataTable.Rows
            emailAddress = myRow.Item("EmailAddress").ToString.Trim
            If emailAddress <> "" Then
                If emailAddress <> myFromAddress Then
                    mySQL = "execute dbo.SendEmail @From, @To, @Subject, @Body, @ReturnMsg"
                    Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                    With myAdapter
                        Try
                            .SelectCommand.Parameters.Add("@From", SqlDbType.VarChar).Value = gEmailFromAddress
                            .SelectCommand.Parameters.Add("@To", SqlDbType.Char).Value = emailAddress
                            .SelectCommand.Parameters.Add("@Subject", SqlDbType.Char).Value = mySubject
                            .SelectCommand.Parameters.Add("@Body", SqlDbType.Char).Value = "The ABACUS Interface has failed for the following reason: " & vbCrLf & vbCrLf & myEmailStatement
                            .SelectCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar).Value = ""

                            myData = New DataTable("Response")
                            Using myAdapter
                                .Fill(myData)
                            End Using
                        Catch ex As Exception
                            AppendText(ex.Message.ToString)
                            'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                        End Try
                    End With
                End If
            End If

            If myTimeOut = True Then
                myEmailSentTime = Now.AddMinutes(5)
            End If

        Next

        'Used when XML server is down
        If pURIResponse = "No Reply" Then
            mySQL = "execute dbo.SendEmail @From, @To, @Subject, @Body, @ReturnMsg"
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
            With myAdapter
                Try
                    .SelectCommand.Parameters.Add("@From", SqlDbType.VarChar).Value = myFromAddress
                    .SelectCommand.Parameters.Add("@To", SqlDbType.Char).Value = "2147071687@txt.att.net; 2147256985@txt.att.net"
                    .SelectCommand.Parameters.Add("@Subject", SqlDbType.Char).Value = myOOXMLSRVSubject
                    .SelectCommand.Parameters.Add("@Body", SqlDbType.Char).Value = myOOXMLSRVBody
                    .SelectCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar).Value = ""

                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using
                Catch ex As Exception
                    AppendText(ex.Message.ToString)
                    'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                End Try
            End With
        End If
    End Sub

    Private Sub SendIMDiscoErrorEmail(ByRef pDS1 As DataSet, ByVal pURIResponse As String, ByVal pOldAppUserId As String, ByVal pNewAppUserId As String, ByVal pPagerNumber As String)
        Dim myDataTable As DataTable = Nothing
        Dim mySendSuccess As Boolean = False
        Dim emailAddress As String = ""
        Dim myEmailStatement As String = pURIResponse
        Dim myFirstEmail As Integer = 1
        Dim mySQL As String = ""
        Dim myData As DataTable = Nothing
        Dim myFromAddress As String = ""
        Dim mySubject As String = ""
        Dim myOOXMLSRVSubject As String = ""
        Dim myOOXMLSRVBody As String = ""

        If myTimeOut = True Then
            If myEmailSentTime > Now Then
                Return
            End If
            myEmailStatement = "No Response from Server"
        End If

        If gEnvironment = "P" Then
            'myFromAddress = "administrator@intellimsg.net"
            mySubject = "ABACUS Interface failure"
            myOOXMLSRVSubject = "OOXMLSRV"
            myOOXMLSRVBody = "The OOXMLSRV process has stopped on atprod2"
        End If

        If gEnvironment = "T" Then
            'myFromAddress = "administrator@dev.intellimsg.net"      ' use Dev for now
            mySubject = "DEV: ABACUS Interface failure"             ' use Dev for now
            myOOXMLSRVSubject = "DEV Email: OOXMLSRV"               ' use Dev for now
            myOOXMLSRVBody = "The OOXMLSRV process has stopped on devm" ' use Dev for now
        End If

        If gEnvironment = "D" Then
            'myFromAddress = "administrator@dev.intellimsg.net"
            mySubject = "DEV: ABACUS Interface failure"
            myOOXMLSRVSubject = "DEV Email: OOXMLSRV"
            myOOXMLSRVBody = "The OOXMLSRV process has stopped on devm"
        End If

        myEmailStatement = myEmailStatement & vbCrLf & vbCrLf
        myEmailStatement = myEmailStatement & "OldAppUserID: " & pOldAppUserId & vbCrLf
        myEmailStatement = myEmailStatement & "NewAppUserID: " & pNewAppUserId & vbCrLf
        myEmailStatement = myEmailStatement & "PagerNumber: " & pPagerNumber & vbCrLf
        myEmailStatement = myEmailStatement & "Intellimsg Rename" & vbCrLf

        myDataTable = New DataTable
        myDataTable.Columns.Add("EmailAddress", Type.GetType("System.String"))

        Dim myNewRow As DataRow = myDataTable.NewRow
        myDataTable.Rows.Add("larryr@abw.com")
        myDataTable.Rows.Add("amy.williams@americanmessaging.net")
        myDataTable.Rows.Add("zerrialb@abw.com")
        myDataTable.Rows.Add("2147071687@txt.att.net")

        For Each myRow As DataRow In myDataTable.Rows
            emailAddress = myRow.Item("EmailAddress").ToString.Trim
            If emailAddress <> "" Then
                If emailAddress <> myFromAddress Then
                    mySQL = "execute dbo.SendEmail @From, @To, @Subject, @Body, @ReturnMsg"
                    Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                    With myAdapter
                        Try
                            .SelectCommand.Parameters.Add("@From", SqlDbType.VarChar).Value = gEmailFromAddress
                            .SelectCommand.Parameters.Add("@To", SqlDbType.Char).Value = emailAddress
                            .SelectCommand.Parameters.Add("@Subject", SqlDbType.Char).Value = mySubject
                            .SelectCommand.Parameters.Add("@Body", SqlDbType.Char).Value = "The ABACUS Interface has failed for the following reason: " & vbCrLf & vbCrLf & myEmailStatement
                            .SelectCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar).Value = ""

                            myData = New DataTable("Response")
                            Using myAdapter
                                .Fill(myData)
                            End Using
                        Catch ex As Exception
                            AppendText(ex.Message.ToString)
                            'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                        End Try
                    End With
                End If
            End If

            If myTimeOut = True Then
                myEmailSentTime = Now.AddMinutes(5)
            End If

        Next

        'Used when XML server is down
        If pURIResponse = "No Reply" Then
            mySQL = "execute dbo.SendEmail @From, @To, @Subject, @Body, @ReturnMsg"
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
            With myAdapter
                Try
                    .SelectCommand.Parameters.Add("@From", SqlDbType.VarChar).Value = myFromAddress
                    .SelectCommand.Parameters.Add("@To", SqlDbType.Char).Value = "2147071687@txt.att.net; 2147256985@txt.att.net"
                    .SelectCommand.Parameters.Add("@Subject", SqlDbType.Char).Value = myOOXMLSRVSubject
                    .SelectCommand.Parameters.Add("@Body", SqlDbType.Char).Value = myOOXMLSRVBody
                    .SelectCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar).Value = ""

                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using
                Catch ex As Exception
                    AppendText(ex.Message.ToString)
                    'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                End Try
            End With
        End If
    End Sub

    Private Sub ReadEmailAddresses(ByRef pDS1 As DataSet)
        Dim mySql As String = ""
        Dim myFlag As Boolean = False

        mySql = "exec GetErrorEmailAddresses"
        myFlag = FillSqlDataSet(mySql, pDS1, gErrorMessage)
    End Sub

    Public Sub AppendText(ByVal pText As String)
        If Me.myAppendLineCountMax > 0 Then
            If Me.myAppendLineCount > Me.myAppendLineCountMax Then
                Me.txtText.Text = ""
                Me.myAppendLineCount = 0
                GC.Collect()
            End If
        End If
        Me.myAppendLineCount = Me.myAppendLineCount + 1
        Me.txtText.AppendText(Format(DateTime.Now, "MM/dd/yy HH:mm:ss") & " " & pText & vbCrLf)
        Application.DoEvents()
    End Sub

    Private Sub ReadHarkNumbers(ByRef pHarkDS As DataSet)
        Dim mySendSuccess As Boolean = False
        Dim myURIData As String = ""

        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=READ+WIFHARKT"   'Dev outward facing IP - Use normally
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=READ+WIFHARKT"     'dev inside IP
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=READ+WIFHARKT"      'BamBam\test IP

        'Dim myURIBase As String = gURIBase & "xmlrequest=READ+WIFHARKT"
        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMHARKT"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        myURIRequest = myURIBase

        mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

        If mySendSuccess = True Then
            If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                Dim P1 As Integer = 0
                Dim P2 As Integer = 0
                Dim myError As String = myURIResponse
                While myError.IndexOf("<ERROR>") >= 0
                    P1 = myError.IndexOf("<ERROR>") + 7
                    P2 = myError.LastIndexOf("</ERROR>")
                    myError = myError.Substring(P1, P2 - P1)
                End While
                AppendText("XML Error: " & myError)
                pHarkDS = Nothing
            Else
                Dim dstHARKTXMLData As New DataSet()
                dstHARKTXMLData = ConvertXMLToDataSet(myURIResponse)
                pHarkDS = dstHARKTXMLData
            End If
        Else
            pHarkDS = Nothing
        End If
    End Sub

    Private Sub ReadCapcodeGroup()
        Dim myRead As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 10

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCAPCG+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Dim YYYYMMDDHHMISS As String = ""
        Dim DATA_ACTION As String = ""
        Dim CHANNEL_CODE As String = ""
        Dim PAGER_FORMAT As String = ""
        Dim CAPCODE_IND As String = ""
        Dim CAPCODE_NAME_IND As String = ""
        Dim CO_CODE As String = ""
        Dim ACCT_NO As String = ""
        Dim CAPCODE_OWNER_NAME As String = ""
        Dim CAPCODE_OWNER_PHONE As String = ""
        Dim CAPCODE_OWNER_EMAIL As String = ""
        Dim CUE_COLOR As String = ""
        Dim CUE_ALERT_TONE As String = ""
        Dim FOLDER As String = ""
        Dim UNLOCK_DISPLAY As String = ""
        Dim ENCRYPT As String = ""
        Dim MAIL_DROP As String = ""
        Dim FLAG_UI As String = ""
        Dim FLAG_DEFERRED As String = ""
        Dim FLAG_LOCAL As String = ""
        Dim CAPCODE_PHONE_NO As String = ""
        Dim STMT_DESCRIPTION As String = ""

        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing

        While myContinue = True And myCount < myLimit

            myURIRequest = myURIBase & "<MSG>" & HTA(YYYYMMDDHHMISS) & "/" _
                                                & HTA(DATA_ACTION) & "/" _
                                                & HTA(CHANNEL_CODE) & "/" _
                                                & HTA(PAGER_FORMAT) & "/" _
                                                & HTA(CAPCODE_IND) & "</MSG>"

            myRead = SendHTTPRequest(myURIRequest, myURIResponse)

            If myRead = True Then
                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                    Dim P1 As Integer = 0
                    Dim P2 As Integer = 0
                    Dim myError As String = myURIResponse
                    While myError.IndexOf("<ERROR>") >= 0
                        P1 = myError.IndexOf("<ERROR>") + 7
                        P2 = myError.LastIndexOf("</ERROR>")
                        myError = myError.Substring(P1, P2 - P1)
                    End While
                    AppendText("XML Error: " & myError)
                    myContinue = False
                Else
                    Dim dstXMLData As New DataSet()
                    dstXMLData = ConvertXMLToDataSet(myURIResponse)
                    YYYYMMDDHHMISS = ATH(dstXMLData.Tables(0).Rows(0).Item("YYYYMMDDHHMISS"))
                    DATA_ACTION = ATH(dstXMLData.Tables(0).Rows(0).Item("DATA_ACTION"))
                    CHANNEL_CODE = ATH(dstXMLData.Tables(0).Rows(0).Item("CHANNEL_CODE"))
                    PAGER_FORMAT = ATH(dstXMLData.Tables(0).Rows(0).Item("PAGER_FORMAT"))
                    CAPCODE_IND = ATH(dstXMLData.Tables(0).Rows(0).Item("CAPCODE_IND"))
                    CAPCODE_NAME_IND = ATH(dstXMLData.Tables(0).Rows(0).Item("CAPCODE_NAME_IND"))
                    CO_CODE = ATH(dstXMLData.Tables(0).Rows(0).Item("CO_CODE"))
                    ACCT_NO = ATH(dstXMLData.Tables(0).Rows(0).Item("ACCT_NO"))
                    CAPCODE_OWNER_NAME = ATH(dstXMLData.Tables(0).Rows(0).Item("CAPCODE_OWNER_NAME"))
                    CAPCODE_OWNER_PHONE = ATH(dstXMLData.Tables(0).Rows(0).Item("CAPCODE_OWNER_PHONE"))
                    CAPCODE_OWNER_EMAIL = ATH(dstXMLData.Tables(0).Rows(0).Item("CAPCODE_OWNER_EMAIL"))
                    CUE_COLOR = ATH(dstXMLData.Tables(0).Rows(0).Item("CUE_COLOR"))
                    CUE_ALERT_TONE = ATH(dstXMLData.Tables(0).Rows(0).Item("CUE_ALERT_TONE"))
                    FOLDER = ATH(dstXMLData.Tables(0).Rows(0).Item("FOLDER"))
                    UNLOCK_DISPLAY = ATH(dstXMLData.Tables(0).Rows(0).Item("UNLOCK_DISPLAY"))
                    ENCRYPT = ATH(dstXMLData.Tables(0).Rows(0).Item("ENCRYPT"))
                    MAIL_DROP = ATH(dstXMLData.Tables(0).Rows(0).Item("MAIL_DROP"))

                    FLAG_UI = ATH(dstXMLData.Tables(0).Rows(0).Item("FLAG_UI"))
                    FLAG_DEFERRED = ATH(dstXMLData.Tables(0).Rows(0).Item("FLAG_DEFERRED"))
                    FLAG_LOCAL = ATH(dstXMLData.Tables(0).Rows(0).Item("FLAG_LOCAL"))
                    CAPCODE_PHONE_NO = ATH(dstXMLData.Tables(0).Rows(0).Item("CAPCODE_PHONE_NO"))
                    STMT_DESCRIPTION = ATH(dstXMLData.Tables(0).Rows(0).Item("STMT_DESCRIPTION"))
                    If YYYYMMDDHHMISS.Trim = "" Then
                        myContinue = False
                    Else
                        mySql = "execute dbo.SaveAbacusWIFCAPCG @YYYYMMDDHHMISS, @DATA_ACTION, " _
                            & "@CHANNEL_CODE, @PAGER_FORMAT, @CAPCODE_IND, @CAPCODE_NAME_IND, " _
                            & "@CO_CODE, @ACCT_NO, @CAPCODE_OWNER_NAME, @CAPCODE_OWNER_PHONE, @CAPCODE_OWNER_EMAIL, " _
                            & "@CUE_COLOR, @CUE_ALERT_TONE, @FOLDER, @UNLOCK_DISPLAY, @ENCRYPT, @MAIL_DROP,@FLAG_UI, @FLAG_DEFERRED, @FLAG_LOCAL, @CAPCODE_PHONE_NO, @STMT_DESCRIPTION"
                        Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                        With myAdapter
                            Try
                                .SelectCommand.Parameters.Add("@YYYYMMDDHHMISS", SqlDbType.VarChar).Value = YYYYMMDDHHMISS
                                .SelectCommand.Parameters.Add("@DATA_ACTION", SqlDbType.Char).Value = DATA_ACTION
                                .SelectCommand.Parameters.Add("@CHANNEL_CODE", SqlDbType.Char).Value = CHANNEL_CODE
                                .SelectCommand.Parameters.Add("@PAGER_FORMAT", SqlDbType.Char).Value = PAGER_FORMAT
                                .SelectCommand.Parameters.Add("@CAPCODE_IND", SqlDbType.VarChar).Value = CAPCODE_IND
                                .SelectCommand.Parameters.Add("@CAPCODE_NAME_IND", SqlDbType.Char).Value = CAPCODE_NAME_IND
                                .SelectCommand.Parameters.Add("@CO_CODE", SqlDbType.VarChar).Value = CO_CODE
                                .SelectCommand.Parameters.Add("@ACCT_NO", SqlDbType.VarChar).Value = ACCT_NO
                                .SelectCommand.Parameters.Add("@CAPCODE_OWNER_NAME", SqlDbType.VarChar).Value = CAPCODE_OWNER_NAME
                                .SelectCommand.Parameters.Add("@CAPCODE_OWNER_PHONE", SqlDbType.VarChar).Value = CAPCODE_OWNER_PHONE
                                .SelectCommand.Parameters.Add("@CAPCODE_OWNER_EMAIL", SqlDbType.VarChar).Value = CAPCODE_OWNER_EMAIL
                                .SelectCommand.Parameters.Add("@CUE_COLOR", SqlDbType.VarChar).Value = CUE_COLOR
                                .SelectCommand.Parameters.Add("@CUE_ALERT_TONE", SqlDbType.VarChar).Value = CUE_ALERT_TONE
                                .SelectCommand.Parameters.Add("@FOLDER", SqlDbType.VarChar).Value = FOLDER
                                .SelectCommand.Parameters.Add("@UNLOCK_DISPLAY", SqlDbType.VarChar).Value = UNLOCK_DISPLAY
                                .SelectCommand.Parameters.Add("@ENCRYPT", SqlDbType.VarChar).Value = ENCRYPT
                                .SelectCommand.Parameters.Add("@MAIL_DROP", SqlDbType.VarChar).Value = MAIL_DROP

                                .SelectCommand.Parameters.Add("@FLAG_UI", SqlDbType.VarChar).Value = FLAG_UI
                                .SelectCommand.Parameters.Add("@FLAG_DEFERRED", SqlDbType.VarChar).Value = FLAG_DEFERRED
                                .SelectCommand.Parameters.Add("@FLAG_LOCAL", SqlDbType.VarChar).Value = FLAG_LOCAL
                                .SelectCommand.Parameters.Add("@CAPCODE_PHONE_NO", SqlDbType.VarChar).Value = CAPCODE_PHONE_NO
                                .SelectCommand.Parameters.Add("@STMT_DESCRIPTION", SqlDbType.VarChar).Value = STMT_DESCRIPTION


                                myData = New DataTable("Response")
                                Using myAdapter
                                    .Fill(myData)
                                End Using
                                ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                                If ReturnMsg <> "." Then
                                    'MsgBox("WIFCAPCG " & ReturnMsg)
                                    AppendText("WIFCAPCG " & ReturnMsg.ToString)
                                    WriteToErrorEmail("WIFCAPCG " & ReturnMsg.ToString)
                                    myContinue = False
                                End If

                                myData = Nothing
                                myAdapter = Nothing
                            Catch ex As Exception

                                WriteToErrorEmail(ex.Message.ToString)
                                'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                            End Try
                        End With
                        myCount = myCount + 1
                        If myCount = myLimit Then 'Getting out of loop, need to delete last transaction
                            myURIRequest = myURIBase & "<MSG>" & HTA(YYYYMMDDHHMISS) & "/" _
                                                                & HTA(DATA_ACTION) & "/" _
                                                                & HTA(CHANNEL_CODE) & "/" _
                                                                & HTA(PAGER_FORMAT) & "/" _
                                                                & HTA(CAPCODE_IND) & "</MSG>"
                            myRead = SendHTTPRequest(myURIRequest, myURIResponse)
                        End If
                    End If
                End If
            Else
                myContinue = False
            End If

        End While

    End Sub

    Private Sub ReadCapcodeGroupAccounts()
        Dim myRead As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 10

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCAPCH+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Dim YYYYMMDDHHMISS As String = ""
        Dim DATA_ACTION As String = ""
        Dim CHANNEL_CODE As String = ""
        Dim PAGER_FORMAT As String = ""
        Dim CAPCODE_IND As String = ""
        Dim CO_CODE As String = ""
        Dim ACCT_NO As String = ""

        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing

        While myContinue = True And myCount < myLimit

            myURIRequest = myURIBase & "<MSG>" & HTA(YYYYMMDDHHMISS) & "/" _
                                                & HTA(DATA_ACTION) & "/" _
                                                & HTA(CHANNEL_CODE) & "/" _
                                                & HTA(PAGER_FORMAT) & "/" _
                                                & HTA(CAPCODE_IND) & "/" _
                                                & HTA(CO_CODE) & "/" _
                                                & HTA(ACCT_NO) & "</MSG>"

            myRead = SendHTTPRequest(myURIRequest, myURIResponse)

            If myRead = True Then
                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                    Dim P1 As Integer = 0
                    Dim P2 As Integer = 0
                    Dim myError As String = myURIResponse
                    While myError.IndexOf("<ERROR>") >= 0
                        P1 = myError.IndexOf("<ERROR>") + 7
                        P2 = myError.LastIndexOf("</ERROR>")
                        myError = myError.Substring(P1, P2 - P1)
                    End While
                    AppendText("XML Error: " & myError)
                    myContinue = False
                Else
                    Dim dstXMLData As New DataSet()
                    dstXMLData = ConvertXMLToDataSet(myURIResponse)
                    YYYYMMDDHHMISS = ATH(dstXMLData.Tables(0).Rows(0).Item("YYYYMMDDHHMISS"))
                    DATA_ACTION = ATH(dstXMLData.Tables(0).Rows(0).Item("DATA_ACTION"))
                    CHANNEL_CODE = ATH(dstXMLData.Tables(0).Rows(0).Item("CHANNEL_CODE"))
                    PAGER_FORMAT = ATH(dstXMLData.Tables(0).Rows(0).Item("PAGER_FORMAT"))
                    CAPCODE_IND = ATH(dstXMLData.Tables(0).Rows(0).Item("CAPCODE_IND"))
                    CO_CODE = ATH(dstXMLData.Tables(0).Rows(0).Item("CO_CODE"))
                    ACCT_NO = ATH(dstXMLData.Tables(0).Rows(0).Item("ACCT_NO"))
                    If YYYYMMDDHHMISS.Trim = "" Then
                        myContinue = False
                    Else
                        mySql = "execute dbo.SaveAbacusWIFCAPCH @YYYYMMDDHHMISS, @DATA_ACTION, " _
                            & "@CHANNEL_CODE, @PAGER_FORMAT, @CAPCODE_IND, @CO_CODE, @ACCT_NO"
                        Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                        With myAdapter
                            Try
                                .SelectCommand.Parameters.Add("@YYYYMMDDHHMISS", SqlDbType.VarChar).Value = YYYYMMDDHHMISS
                                .SelectCommand.Parameters.Add("@DATA_ACTION", SqlDbType.Char).Value = DATA_ACTION
                                .SelectCommand.Parameters.Add("@CHANNEL_CODE", SqlDbType.Char).Value = CHANNEL_CODE
                                .SelectCommand.Parameters.Add("@PAGER_FORMAT", SqlDbType.Char).Value = PAGER_FORMAT
                                .SelectCommand.Parameters.Add("@CAPCODE_IND", SqlDbType.VarChar).Value = CAPCODE_IND
                                .SelectCommand.Parameters.Add("@CO_CODE", SqlDbType.VarChar).Value = CO_CODE
                                .SelectCommand.Parameters.Add("@ACCT_NO", SqlDbType.VarChar).Value = ACCT_NO

                                myData = New DataTable("Response")
                                Using myAdapter
                                    .Fill(myData)
                                End Using
                                ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                                If ReturnMsg <> "." Then
                                    'MsgBox("WIFCAPCH " & ReturnMsg)
                                    AppendText("WIFCAPCH " & ReturnMsg)
                                    WriteToErrorEmail("WIFCAPCH " & ReturnMsg)
                                    myContinue = False
                                End If

                                myData = Nothing
                                myAdapter = Nothing
                            Catch ex As Exception

                                WriteToErrorEmail(ex.Message.ToString)
                                'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                            End Try
                        End With
                        myCount = myCount + 1
                        If myCount = myLimit Then 'Getting out of loop, need to delete last transaction
                            myURIRequest = myURIBase & "<MSG>" & HTA(YYYYMMDDHHMISS) & "/" _
                                                                & HTA(DATA_ACTION) & "/" _
                                                                & HTA(CHANNEL_CODE) & "/" _
                                                                & HTA(PAGER_FORMAT) & "/" _
                                                                & HTA(CAPCODE_IND) & "/" _
                                                                & HTA(CO_CODE) & "/" _
                                                                & HTA(ACCT_NO) & "</MSG>"
                            myRead = SendHTTPRequest(myURIRequest, myURIResponse)
                        End If
                    End If
                End If
            Else
                myContinue = False
            End If

        End While

    End Sub

    Private Sub ReadChannelCode()
        Dim myRead As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 10

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCHANL+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Dim YYYYMMDDHHMISS As String = ""
        Dim DATA_ACTION As String = ""
        Dim CHANNEL_CODE As String = ""
        Dim FREQUENCY As String = ""

        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing

        While myContinue = True And myCount < myLimit

            myURIRequest = myURIBase & "<MSG>" & HTA(YYYYMMDDHHMISS) & "/" _
                                                & HTA(DATA_ACTION) & "/" _
                                                & HTA(CHANNEL_CODE) & "</MSG>"

            myRead = SendHTTPRequest(myURIRequest, myURIResponse)

            If myRead = True Then
                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                    Dim P1 As Integer = 0
                    Dim P2 As Integer = 0
                    Dim myError As String = myURIResponse
                    While myError.IndexOf("<ERROR>") >= 0
                        P1 = myError.IndexOf("<ERROR>") + 7
                        P2 = myError.LastIndexOf("</ERROR>")
                        myError = myError.Substring(P1, P2 - P1)
                    End While
                    AppendText("XML Error: " & myError)
                    myContinue = False
                Else
                    Dim dstXMLData As New DataSet()
                    dstXMLData = ConvertXMLToDataSet(myURIResponse)
                    YYYYMMDDHHMISS = ATH(dstXMLData.Tables(0).Rows(0).Item("YYYYMMDDHHMISS"))
                    DATA_ACTION = ATH(dstXMLData.Tables(0).Rows(0).Item("DATA_ACTION"))
                    CHANNEL_CODE = ATH(dstXMLData.Tables(0).Rows(0).Item("CHANNEL_CODE"))
                    FREQUENCY = ATH(dstXMLData.Tables(0).Rows(0).Item("FREQUENCY"))
                    If YYYYMMDDHHMISS.Trim = "" Then
                        myContinue = False
                    Else
                        mySql = "execute dbo.SaveAbacusWIFCHANL @YYYYMMDDHHMISS, @DATA_ACTION, " _
                            & "@CHANNEL_CODE, @FREQUENCY"
                        Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                        With myAdapter
                            Try
                                .SelectCommand.Parameters.Add("@YYYYMMDDHHMISS", SqlDbType.VarChar).Value = YYYYMMDDHHMISS
                                .SelectCommand.Parameters.Add("@DATA_ACTION", SqlDbType.Char).Value = DATA_ACTION
                                .SelectCommand.Parameters.Add("@CHANNEL_CODE", SqlDbType.Char).Value = CHANNEL_CODE
                                .SelectCommand.Parameters.Add("@FREQUENCY", SqlDbType.Char).Value = FREQUENCY

                                myData = New DataTable("Response")
                                Using myAdapter
                                    .Fill(myData)
                                End Using
                                ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                                If ReturnMsg <> "." Then
                                    'MsgBox("WIFCHANL " & ReturnMsg)
                                    AppendText("WIFCHANL " & ReturnMsg)
                                    WriteToErrorEmail("WIFCHANL " & ReturnMsg)
                                    myContinue = False
                                End If

                                myData = Nothing
                                myAdapter = Nothing
                            Catch ex As Exception

                                WriteToErrorEmail(ex.Message.ToString)
                                'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                            End Try
                        End With
                        myCount = myCount + 1
                        If myCount = myLimit Then 'Getting out of loop, need to delete last transaction
                            myURIRequest = myURIBase & "<MSG>" & HTA(YYYYMMDDHHMISS) & "/" _
                                                                & HTA(DATA_ACTION) & "/" _
                                                                & HTA(CHANNEL_CODE) & "</MSG>"
                            myRead = SendHTTPRequest(myURIRequest, myURIResponse)
                        End If
                    End If
                End If
            Else
                myContinue = False
            End If

        End While

    End Sub

    Private Sub ReadCueConfiguration()
        Dim myReadCue As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 10

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCUECF+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Dim YYYYMMDDHHMISS As String = ""
        Dim SERIAL_NO As String = ""
        Dim CUE_CAPCODE As String = ""
        Dim CUE_CAPCODE_NAME As String = ""
        Dim CUE_CAPCODE_TYPE As String = ""
        Dim CUE_CAPCODE_ENABLE As String = ""
        Dim CUE_CAPCODE_COLOR As String = ""
        Dim CUE_ALERT_TONE As String = ""
        Dim CUE_FOLDER_ENABLE As String = ""
        Dim CUE_UNLOCK_DISPLAY As String = ""
        Dim CUE_ENCRYPT_ENABLE As String = ""
        Dim CUE_MAIL_DROP As String = ""
        Dim CO_CODE As String = ""
        Dim ACCT_NO As String = ""
        Dim PHONE_NO As String = ""
        Dim PRIMARY_CAPCODE As String = ""
        Dim TRANS_ENTERED_BY As String = ""

        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing

        While myContinue = True And myCount < 10

            myURIRequest = myURIBase & "<MSG>" & YYYYMMDDHHMISS & "/" & SERIAL_NO & "</MSG>"

            myReadCue = SendHTTPRequest(myURIRequest, myURIResponse)

            If myReadCue = True Then
                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                    Dim P1 As Integer = 0
                    Dim P2 As Integer = 0
                    Dim myError As String = myURIResponse
                    While myError.IndexOf("<ERROR>") >= 0
                        P1 = myError.IndexOf("<ERROR>") + 7
                        P2 = myError.LastIndexOf("</ERROR>")
                        myError = myError.Substring(P1, P2 - P1)
                    End While
                    AppendText("XML Error: " & myError)
                    myContinue = False
                Else
                    Dim dstXMLData As New DataSet()
                    dstXMLData = ConvertXMLToDataSet(myURIResponse)
                    YYYYMMDDHHMISS = dstXMLData.Tables(0).Rows(0).Item("YYYYMMDDHHMISS")
                    SERIAL_NO = dstXMLData.Tables(0).Rows(0).Item("SERIAL_NO")
                    CUE_CAPCODE = dstXMLData.Tables(0).Rows(0).Item("CUE_CAPCODE")
                    CUE_CAPCODE_NAME = dstXMLData.Tables(0).Rows(0).Item("CUE_CAPCODE_NAME")
                    CUE_CAPCODE_TYPE = dstXMLData.Tables(0).Rows(0).Item("CUE_CAPCODE_TYPE")
                    CUE_CAPCODE_ENABLE = dstXMLData.Tables(0).Rows(0).Item("CUE_CAPCODE_ENABLE")
                    CUE_CAPCODE_COLOR = dstXMLData.Tables(0).Rows(0).Item("CUE_CAPCODE_COLOR")
                    CUE_ALERT_TONE = dstXMLData.Tables(0).Rows(0).Item("CUE_ALERT_TONE")
                    CUE_FOLDER_ENABLE = dstXMLData.Tables(0).Rows(0).Item("CUE_FOLDER_ENABLE")
                    CUE_UNLOCK_DISPLAY = dstXMLData.Tables(0).Rows(0).Item("CUE_UNLOCK_DISPLAY")
                    CUE_ENCRYPT_ENABLE = dstXMLData.Tables(0).Rows(0).Item("CUE_ENCRYPT_ENABLE")
                    CUE_MAIL_DROP = dstXMLData.Tables(0).Rows(0).Item("CUE_MAIL_DROP")
                    CO_CODE = dstXMLData.Tables(0).Rows(0).Item("CO_CODE")
                    ACCT_NO = dstXMLData.Tables(0).Rows(0).Item("ACCT_NO")
                    PHONE_NO = dstXMLData.Tables(0).Rows(0).Item("PHONE_NO")
                    PRIMARY_CAPCODE = dstXMLData.Tables(0).Rows(0).Item("PRIMARY_CAPCODE")
                    TRANS_ENTERED_BY = dstXMLData.Tables(0).Rows(0).Item("TRANS_ENTERED_BY")
                    If YYYYMMDDHHMISS.Trim = "" Then
                        myContinue = False
                    Else
                        mySql = "execute dbo.SaveAbacusWIFCUECF @YYYYMMDDHHMISS, @SERIAL_NO, " _
                            & "@CUE_CAPCODE, @CUE_CAPCODE_NAME, @CUE_CAPCODE_TYPE, @CUE_CAPCODE_ENABLE, " _
                            & "@CUE_CAPCODE_COLOR, @CUE_ALERT_TONE, @CUE_FOLDER_ENABLE, @CUE_UNLOCK_DISPLAY, " _
                            & "@CUE_ENCRYPT_ENABLE, @CUE_MAIL_DROP, @CO_CODE, @ACCT_NO, @PHONE_NO, " _
                            & "@PRIMARY_CAPCODE, @TRANS_ENTERED_BY"
                        Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                        With myAdapter
                            Try
                                .SelectCommand.Parameters.Add("@YYYYMMDDHHMISS", SqlDbType.VarChar).Value = YYYYMMDDHHMISS
                                .SelectCommand.Parameters.Add("@SERIAL_NO", SqlDbType.Char).Value = SERIAL_NO
                                .SelectCommand.Parameters.Add("@CUE_CAPCODE", SqlDbType.Char).Value = CUE_CAPCODE
                                .SelectCommand.Parameters.Add("@CUE_CAPCODE_NAME", SqlDbType.Char).Value = CUE_CAPCODE_NAME
                                .SelectCommand.Parameters.Add("@CUE_CAPCODE_TYPE", SqlDbType.VarChar).Value = CUE_CAPCODE_TYPE
                                .SelectCommand.Parameters.Add("@CUE_CAPCODE_ENABLE", SqlDbType.Char).Value = CUE_CAPCODE_ENABLE
                                .SelectCommand.Parameters.Add("@CUE_CAPCODE_COLOR", SqlDbType.VarChar).Value = CUE_CAPCODE_COLOR
                                .SelectCommand.Parameters.Add("@CUE_ALERT_TONE", SqlDbType.VarChar).Value = CUE_ALERT_TONE
                                .SelectCommand.Parameters.Add("@CUE_FOLDER_ENABLE", SqlDbType.VarChar).Value = CUE_FOLDER_ENABLE
                                .SelectCommand.Parameters.Add("@CUE_UNLOCK_DISPLAY", SqlDbType.VarChar).Value = CUE_UNLOCK_DISPLAY
                                .SelectCommand.Parameters.Add("@CUE_ENCRYPT_ENABLE", SqlDbType.VarChar).Value = CUE_ENCRYPT_ENABLE
                                .SelectCommand.Parameters.Add("@CUE_MAIL_DROP", SqlDbType.VarChar).Value = CUE_MAIL_DROP
                                .SelectCommand.Parameters.Add("@CO_CODE", SqlDbType.VarChar).Value = CO_CODE
                                .SelectCommand.Parameters.Add("@ACCT_NO", SqlDbType.VarChar).Value = ACCT_NO
                                .SelectCommand.Parameters.Add("@PHONE_NO", SqlDbType.VarChar).Value = PHONE_NO
                                .SelectCommand.Parameters.Add("@PRIMARY_CAPCODE", SqlDbType.VarChar).Value = PRIMARY_CAPCODE
                                .SelectCommand.Parameters.Add("@TRANS_ENTERED_BY", SqlDbType.VarChar).Value = TRANS_ENTERED_BY

                                myData = New DataTable("Response")
                                Using myAdapter
                                    .Fill(myData)
                                End Using
                                ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                                If ReturnMsg <> "." Then
                                    'MsgBox("WIFCUECF " & ReturnMsg)
                                    AppendText("WIFCUECF " & ReturnMsg)
                                    WriteToErrorEmail("WIFCUECF " & ReturnMsg)
                                    myContinue = False
                                End If

                                myData = Nothing
                                myAdapter = Nothing
                            Catch ex As Exception

                                WriteToErrorEmail(ex.Message.ToString)
                                'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                            End Try
                        End With
                        If myCount = myLimit Then 'Getting out of loop, need to delete last transaction
                            myURIRequest = myURIBase & "<MSG>" & YYYYMMDDHHMISS & "/" & SERIAL_NO & "</MSG>"
                            myReadCue = SendHTTPRequest(myURIRequest, myURIResponse)
                        End If
                        myCount = myCount + 1
                    End If
                End If
            Else
                myContinue = False
            End If

        End While

    End Sub

    Private Sub ReadCueEquipment()
        Dim myReadCue As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 10

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCUEEQ+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Dim YYYYMMDDHHMISS As String = ""
        Dim SERIAL_NO As String = ""
        Dim CAPCODE_1 As String = ""
        Dim EQUIPMENT_TYPE As String = ""
        Dim PAGER_FORMAT As String = ""
        Dim CHANNEL_CODE As String = ""
        Dim DUAL_CHANNEL_1 As String = ""
        Dim DUAL_CHANNEL_2 As String = ""
        Dim PRIMARY_FREQUENCY As String = ""
        Dim SECONDARY_FREQUENCY As String = ""
        Dim CO_CODE As String = ""
        Dim ACCT_NO As String = ""
        Dim PHONE_NO_1 As String = ""
        Dim PHONE_NO_2 As String = ""
        Dim PHONE_NO_3 As String = ""
        Dim PHONE_NO_4 As String = ""
        Dim TRANS_ENTERED_BY As String = ""

        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing

        While myContinue = True And myCount < myLimit

            myURIRequest = myURIBase & "<MSG>" & YYYYMMDDHHMISS & "/" & SERIAL_NO & "</MSG>"

            myReadCue = SendHTTPRequest(myURIRequest, myURIResponse)

            If myReadCue = True Then
                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                    Dim P1 As Integer = 0
                    Dim P2 As Integer = 0
                    Dim myError As String = myURIResponse
                    While myError.IndexOf("<ERROR>") >= 0
                        P1 = myError.IndexOf("<ERROR>") + 7
                        P2 = myError.LastIndexOf("</ERROR>")
                        myError = myError.Substring(P1, P2 - P1)
                    End While
                    AppendText("XML Error: " & myError)
                    myContinue = False
                Else
                    Dim dstXMLData As New DataSet()
                    dstXMLData = ConvertXMLToDataSet(myURIResponse)
                    YYYYMMDDHHMISS = dstXMLData.Tables(0).Rows(0).Item("YYYYMMDDHHMISS")
                    SERIAL_NO = dstXMLData.Tables(0).Rows(0).Item("SERIAL_NO")
                    CAPCODE_1 = dstXMLData.Tables(0).Rows(0).Item("CAPCODE_1")
                    EQUIPMENT_TYPE = dstXMLData.Tables(0).Rows(0).Item("EQUIPMENT_TYPE")
                    PAGER_FORMAT = dstXMLData.Tables(0).Rows(0).Item("PAGER_FORMAT")
                    CHANNEL_CODE = dstXMLData.Tables(0).Rows(0).Item("CHANNEL_CODE")
                    DUAL_CHANNEL_1 = dstXMLData.Tables(0).Rows(0).Item("DUAL_CHANNEL_1")
                    DUAL_CHANNEL_2 = dstXMLData.Tables(0).Rows(0).Item("DUAL_CHANNEL_2")
                    PRIMARY_FREQUENCY = dstXMLData.Tables(0).Rows(0).Item("PRIMARY_FREQUENCY")
                    SECONDARY_FREQUENCY = dstXMLData.Tables(0).Rows(0).Item("SECONDARY_FREQUENCY")
                    CO_CODE = dstXMLData.Tables(0).Rows(0).Item("CO_CODE")
                    ACCT_NO = dstXMLData.Tables(0).Rows(0).Item("ACCT_NO")
                    PHONE_NO_1 = dstXMLData.Tables(0).Rows(0).Item("PHONE_NO_1")
                    PHONE_NO_2 = dstXMLData.Tables(0).Rows(0).Item("PHONE_NO_2")
                    PHONE_NO_3 = dstXMLData.Tables(0).Rows(0).Item("PHONE_NO_3")
                    PHONE_NO_4 = dstXMLData.Tables(0).Rows(0).Item("PHONE_NO_4")
                    TRANS_ENTERED_BY = dstXMLData.Tables(0).Rows(0).Item("TRANS_ENTERED_BY")
                    If YYYYMMDDHHMISS.Trim = "" Then
                        myContinue = False
                    Else
                        mySql = "execute dbo.SaveAbacusWIFCUEEQ @YYYYMMDDHHMISS, @SERIAL_NO, " _
                            & "@CAPCODE_1, @EQUIPMENT_TYPE, @PAGER_FORMAT, @CHANNEL_CODE, @DUAL_CHANNEL_1, " _
                            & "@DUAL_CHANNEL_2, @PRIMARY_FREQUENCY, @SECONDARY_FREQUENCY, @CO_CODE, @ACCT_NO, " _
                            & "@PHONE_NO_1, @PHONE_NO_2, @PHONE_NO_3, @PHONE_NO_4, @TRANS_ENTERED_BY"
                        Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                        With myAdapter
                            Try
                                .SelectCommand.Parameters.Add("@YYYYMMDDHHMISS", SqlDbType.VarChar).Value = YYYYMMDDHHMISS
                                .SelectCommand.Parameters.Add("@SERIAL_NO", SqlDbType.Char).Value = SERIAL_NO
                                .SelectCommand.Parameters.Add("@CAPCODE_1", SqlDbType.Char).Value = CAPCODE_1
                                .SelectCommand.Parameters.Add("@EQUIPMENT_TYPE", SqlDbType.Char).Value = EQUIPMENT_TYPE
                                .SelectCommand.Parameters.Add("@PAGER_FORMAT", SqlDbType.Char).Value = PAGER_FORMAT
                                .SelectCommand.Parameters.Add("@CHANNEL_CODE", SqlDbType.VarChar).Value = CHANNEL_CODE
                                .SelectCommand.Parameters.Add("@DUAL_CHANNEL_1", SqlDbType.Char).Value = DUAL_CHANNEL_1
                                .SelectCommand.Parameters.Add("@DUAL_CHANNEL_2", SqlDbType.VarChar).Value = DUAL_CHANNEL_2
                                .SelectCommand.Parameters.Add("@PRIMARY_FREQUENCY", SqlDbType.VarChar).Value = PRIMARY_FREQUENCY
                                .SelectCommand.Parameters.Add("@SECONDARY_FREQUENCY", SqlDbType.VarChar).Value = SECONDARY_FREQUENCY
                                .SelectCommand.Parameters.Add("@CO_CODE", SqlDbType.VarChar).Value = CO_CODE
                                .SelectCommand.Parameters.Add("@ACCT_NO", SqlDbType.VarChar).Value = ACCT_NO
                                .SelectCommand.Parameters.Add("@PHONE_NO_1", SqlDbType.VarChar).Value = PHONE_NO_1
                                .SelectCommand.Parameters.Add("@PHONE_NO_2", SqlDbType.VarChar).Value = PHONE_NO_2
                                .SelectCommand.Parameters.Add("@PHONE_NO_3", SqlDbType.VarChar).Value = PHONE_NO_3
                                .SelectCommand.Parameters.Add("@PHONE_NO_4", SqlDbType.VarChar).Value = PHONE_NO_4
                                .SelectCommand.Parameters.Add("@TRANS_ENTERED_BY", SqlDbType.VarChar).Value = TRANS_ENTERED_BY

                                myData = New DataTable("Response")
                                Using myAdapter
                                    .Fill(myData)
                                End Using
                                ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                                If ReturnMsg <> "." Then
                                    'MsgBox("WIFCUEEQ " & ReturnMsg)
                                    AppendText("WIFCUEEQ " & ReturnMsg)
                                    WriteToErrorEmail("WIFCUEEQ " & ReturnMsg)
                                    myContinue = False
                                End If

                                myData = Nothing
                                myAdapter = Nothing
                            Catch ex As Exception

                                WriteToErrorEmail(ex.Message.ToString)
                                'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                            End Try
                        End With
                        myCount = myCount + 1
                        If myCount = myLimit Then 'Getting out of loop, need to delete last transaction
                            myURIRequest = myURIBase & "<MSG>" & YYYYMMDDHHMISS & "/" & SERIAL_NO & "</MSG>"
                            myReadCue = SendHTTPRequest(myURIRequest, myURIResponse)
                        End If
                    End If
                End If
            Else
                myContinue = False
            End If

        End While

    End Sub

    Private Sub ReadPagerConfigurationTransfer()
        Dim myReadCue As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 10

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCUEPC+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Dim SerialNumber As String = ""
        Dim TransactionDate As Date = "1/1/1900"
        Dim SequenceNumber As Integer = 0
        Dim CUE_SERIAL_NUMBER As String = ""
        Dim CAPCODES As String = ""
        Dim CAPCODE_NAMES As String = ""
        Dim CAPCODE_TYPES As String = ""
        Dim CAPCODE_ENABLES As String = ""
        Dim CAPCODE_COLORS As String = ""
        Dim ALERT_TONES As String = ""
        Dim FOLDER_ENABLES As String = ""
        Dim UNLOCK_DISPLAYS As String = ""
        Dim ENCRYPTION_ENABLES As String = ""
        Dim MAIL_DROP_STORAGES As String = ""
        Dim REWORK_CODE As String = ""
        Dim BATTERY_CONSTANTS As String = ""
        Dim CHARGE_CONTROL As String = ""
        Dim BATTERY_DISPLAY_VOLT As String = ""
        Dim CARRIER_PASSWORD As String = ""
        Dim PRIMARY_FREQUENCY As String = ""
        Dim SECONDARY_FREQUENCY As String = ""
        Dim DUAL_FREQUENCY_ENAB As String = ""
        Dim DUPLICATE_MSG_TIME As String = ""
        Dim ABOUT_MESSAGE As String = ""
        Dim DEVICE_COLLAPSE_VAL As String = ""
        Dim ADMIN_PASSWORD As String = ""
        Dim MAX_DISPLAYED_MSGS As String = ""
        Dim ENCRYPTION_KEY As String = ""
        Dim USER_PASSWORD As String = ""
        Dim INTELLIMESSAGE_DATE As String = ""
        Dim INTELLIMESSAGE_TIME As String = ""
        Dim REMINDER_TONE As String = ""
        Dim REMINDER_DURATION As String = ""
        Dim POWER_OFF_DURATION As String = ""
        Dim VERSION As String = ""
        Dim ALERT_TONE_DURATION As String = ""
        Dim UNLOCK_CODE As String = ""
        Dim PRIMARY_CAPCODE As String = ""
        Dim CHANNEL_CODE As String = ""
        Dim PAGER_FORMAT As String = ""
        Dim EQUIPMENT_TYPE As String = ""
        Dim CO_CODE As String = ""
        Dim ACCT_NO As String = ""
        Dim TRANS_ENTERED_BY As String = ""

        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing

        While myContinue = True And myCount < myLimit

            mySql = "execute dbo.PagerConfigurationTransferUpdate @SerialNumber, @TransactionDate, @SequenceNumber"
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myAdapter
                Try
                    .SelectCommand.Parameters.Add("@SerialNumber", SqlDbType.VarChar).Value = SerialNumber
                    .SelectCommand.Parameters.Add("@TransactionDate", SqlDbType.Char).Value = TransactionDate
                    .SelectCommand.Parameters.Add("@SequenceNumber", SqlDbType.Char).Value = SequenceNumber

                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using
                    ReturnMsg = myData.Rows(0).Item("ReturnMsg")

                    'Keep the key to send back to delete if update to Abacus is completed properly
                    SerialNumber = myData.Rows(0).Item("SerialNumber")
                    TransactionDate = myData.Rows(0).Item("TransactionDate")
                    SequenceNumber = myData.Rows(0).Item("SequenceNumber")

                    'Get the next record to send to Abacus (if any)
                    CUE_SERIAL_NUMBER = myData.Rows(0).Item("SerialNumber")
                    CAPCODES = myData.Rows(0).Item("Capcode")
                    CAPCODE_NAMES = myData.Rows(0).Item("CapcodeName")
                    CAPCODE_TYPES = myData.Rows(0).Item("CapcodeType")
                    CAPCODE_ENABLES = myData.Rows(0).Item("CapcodeEnable")
                    CAPCODE_COLORS = myData.Rows(0).Item("CapcodeColor")
                    ALERT_TONES = myData.Rows(0).Item("AlertTone")
                    FOLDER_ENABLES = myData.Rows(0).Item("FolderEnable")
                    UNLOCK_DISPLAYS = myData.Rows(0).Item("UnlockDisplay")
                    ENCRYPTION_ENABLES = myData.Rows(0).Item("EncryptionEnable")
                    MAIL_DROP_STORAGES = myData.Rows(0).Item("MailDropStorage")
                    REWORK_CODE = myData.Rows(0).Item("ReworkCode").ToString
                    BATTERY_CONSTANTS = myData.Rows(0).Item("BatteryConstants")
                    CHARGE_CONTROL = myData.Rows(0).Item("ChargeControl").ToString
                    BATTERY_DISPLAY_VOLT = myData.Rows(0).Item("BatteryDisplayVolts").ToString
                    CARRIER_PASSWORD = myData.Rows(0).Item("CarrierPassword")
                    PRIMARY_FREQUENCY = myData.Rows(0).Item("PrimaryFrequency")
                    SECONDARY_FREQUENCY = myData.Rows(0).Item("SecondaryFrequency")
                    DUAL_FREQUENCY_ENAB = myData.Rows(0).Item("DualFrequencyEnable")
                    DUPLICATE_MSG_TIME = myData.Rows(0).Item("DuplicateMessageTime")
                    ABOUT_MESSAGE = myData.Rows(0).Item("AboutMessage")
                    DEVICE_COLLAPSE_VAL = myData.Rows(0).Item("DeviceCollapseValue")
                    ADMIN_PASSWORD = myData.Rows(0).Item("AdminPassword")
                    MAX_DISPLAYED_MSGS = myData.Rows(0).Item("MaximumDisplayedMessages")
                    ENCRYPTION_KEY = myData.Rows(0).Item("EncryptionKey")
                    USER_PASSWORD = myData.Rows(0).Item("UserPassword")
                    INTELLIMESSAGE_DATE = myData.Rows(0).Item("IntellimessageDate")
                    INTELLIMESSAGE_TIME = myData.Rows(0).Item("IntellimessageTime")
                    REMINDER_TONE = myData.Rows(0).Item("ReminderTone")
                    REMINDER_DURATION = myData.Rows(0).Item("ReminderDuration")
                    POWER_OFF_DURATION = myData.Rows(0).Item("PowerOffDuration")
                    VERSION = myData.Rows(0).Item("Version")
                    ALERT_TONE_DURATION = myData.Rows(0).Item("AlertToneDuration")
                    UNLOCK_CODE = myData.Rows(0).Item("UnlockCode")
                    PRIMARY_CAPCODE = myData.Rows(0).Item("PrimaryCapcode")
                    CHANNEL_CODE = myData.Rows(0).Item("ChannelCode")
                    PAGER_FORMAT = myData.Rows(0).Item("PagerFormat")
                    EQUIPMENT_TYPE = myData.Rows(0).Item("EquipmentType")
                    CO_CODE = myData.Rows(0).Item("CoCode")
                    ACCT_NO = myData.Rows(0).Item("AcctNo")
                    TRANS_ENTERED_BY = myData.Rows(0).Item("TransEnteredBy")

                    If ReturnMsg <> "." Then
                        'MsgBox("PagerConfigurationTransferUpdate " & ReturnMsg)
                        AppendText("PagerConfigurationTransferUpdate " & ReturnMsg)
                        WriteToErrorEmail("PagerConfigurationTransferUpdate " & ReturnMsg)
                        myContinue = False
                    End If

                    myData = Nothing
                    myAdapter = Nothing
                Catch ex As Exception
                    ' MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                    WriteToErrorEmail(ex.Message.ToString)
                End Try
            End With

            If CUE_SERIAL_NUMBER.Trim <> "" Then
                myURIRequest = myURIBase & "<MSG>" & HTA(CUE_SERIAL_NUMBER) & "/" _
                                                    & HTA(CAPCODES) & "/" _
                                                    & HTA(CAPCODE_NAMES) & "/" _
                                                    & HTA(CAPCODE_TYPES) & "/" _
                                                    & HTA(CAPCODE_ENABLES) & "/" _
                                                    & HTA(CAPCODE_COLORS) & "/" _
                                                    & HTA(ALERT_TONES) & "/" _
                                                    & HTA(FOLDER_ENABLES) & "/" _
                                                    & HTA(UNLOCK_DISPLAYS) & "/" _
                                                    & HTA(ENCRYPTION_ENABLES) & "/" _
                                                    & HTA(MAIL_DROP_STORAGES) & "/" _
                                                    & HTA(REWORK_CODE) & "/" _
                                                    & HTA(BATTERY_CONSTANTS) & "/" _
                                                    & HTA(CHARGE_CONTROL) & "/" _
                                                    & HTA(BATTERY_DISPLAY_VOLT) & "/" _
                                                    & HTA(CARRIER_PASSWORD) & "/" _
                                                    & HTA(PRIMARY_FREQUENCY) & "/" _
                                                    & HTA(SECONDARY_FREQUENCY) & "/" _
                                                    & HTA(DUAL_FREQUENCY_ENAB) & "/" _
                                                    & HTA(DUPLICATE_MSG_TIME) & "/" _
                                                    & HTA(ABOUT_MESSAGE) & "/" _
                                                    & HTA(DEVICE_COLLAPSE_VAL) & "/" _
                                                    & HTA(ADMIN_PASSWORD) & "/" _
                                                    & HTA(MAX_DISPLAYED_MSGS) & "/" _
                                                    & HTA(ENCRYPTION_KEY) & "/" _
                                                    & HTA(USER_PASSWORD) & "/" _
                                                    & HTA(INTELLIMESSAGE_DATE) & "/" _
                                                    & HTA(INTELLIMESSAGE_TIME) & "/" _
                                                    & HTA(REMINDER_TONE) & "/" _
                                                    & HTA(REMINDER_DURATION) & "/" _
                                                    & HTA(POWER_OFF_DURATION) & "/" _
                                                    & HTA(VERSION) & "/" _
                                                    & HTA(ALERT_TONE_DURATION) & "/" _
                                                    & HTA(UNLOCK_CODE) & "/" _
                                                    & HTA(PRIMARY_CAPCODE) & "/" _
                                                    & HTA(CHANNEL_CODE) & "/" _
                                                    & HTA(PAGER_FORMAT) & "/" _
                                                    & HTA(EQUIPMENT_TYPE) & "/" _
                                                    & HTA(CO_CODE) & "/" _
                                                    & HTA(ACCT_NO) & "/" _
                                                    & HTA(TRANS_ENTERED_BY) & "</MSG>"

                myReadCue = SendHTTPRequest(myURIRequest, myURIResponse)

                If myReadCue = True Then
                    If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                        Dim P1 As Integer = 0
                        Dim P2 As Integer = 0
                        Dim myError As String = myURIResponse
                        While myError.IndexOf("<ERROR>") >= 0
                            P1 = myError.IndexOf("<ERROR>") + 7
                            P2 = myError.LastIndexOf("</ERROR>")
                            myError = myError.Substring(P1, P2 - P1)
                        End While
                        AppendText("XML Error: " & myError)
                        myContinue = False
                    End If
                    myCount = myCount + 1
                    If myCount = myLimit Then 'Getting out of loop, need to delete last transaction
                        mySql = "execute dbo.PagerConfigurationTransferUpdate @SerialNumber, @TransactionDate, @SequenceNumber"
                        Dim myDoneAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                        With myDoneAdapter
                            Try
                                .SelectCommand.Parameters.Add("@SerialNumber", SqlDbType.VarChar).Value = SerialNumber
                                .SelectCommand.Parameters.Add("@TransactionDate", SqlDbType.Char).Value = TransactionDate
                                .SelectCommand.Parameters.Add("@SequenceNumber", SqlDbType.Char).Value = SequenceNumber

                                myData = New DataTable("Response")
                                Using myDoneAdapter
                                    .Fill(myData)
                                End Using
                                ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                                myData = Nothing
                                myDoneAdapter = Nothing
                            Catch ex As Exception
                                'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                                WriteToErrorEmail(ex.Message.ToString)
                            End Try
                        End With
                    End If
                End If
            Else
                myContinue = False
            End If

        End While

    End Sub


    Private Sub ReadPagerGCTransfer()
        Dim myReadCue As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 20

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMXTRCC+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""
        Dim mySERIAL_NO As String = ""
        Dim mySLOT As String = ""
        Dim myYYYYMMDDHHMISS As String = ""
        Dim mySEQ_NO As String = ""
        Dim myCHANNEL_CODE As String = ""
        Dim myPAGER_FORMAT As String = ""
        Dim myCAPCODE_IND As String = ""



        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing
        'Limit how many records processed so other processes dont timeout. 
        While myContinue = True And myCount < myLimit

            mySql = "execute dbo.PagerGCTransferGet"
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myAdapter
                Try
                   
                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using
                    If myData.Rows.Count = 0 Then
                        Exit Sub
                    End If
                    'Keep the key to send back to delete if update to Abacus is completed properly
                    mySERIAL_NO = myData.Rows(0).Item("SERIAL_NO")
                    mySLOT = myData.Rows(0).Item("SLOT")
                    myYYYYMMDDHHMISS = myData.Rows(0).Item("YYYYMMDDHHMISS")
                    mySEQ_NO = myData.Rows(0).Item("SEQ_NO")
                    myCHANNEL_CODE = myData.Rows(0).Item("CHANNEL_CODE")
                    myPAGER_FORMAT = myData.Rows(0).Item("PAGER_FORMAT")
                    myCAPCODE_IND = myData.Rows(0).Item("CAPCODE_IND")
                  

                    myData = Nothing
                    myAdapter = Nothing
                Catch ex As Exception
                    ' MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                    WriteToErrorEmail(ex.Message.ToString)
                End Try
            End With

            myURIRequest = myURIBase & "<MSG>" & HTA(mySERIAL_NO) & "/" _
                                                & HTA(mySLOT) & "/" _
                                                & HTA(myYYYYMMDDHHMISS) & "/" _
                                                & HTA(mySEQ_NO) & "/" _
                                                & HTA(myCHANNEL_CODE) & "/" _
                                                & HTA(myPAGER_FORMAT) & "/" _
                                                & HTA(myCAPCODE_IND) & "</MSG>"

            myReadCue = SendHTTPRequest(myURIRequest, myURIResponse)

            If myReadCue = True Then
                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                    Dim P1 As Integer = 0
                    Dim P2 As Integer = 0
                    Dim myError As String = myURIResponse
                    While myError.IndexOf("<ERROR>") >= 0
                        P1 = myError.IndexOf("<ERROR>") + 7
                        P2 = myError.LastIndexOf("</ERROR>")
                        myError = myError.Substring(P1, P2 - P1)
                    End While
                    AppendText("XML Error: " & myError)
                    WriteToErrorEmail(myError)
                End If
            Else
                WriteToErrorEmail("Failed " & myURIRequest)
            End If


            mySql = "execute dbo.PagerGCTransferHistoryUpdate @SERIAL_NO, @SLOT, @YYYYMMDDHHMISS, @SEQ_NO"
            Dim myDoneAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myDoneAdapter
                Try
                    .SelectCommand.Parameters.Add("@SERIAL_NO", SqlDbType.VarChar).Value = mySERIAL_NO
                    .SelectCommand.Parameters.Add("@SLOT", SqlDbType.Char).Value = mySLOT
                    .SelectCommand.Parameters.Add("@YYYYMMDDHHMISS", SqlDbType.Char).Value = myYYYYMMDDHHMISS
                    .SelectCommand.Parameters.Add("@SEQ_NO", SqlDbType.Char).Value = mySEQ_NO

                    myData = New DataTable("Response")
                    Using myDoneAdapter
                        .Fill(myData)
                    End Using
                    ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                    If ReturnMsg <> "." Then
                        WriteToErrorEmail("PagerGCTransferHistoryUpdate Failed " & myURIRequest & vbCrLf & _
                                  "SERIAL_NO: " & mySERIAL_NO & "|SLOT:" & mySLOT & "|YYYYMMDDHHMISS:" & myYYYYMMDDHHMISS * ":SEQ: " & myYYYYMMDDHHMISS)
                    End If
                    myData = Nothing
                    myDoneAdapter = Nothing
                Catch ex As Exception

                    WriteToErrorEmail(ex.Message.ToString)
                End Try
            End With
            myCount = myCount + 1

        End While

    End Sub
    Private Sub SendHarkNumbersToAdd(ByRef pDS As DataSet, ByRef pRecords As Integer)
        Dim myDataTable As DataTable = Nothing
        Dim myTranDate As String = ""
        Dim myHarkDate As String = ""
        Dim myTranTime As String = ""
        Dim myTranDateTime As DateTime = Nothing
        Dim myPhoneNo As String = ""
        Dim myAlphaOrNumeric As String = ""
        Dim myInService As String = ""
        Dim myEncrypted As String = ""
        Dim myPrimCapcode As String = ""
        Dim myCoCode As String = ""
        Dim myAcctNo As String = ""
        Dim mySystemID As String = ""
        Dim mySwitchType As String = ""
        Dim myInterfaced As String = ""
        Dim myInServEquipType As String = ""
        Dim myIUD As String = ""
        Dim myPagerTypeCode As String = ""
        Dim myOnMachineAvailable As String = ""
        Dim myReserved As String = ""
        Dim myReservedCoCode As String = ""
        Dim myReservedAcctNo As String = ""
        Dim myAlias As String = ""
        Dim myHolderName As String = ""
        Dim myWorkNumber As String = ""
        Dim myWorkDate As String = ""
        Dim mySQL As String = ""
        Dim myData As DataTable = Nothing
        Dim myDeleteData As DataTable = Nothing
        Dim myResponse As String = ""
        Dim myKeyDT As DataTable = Nothing
        Dim myKeyDTRow As DataRow = Nothing
        Dim myKeyData As String = ""
        Dim myKeyDS As DataSet = Nothing
        Dim P1 As Integer = 0
        Dim p2 As Integer = 0
        Dim mySendSuccess As Boolean = False

        'Dim myURIData As String = ""
        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=DELETE+WILHARKT+"   'Dev outward facing IP - Use after testing new delete link functionality
        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=METHOD+WIQHARKT+"    'Dev outward facing IP - Use normally
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=METHOD+WIQHARKT+"  'dev inside IP
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=DELETE+WILHARKT+"  'dev inside IP - Use after testing new delete link functionality
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=METHOD+WIQHARKT+"      'BamBam\test IP
        'Dim myURIBASE As String = "http://10.200.2.50/xml.tbred?xmlrequest=DELETE+WILHARKT+"   'Phred2 IP

        'Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIQHARKT+"
        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIDHARKT+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Try
            If pDS.Tables.Count > 0 Then
                myDataTable = pDS.Tables(0)
                For Each myRow As DataRow In myDataTable.Rows
                    myTranDate = myRow.Item("ACTUAL_TRAN_DATE").ToString
                    myTranTime = myRow.Item("ACTUAL_TRAN_TIME").ToString
                    myPhoneNo = myRow.Item("PHONE_NO").ToString
                    myAlphaOrNumeric = myRow.Item("ALPHA_OR_NUMERIC").ToString
                    myInService = myRow.Item("IN_SERVICE").ToString
                    myEncrypted = myRow.Item("ENCRYPTED").ToString
                    myPrimCapcode = myRow.Item("PRIM_CAPCODE").ToString
                    myCoCode = myRow.Item("CO_CODE").ToString
                    myAcctNo = myRow.Item("ACCT_NO").ToString
                    mySystemID = myRow.Item("SYSTEM_ID").ToString
                    mySwitchType = myRow.Item("SWITCH_TYPE").ToString
                    myInterfaced = myRow.Item("INTERFACED").ToString
                    myInServEquipType = myRow.Item("IN_SERV_EQUIP_TYPE").ToString
                    myIUD = myRow.Item("I_U_D").ToString
                    myPagerTypeCode = myRow.Item("PAGER_TYPE_CODE").ToString
                    myOnMachineAvailable = myRow.Item("ON_MACHINE_AVAILABLE").ToString
                    myReserved = myRow.Item("RESERVED").ToString
                    myReservedCoCode = myRow.Item("RESERVED_CO_CODE").ToString
                    myReservedAcctNo = myRow.Item("RESERVED_ACCT_NO").ToString
                    myAlias = myRow.Item("ALIAS").ToString
                    myHolderName = myRow.Item("HOLDER_NAME").ToString
                    If myPhoneNo.Trim <> "" Then
                        myWorkNumber = ""
                        For ch = 0 To Len(myPhoneNo) - 1
                            If IsNumeric(myPhoneNo.Substring(ch, 1)) Then
                                myWorkNumber = myWorkNumber & myPhoneNo.Substring(ch, 1)
                            End If
                        Next
                        If myWorkNumber.Trim <> "" Then
                            myPhoneNo = myWorkNumber
                        End If
                    End If
                    If myTranDate.Length <> 10 Or myTranTime.Length <> 6 Then
                        myTranDateTime = Now
                    Else
                        myWorkDate = myTranDate.Substring(6, 4) & "-" & myTranDate.Substring(0, 2) & "-" & myTranDate.Substring(3, 2)
                        myWorkDate = myWorkDate & " " & myTranTime.Substring(0, 2) & ":" & myTranTime.Substring(2, 2) & ":" & myTranTime.Substring(4, 2) & "." & "000"
                        'myTranDateTime = DirectCast(myWorkDate, DateTime)
                        myTranDateTime = DirectCast(myWorkDate, String)
                        myWorkDate = ""
                        myHarkDate = myTranDate.Substring(0, 2) & myTranDate.Substring(3, 2) & myTranDate.Substring(8, 2)
                    End If
                    myKeyDT = New DataTable("KeyTable")
                    myKeyDT.Columns.Add("ACTUAL-TRAN-DATE", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("ACTUAL-TRAN-TIME", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("PHONE-NO", Type.GetType("System.String"))
                    myKeyDTRow = myKeyDT.NewRow
                    'myKeyDTRow.Item("ACTUAL-TRAN-DATE") = myTranDate
                    myKeyDTRow.Item("ACTUAL-TRAN-DATE") = myHarkDate
                    myKeyDTRow.Item("ACTUAL-TRAN-TIME") = myTranTime
                    myKeyDTRow.Item("PHONE-NO") = myPhoneNo
                    myKeyDT.Rows.Add(myKeyDTRow)
                    myKeyDS = New DataSet("KeyDS")
                    myKeyDS.Tables.Add(myKeyDT)

                    myKeyData = ConvertDataSetToXML(myKeyDS)

                    P1 = myKeyData.IndexOf("<KeyTable>") + 10
                    p2 = myKeyData.LastIndexOf("</KeyTable>")
                    myKeyData = myKeyData.Substring(P1, p2 - P1)

                    mySQL = "execute dbo.HarkPhoneNo @ActualTranDateTime, @PhoneNo, @AlphaOrNumeric, @IUD, @InService, @Encrypted, @PrimCapcode, @CoCode, @AcctNo, @SystemID, @SwitchType, @Interfaced, @InServEquipType, @PagerTypeCode, @OnMachineAvailable, @Reserved, @ReservedCoCode, @ReservedAcctNo, @Alias, @HolderName"
                    Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                    With myAdapter
                        Try
                            .SelectCommand.Parameters.Add("@ActualTranDateTime", SqlDbType.DateTime).Value = myTranDateTime
                            .SelectCommand.Parameters.Add("@PhoneNo", SqlDbType.Char).Value = myPhoneNo
                            .SelectCommand.Parameters.Add("@AlphaOrNumeric", SqlDbType.Char).Value = myAlphaOrNumeric
                            .SelectCommand.Parameters.Add("@IUD", SqlDbType.Char).Value = myIUD
                            .SelectCommand.Parameters.Add("@InService", SqlDbType.Char).Value = myInService
                            .SelectCommand.Parameters.Add("@Encrypted", SqlDbType.Char).Value = myEncrypted
                            .SelectCommand.Parameters.Add("@PrimCapcode", SqlDbType.Char).Value = myPrimCapcode
                            .SelectCommand.Parameters.Add("@CoCode", SqlDbType.Char).Value = myCoCode
                            .SelectCommand.Parameters.Add("@AcctNo", SqlDbType.Char).Value = myAcctNo
                            .SelectCommand.Parameters.Add("@SystemID", SqlDbType.Char).Value = mySystemID
                            .SelectCommand.Parameters.Add("@SwitchType", SqlDbType.Char).Value = mySwitchType
                            .SelectCommand.Parameters.Add("@Interfaced", SqlDbType.Char).Value = myInterfaced
                            .SelectCommand.Parameters.Add("@InServEquipType", SqlDbType.Char).Value = myInServEquipType
                            .SelectCommand.Parameters.Add("@PagerTypeCode", SqlDbType.Char).Value = myPagerTypeCode
                            .SelectCommand.Parameters.Add("@OnMachineAvailable", SqlDbType.Char).Value = myOnMachineAvailable
                            .SelectCommand.Parameters.Add("@Reserved", SqlDbType.Char).Value = myReserved
                            .SelectCommand.Parameters.Add("@ReservedCoCode", SqlDbType.Char).Value = myReservedCoCode
                            .SelectCommand.Parameters.Add("@ReservedAcctNo", SqlDbType.Char).Value = myReservedAcctNo
                            .SelectCommand.Parameters.Add("@Alias", SqlDbType.Char).Value = myAlias
                            .SelectCommand.Parameters.Add("@HolderName", SqlDbType.Char).Value = myHolderName

                            myData = New DataTable("Response")
                            Using myAdapter
                                .Fill(myData)
                            End Using

                            If myData.Rows.Count > 0 Then
                                myResponse = myData.Rows(0).Item("ReturnMsg").ToString.Trim
                            Else
                                myResponse = "No Response"
                            End If
                            AppendText("HarkNumber Response: " & myResponse)

                            'If myResponse = "." Then
                            'AppendText("HARK Number Add to IntelliMessage Successful")

                            myURIRequest = myURIBase & "<msg>" & myKeyData & "</msg>"

                            'myURIRequest = myURIBASE & myKeyData
                            mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

                            If mySendSuccess = True Then
                                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                                    Dim P3 As Integer = 0
                                    Dim P4 As Integer = 0
                                    Dim myError As String = myURIResponse
                                    While myError.IndexOf("<ERROR>") >= 0
                                        P3 = myError.IndexOf("<ERROR>") + 7
                                        P4 = myError.LastIndexOf("</ERROR>")
                                        myError = myError.Substring(P3, P4 - P3)
                                    End While
                                    AppendText("XML Error: " & myError)
                                    pDS = Nothing
                                Else
                                    'Dim dstHARKTXMLData As New DataSet()
                                    'dstHARKTXMLData = ConvertXMLToDataSet(myURIResponse)
                                    'pDS = dstHARKTXMLData
                                    Dim myHarkDeleteDS As New DataSet()
                                    myHarkDeleteDS = ConvertXMLToDataSet(myURIResponse)
                                    If myHarkDeleteDS.Tables.Count > 0 Then
                                        myDeleteData = myHarkDeleteDS.Tables(0)

                                        If myDeleteData.Rows.Count > 0 Then
                                            myResponse = myDeleteData.Rows(0).Item("DELETE").ToString.Trim
                                        Else
                                            myResponse = "No Response"
                                        End If

                                        If myResponse = "." Then
                                            AppendText("Delete Successful")
                                        Else
                                            AppendText("Delete Not Successful")
                                        End If
                                    End If
                                End If
                            Else
                                pDS = Nothing
                            End If

                            'End If

                            pRecords = pRecords + 1

                        Catch ex As Exception
                            'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                            WriteToErrorEmail(ex.Message.ToString)
                        End Try

                    End With

                Next
            End If
        Catch ex As Exception
            pDS = Nothing
        End Try
    End Sub

    Private Sub CheckForIntelliMsgDisconnects(ByRef pIntelliMsgDisconnects As Boolean)

        'Dim myURIBase As String = "http://63.97.58.99/xml.tbred?xmlrequest=METHOD+WIMCCAPW"     'Dev outward facing IP - Use normally
        'Dim myURIBase As String = "http://10.214.34.75/xml.tbred?xmlrequest=METHOD+WIMCCAPW"   'dev inside IP
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=METHOD+WIMCCAPW"      'BamBam\test IP

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCCAPW"
        Dim mySendSuccess As Boolean = False
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        myURIRequest = myURIBase

        mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

        If mySendSuccess = True Then
            If myURIResponse = "Y" Then
                pIntelliMsgDisconnects = True
            End If
        End If

    End Sub

    Private Sub ReadIntelliMsgDisconnects(ByRef pImsgDiscoDS As DataSet)
        Dim mySendSuccess As Boolean = False
        Dim myURIData As String = ""

        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=READ+WIFCCAPD"   'Dev outward facing IP - Use normally
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=READ+WIFCCAPD"     'dev inside IP
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=READ+WIFCCAPD"      'BamBam\test IP

        Dim myURIBase As String = gURIBase & "xmlrequest=READ+WIFCCAPD"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        myURIRequest = myURIBase

        mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

        If mySendSuccess = True Then
            If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                Dim P1 As Integer = 0
                Dim P2 As Integer = 0
                Dim myError As String = myURIResponse
                While myError.IndexOf("<ERROR>") >= 0
                    P1 = myError.IndexOf("<ERROR>") + 7
                    P2 = myError.LastIndexOf("</ERROR>")
                    myError = myError.Substring(P1, P2 - P1)
                End While
                AppendText("XML Error: " & myError)
                pImsgDiscoDS = Nothing
            Else
                Dim dstImsgDiscoXMLData As New DataSet()
                dstImsgDiscoXMLData = ConvertXMLToDataSet(myURIResponse)
                pImsgDiscoDS = dstImsgDiscoXMLData
            End If
        Else
            pImsgDiscoDS = Nothing
        End If
    End Sub

    Private Sub SendIntelliMsgDisconnects(ByRef pDS As DataSet, ByRef pRecords As Integer)
        Dim myDataTable As DataTable = Nothing
        Dim myTranDate As String = ""
        Dim myIntelliMsgDisoDate As String = ""
        Dim myTranTime As String = ""
        Dim myTranDateTime As DateTime = Nothing
        Dim myCoCode As String = ""
        Dim myAcctNo As String = ""
        Dim myPhoneNo As String = ""
        Dim myDestination As String = ""
        Dim myDateOutService As String = ""
        Dim myOperCode As String = ""
        Dim myWorkNumber As String = ""
        Dim myWorkDate As String = ""
        Dim mySQL As String = ""
        Dim myData As DataTable = Nothing
        Dim myDeleteData As DataTable = Nothing
        Dim myResponse As String = ""
        Dim myKeyDT As DataTable = Nothing
        Dim myKeyDTRow As DataRow = Nothing
        Dim myKeyData As String = ""
        Dim myKeyDS As DataSet = Nothing
        Dim P1 As Integer = 0
        Dim p2 As Integer = 0
        Dim mySendSuccess As Boolean = False

        'Dim myURIData As String = ""
        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=DELETE+WILCCAPD+"       'Dev outward facing IP - Use after testing delete link functionality
        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=METHOD+WIQCCAPD+"        'Dev outward facing IP - Use normally
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=METHOD+WIQCCAPD+"      'dev inside IP
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=DELETE+WILCCAPD+"      'dev inside IP - Use after testing delete link functionality
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=METHOD+WIQCCAPD+"      'BamBam\test IP
        'Dim myURIBASE As String = "http://10.200.2.50/xml.tbred?xmlrequest=DELETE+WILCCAPD+"       'Phred2 IP

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIQCCAPD+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Try
            If pDS.Tables.Count > 0 Then
                myDataTable = pDS.Tables(0)
                For Each myRow As DataRow In myDataTable.Rows
                    myTranDate = myRow.Item("ACTUAL_TRAN_DATE").ToString
                    myTranTime = myRow.Item("ACTUAL_TRAN_TIME").ToString
                    myCoCode = myRow.Item("CO_CODE").ToString
                    myAcctNo = myRow.Item("ACCT_NO").ToString
                    myPhoneNo = myRow.Item("PHONE_NO").ToString
                    myDestination = myRow.Item("DESTINATION").ToString
                    myDateOutService = myRow.Item("DATE_OUT_SERVICE").ToString
                    myOperCode = myRow.Item("OPER_CODE").ToString
                    If myPhoneNo.Trim <> "" Then
                        myWorkNumber = ""
                        For ch = 0 To Len(myPhoneNo) - 1
                            If IsNumeric(myPhoneNo.Substring(ch, 1)) Then
                                myWorkNumber = myWorkNumber & myPhoneNo.Substring(ch, 1)
                            End If
                        Next
                        If myWorkNumber.Trim <> "" Then
                            myPhoneNo = myWorkNumber
                        End If
                    End If
                    If myTranDate.Length <> 10 Or myTranTime.Length <> 6 Then
                        myTranDateTime = Now
                    Else
                        myWorkDate = myTranDate.Substring(6, 4) & "-" & myTranDate.Substring(0, 2) & "-" & myTranDate.Substring(3, 2)
                        myWorkDate = myWorkDate & " " & myTranTime.Substring(0, 2) & ":" & myTranTime.Substring(2, 2) & ":" & myTranTime.Substring(4, 2) & "." & "000"
                        'myTranDateTime = DirectCast(myWorkDate, DateTime)
                        myTranDateTime = DirectCast(myWorkDate, String)
                        myWorkDate = ""
                        myIntelliMsgDisoDate = myTranDate.Substring(0, 2) & myTranDate.Substring(3, 2) & myTranDate.Substring(8, 2)
                    End If
                    myKeyDT = New DataTable("KeyTable")
                    myKeyDT.Columns.Add("ACTUAL-TRAN-DATE", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("ACTUAL-TRAN-TIME", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("CO-CODE", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("ACCT-NO", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("PHONE-NO", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("DESTINATION", Type.GetType("System.String"))
                    'myKeyDT.Columns.Add("DATE-OUT-SERVICE", Type.GetType("System.String"))
                    'myKeyDT.Columns.Add("OPER-CODE", Type.GetType("System.String"))
                    myKeyDTRow = myKeyDT.NewRow
                    myKeyDTRow.Item("ACTUAL-TRAN-DATE") = myIntelliMsgDisoDate
                    myKeyDTRow.Item("ACTUAL-TRAN-TIME") = myTranTime
                    myKeyDTRow.Item("CO-CODE") = myCoCode
                    myKeyDTRow.Item("ACCT-NO") = myAcctNo
                    myKeyDTRow.Item("PHONE-NO") = myPhoneNo
                    myKeyDTRow.Item("DESTINATION") = myDestination
                    'myKeyDTRow.Item("DATE-OUT-SERVICE") = myDateOutService
                    'myKeyDTRow.Item("OPER-CODE") = myOperCode
                    myKeyDT.Rows.Add(myKeyDTRow)
                    myKeyDS = New DataSet("KeyDS")
                    myKeyDS.Tables.Add(myKeyDT)

                    myKeyData = ConvertDataSetToXML(myKeyDS)

                    P1 = myKeyData.IndexOf("<KeyTable>") + 10
                    p2 = myKeyData.LastIndexOf("</KeyTable>")
                    myKeyData = myKeyData.Substring(P1, p2 - P1)


                    mySQL = "execute dbo.ProcessAbacusDisconnects @ActualTranDateTime, @CoCode, @AcctNo, @PhoneNo, @Destination, @DateOutService, @OperCode"
                    Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                    With myAdapter
                        Try
                            .SelectCommand.Parameters.Add("@ActualTranDateTime", SqlDbType.DateTime).Value = myTranDateTime
                            .SelectCommand.Parameters.Add("@CoCode", SqlDbType.Char).Value = myCoCode
                            .SelectCommand.Parameters.Add("@AcctNo", SqlDbType.Char).Value = myAcctNo
                            .SelectCommand.Parameters.Add("@PhoneNo", SqlDbType.Char).Value = myPhoneNo
                            .SelectCommand.Parameters.Add("@Destination", SqlDbType.Char).Value = myDestination
                            .SelectCommand.Parameters.Add("@DateOutService", SqlDbType.DateTime).Value = myDateOutService
                            .SelectCommand.Parameters.Add("@OperCode", SqlDbType.Char).Value = myOperCode

                            myData = New DataTable("Response")
                            Using myAdapter
                                .Fill(myData)
                            End Using

                            If myData.Rows.Count > 0 Then
                                myResponse = myData.Rows(0).Item("ReturnMsg").ToString.Trim
                            Else
                                myResponse = "No Response"
                            End If
                            AppendText("ProcessAbacusDisconnects Response: " & myResponse)

                            'If myResponse = "." Then
                            'AppendText("HARK Number Add to IntelliMessage Successful")

                            myURIRequest = myURIBase & "<msg>" & myKeyData & "</msg>"

                            'myURIRequest = myURIBASE & myKeyData
                            mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

                            If mySendSuccess = True Then
                                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                                    Dim P3 As Integer = 0
                                    Dim P4 As Integer = 0
                                    Dim myError As String = myURIResponse
                                    While myError.IndexOf("<ERROR>") >= 0
                                        P3 = myError.IndexOf("<ERROR>") + 7
                                        P4 = myError.LastIndexOf("</ERROR>")
                                        myError = myError.Substring(P3, P4 - P3)
                                    End While
                                    AppendText("XML Error: " & myError)
                                    pDS = Nothing
                                Else
                                    Dim myIntelliMsgDeleteDS As New DataSet()
                                    myIntelliMsgDeleteDS = ConvertXMLToDataSet(myURIResponse)
                                    If myIntelliMsgDeleteDS.Tables.Count > 0 Then
                                        myDeleteData = myIntelliMsgDeleteDS.Tables(0)

                                        If myDeleteData.Rows.Count > 0 Then
                                            myResponse = myDeleteData.Rows(0).Item("DELETE").ToString.Trim
                                        Else
                                            myResponse = "No Response"
                                        End If

                                        If myResponse = "." Then
                                            AppendText("Delete Successful")
                                        Else
                                            AppendText("Delete Not Successful")
                                        End If
                                    End If
                                End If
                            Else
                                pDS = Nothing
                            End If

                            'End If

                            pRecords = pRecords + 1

                        Catch ex As Exception
                            'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                            WriteToErrorEmail(ex.Message.ToString)
                        End Try

                    End With

                Next
            End If
        Catch ex As Exception
            pDS = Nothing
        End Try
    End Sub

    Private Sub CheckXMLHeartbeat(ByVal pHeartBeatURIBase As String)
        Dim myDataTable As DataTable = Nothing
        Dim mySendSuccess As Boolean = False
        Dim myURIBase As String = pHeartBeatURIBase & "xmlrequest=METHOD+WIMXMLHB" + ""
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        myURIRequest = myURIBase

        mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

        If mySendSuccess = True Then
            If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                Dim P3 As Integer = 0
                Dim P4 As Integer = 0
                Dim myError As String = myURIResponse
                While myError.IndexOf("<ERROR>") >= 0
                    P3 = myError.IndexOf("<ERROR>") + 7
                    P4 = myError.LastIndexOf("</ERROR>")
                    myError = myError.Substring(P3, P4 - P3)
                End While
                AppendText(pHeartBeatURIBase.ToString & "XML Heartbeat Error: " & myError)
                myURIResponse = myError
                'pWIFIMSGU = Nothing
            Else
                If myURIResponse = "." Then
                    AppendText(pHeartBeatURIBase.ToString & "XML Heartbeat Successful")
                End If
            End If
        End If

        If mySendSuccess = False Then
            myURIResponse = myExceptionText
        End If

        If myURIResponse <> "." Or myTimeOut = True Then
            SendHeartbeatErrorEmail(myURIResponse, pHeartBeatURIBase)
        End If
    End Sub

    Private Sub SendHeartbeatErrorEmail(ByVal pURIResponse As String, ByVal pMachine As String)
        Dim myDataTable As DataTable = Nothing
        Dim mySendSuccess As Boolean = False
        Dim emailAddress As String = ""
        Dim myEmailStatement As String = pURIResponse
        Dim myFirstEmail As Integer = 1
        Dim mySQL As String = ""
        Dim myData As DataTable = Nothing

        If myTimeOut = True Then
            If myEmailSentTime > Now Then
                Return
            End If
            myEmailStatement = "No Response from Server"
        End If

        myEmailStatement = myEmailStatement & vbCrLf & vbCrLf & "XML Heartbeat Error"

        myDataTable = New DataTable
        myDataTable.Columns.Add("EmailAddress", Type.GetType("System.String"))
        Dim myNewRow As DataRow = myDataTable.NewRow
        myNewRow("EmailAddress") = "larryr@abw.com; mac.mcmahan@americanmessaging.net"
        'myNewRow("EmailAddress") = "larryr@abw.com"
        myDataTable.Rows.Add(myNewRow)
        Dim myNewRow1 As DataRow = myDataTable.NewRow
        myNewRow1("EmailAddress") = " 2147071687@txt.att.net; 2143943784@txt.att.net; 2142740019@txt.att.net"
        'myNewRow1("EmailAddress") = " 2147071687@txt.att.net"
        myDataTable.Rows.Add(myNewRow1)
        Dim myNewRow2 As DataRow = myDataTable.NewRow
        myNewRow2("EmailAddress") = "zerrialb@intellimsg.net; joeb@intellimsg.net; larryr@intellimsg.net"
        'myNewRow2("EmailAddress") = "larryr@intellimsg.net"
        myDataTable.Rows.Add(myNewRow2)

        For Each myRow As DataRow In myDataTable.Rows
            emailAddress = myRow.Item("EmailAddress").ToString.Trim
            If emailAddress <> "" Then
                If emailAddress <> "administrator@dev.intellimsg.net" Then
                    mySQL = "execute dbo.SendEmail @From, @To, @Subject, @Body, @ReturnMsg"
                    Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                    With myAdapter
                        Try
                            .SelectCommand.Parameters.Add("@From", SqlDbType.VarChar).Value = "administrator@intellimsg.net"
                            .SelectCommand.Parameters.Add("@To", SqlDbType.Char).Value = emailAddress
                            .SelectCommand.Parameters.Add("@Subject", SqlDbType.Char).Value = "XML Heartbeat"
                            .SelectCommand.Parameters.Add("@Body", SqlDbType.Char).Value = "The XML Heartbeat has failed for " & pMachine & "." & vbCrLf & vbCrLf & myEmailStatement
                            .SelectCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar).Value = ""

                            myData = New DataTable("Response")
                            Using myAdapter
                                .Fill(myData)
                            End Using
                        Catch ex As Exception
                            AppendText(ex.Message.ToString)
                        End Try
                    End With
                End If
            End If
        Next

        If pURIResponse = "No Reply" Then
            mySQL = "execute dbo.SendEmail @From, @To, @Subject, @Body, @ReturnMsg"
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
            With myAdapter
                Try
                    .SelectCommand.Parameters.Add("@From", SqlDbType.VarChar).Value = "administrator@intellimsg.net"
                    .SelectCommand.Parameters.Add("@To", SqlDbType.Char).Value = "2147071687@txt.att.net; 2147256985@txt.att.net"
                    .SelectCommand.Parameters.Add("@Subject", SqlDbType.Char).Value = "OOXMLSRV"
                    .SelectCommand.Parameters.Add("@Body", SqlDbType.Char).Value = "The OOXMLSRV process has stopped on Prod2"
                    .SelectCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar).Value = ""

                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using
                Catch ex As Exception
                    AppendText(ex.Message.ToString)
                End Try
            End With
        End If
    End Sub

    Private Sub CheckForIntelliMsgPSChange(ByRef pIntelliMsgPSChanges As Boolean)

        'Dim myURIBase As String = "http://63.97.58.99/xml.tbred?xmlrequest=METHOD+WIMCCAPW"     'Dev outward facing IP - Use normally
        'Dim myURIBase As String = "http://10.214.34.75/xml.tbred?xmlrequest=METHOD+WIMCCAPW"   'dev inside IP
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=METHOD+WIMCCPSW"      'BamBam\test IP

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCCPSW"
        Dim mySendSuccess As Boolean = False
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        myURIRequest = myURIBase

        mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

        If mySendSuccess = True Then
            If myURIResponse = "Y" Then
                pIntelliMsgPSChanges = True
            End If
        End If

    End Sub

    Private Sub ReadIntelliMsgPSChange(ByRef pImsgPSChangeDS As DataSet)
        Dim mySendSuccess As Boolean = False
        Dim myURIData As String = ""

        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=READ+WIFCCAPD"   'Dev outward facing IP - Use normally
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=READ+WIFCCAPD"     'dev inside IP
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=READ+WIFCCPSD"      'BamBam\test IP

        Dim myURIBase As String = gURIBase & "xmlrequest=READ+WIFCCPSD"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        myURIRequest = myURIBase

        mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

        If mySendSuccess = True Then
            If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                Dim P1 As Integer = 0
                Dim P2 As Integer = 0
                Dim myError As String = myURIResponse
                While myError.IndexOf("<ERROR>") >= 0
                    P1 = myError.IndexOf("<ERROR>") + 7
                    P2 = myError.LastIndexOf("</ERROR>")
                    myError = myError.Substring(P1, P2 - P1)
                End While
                AppendText("XML Error: " & myError)
                pImsgPSChangeDS = Nothing
            Else
                Dim dstImsgPSChangeXMLData As New DataSet()
                dstImsgPSChangeXMLData = ConvertXMLToDataSet(myURIResponse)
                pImsgPSChangeDS = dstImsgPSChangeXMLData
            End If
        Else
            pImsgPSChangeDS = Nothing
        End If
    End Sub

    Private Sub SendIntelliMsgPSChanges(ByRef pDS As DataSet, ByRef pRecords As Integer)
        Dim myDataTable As DataTable = Nothing
        Dim myTranDate As String = ""
        Dim myIntelliMsgPSChangeDate As String = ""
        Dim myTranTime As String = ""
        Dim myTranDateTime As DateTime = Nothing
        Dim myCoCode As String = ""
        Dim myAcctNo As String = ""
        Dim myPhoneNo As String = ""
        Dim myDestination As String = ""
        Dim myOldPhoneNo As String = ""
        Dim myOperCode As String = ""
        Dim myWorkNumber As String = ""
        Dim myWorkDate As String = ""
        Dim mySQL As String = ""
        Dim myData As DataTable = Nothing
        Dim myDeleteData As DataTable = Nothing
        Dim myResponse As String = ""
        Dim myKeyDT As DataTable = Nothing
        Dim myKeyDTRow As DataRow = Nothing
        Dim myKeyData As String = ""
        Dim myKeyDS As DataSet = Nothing
        Dim P1 As Integer = 0
        Dim p2 As Integer = 0
        Dim mySendSuccess As Boolean = False

        'Dim myURIData As String = ""
        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=DELETE+WILCCAPD+"       'Dev outward facing IP - Use after testing delete link functionality
        'Dim myURIBASE As String = "http://63.97.58.99/xml.tbred?xmlrequest=METHOD+WIQCCAPD+"        'Dev outward facing IP - Use normally
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=METHOD+WIQCCAPD+"      'dev inside IP
        'Dim myURIBASE As String = "http://10.214.34.75/xml.tbred?xmlrequest=DELETE+WILCCAPD+"      'dev inside IP - Use after testing delete link functionality
        'Dim myURIBase As String = "http://10.200.2.70/testxml.tbred?xmlrequest=METHOD+WIQCCPSD+"      'BamBam\test IP
        'Dim myURIBASE As String = "http://10.200.2.50/xml.tbred?xmlrequest=DELETE+WILCCAPD+"       'Phred2 IP

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIQCCPSD+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Try
            If pDS.Tables.Count > 0 Then
                myDataTable = pDS.Tables(0)
                For Each myRow As DataRow In myDataTable.Rows
                    myTranDate = myRow.Item("ACTUAL_TRAN_DATE").ToString
                    myTranTime = myRow.Item("ACTUAL_TRAN_TIME").ToString
                    myCoCode = myRow.Item("CO_CODE").ToString
                    myAcctNo = myRow.Item("ACCT_NO").ToString
                    myPhoneNo = myRow.Item("PHONE_NO").ToString
                    myDestination = myRow.Item("DESTINATION").ToString
                    myOldPhoneNo = myRow.Item("OLD_PHONE_NO").ToString
                    myOperCode = myRow.Item("OPER_CODE").ToString
                    If myPhoneNo.Trim <> "" Then
                        myWorkNumber = ""
                        For ch = 0 To Len(myPhoneNo) - 1
                            If IsNumeric(myPhoneNo.Substring(ch, 1)) Then
                                myWorkNumber = myWorkNumber & myPhoneNo.Substring(ch, 1)
                            End If
                        Next
                        If myWorkNumber.Trim <> "" Then
                            myPhoneNo = myWorkNumber
                        End If
                    End If
                    If myOldPhoneNo.Trim <> "" Then
                        myWorkNumber = ""
                        For ch = 0 To Len(myOldPhoneNo) - 1
                            If IsNumeric(myOldPhoneNo.Substring(ch, 1)) Then
                                myWorkNumber = myWorkNumber & myOldPhoneNo.Substring(ch, 1)
                            End If
                        Next
                        If myWorkNumber.Trim <> "" Then
                            myOldPhoneNo = myWorkNumber
                        End If
                    End If
                    If myTranDate.Length <> 10 Or myTranTime.Length <> 6 Then
                        myTranDateTime = Now
                    Else
                        myWorkDate = myTranDate.Substring(6, 4) & "-" & myTranDate.Substring(0, 2) & "-" & myTranDate.Substring(3, 2)
                        myWorkDate = myWorkDate & " " & myTranTime.Substring(0, 2) & ":" & myTranTime.Substring(2, 2) & ":" & myTranTime.Substring(4, 2) & "." & "000"
                        'myTranDateTime = DirectCast(myWorkDate, DateTime)
                        myTranDateTime = DirectCast(myWorkDate, String)
                        myWorkDate = ""
                        myIntelliMsgPSChangeDate = myTranDate.Substring(0, 2) & myTranDate.Substring(3, 2) & myTranDate.Substring(8, 2)
                    End If
                    myKeyDT = New DataTable("KeyTable")
                    myKeyDT.Columns.Add("ACTUAL-TRAN-DATE", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("ACTUAL-TRAN-TIME", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("CO-CODE", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("ACCT-NO", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("PHONE-NO", Type.GetType("System.String"))
                    myKeyDT.Columns.Add("DESTINATION", Type.GetType("System.String"))
                    myKeyDTRow = myKeyDT.NewRow
                    myKeyDTRow.Item("ACTUAL-TRAN-DATE") = myIntelliMsgPSChangeDate
                    myKeyDTRow.Item("ACTUAL-TRAN-TIME") = myTranTime
                    myKeyDTRow.Item("CO-CODE") = myCoCode
                    myKeyDTRow.Item("ACCT-NO") = myAcctNo
                    myKeyDTRow.Item("PHONE-NO") = myPhoneNo
                    myKeyDTRow.Item("DESTINATION") = myDestination
                    myKeyDT.Rows.Add(myKeyDTRow)
                    myKeyDS = New DataSet("KeyDS")
                    myKeyDS.Tables.Add(myKeyDT)

                    myKeyData = ConvertDataSetToXML(myKeyDS)

                    P1 = myKeyData.IndexOf("<KeyTable>") + 10
                    p2 = myKeyData.LastIndexOf("</KeyTable>")
                    myKeyData = myKeyData.Substring(P1, p2 - P1)


                    mySQL = "execute dbo.ProcessAbacusPhoneNumberChanges @ActualTranDateTime, @CoCode, @AcctNo, @PhoneNo, @Destination, @OldPhoneNo, @OperCode"
                    Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                    With myAdapter
                        Try
                            .SelectCommand.Parameters.Add("@ActualTranDateTime", SqlDbType.DateTime).Value = myTranDateTime
                            .SelectCommand.Parameters.Add("@CoCode", SqlDbType.Char).Value = myCoCode
                            .SelectCommand.Parameters.Add("@AcctNo", SqlDbType.Char).Value = myAcctNo
                            .SelectCommand.Parameters.Add("@PhoneNo", SqlDbType.Char).Value = myPhoneNo
                            .SelectCommand.Parameters.Add("@Destination", SqlDbType.Char).Value = myDestination
                            .SelectCommand.Parameters.Add("@OldPhoneNo", SqlDbType.Char).Value = myOldPhoneNo
                            .SelectCommand.Parameters.Add("@OperCode", SqlDbType.Char).Value = myOperCode

                            myData = New DataTable("Response")
                            Using myAdapter
                                .Fill(myData)
                            End Using

                            If myData.Rows.Count > 0 Then
                                myResponse = myData.Rows(0).Item("ReturnMsg").ToString.Trim
                            Else
                                myResponse = "No Response"
                            End If
                            AppendText("ProcessAbacusPhoneNumberChanges Response: " & myResponse)

                            myURIRequest = myURIBase & "<msg>" & myKeyData & "</msg>"

                            'myURIRequest = myURIBASE & myKeyData
                            mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

                            If mySendSuccess = True Then
                                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                                    Dim P3 As Integer = 0
                                    Dim P4 As Integer = 0
                                    Dim myError As String = myURIResponse
                                    While myError.IndexOf("<ERROR>") >= 0
                                        P3 = myError.IndexOf("<ERROR>") + 7
                                        P4 = myError.LastIndexOf("</ERROR>")
                                        myError = myError.Substring(P3, P4 - P3)
                                    End While
                                    AppendText("XML Error: " & myError)
                                    pDS = Nothing
                                Else
                                    Dim myIntelliMsgPSChangeDS As New DataSet()
                                    myIntelliMsgPSChangeDS = ConvertXMLToDataSet(myURIResponse)
                                    If myIntelliMsgPSChangeDS.Tables.Count > 0 Then
                                        myDeleteData = myIntelliMsgPSChangeDS.Tables(0)

                                        If myDeleteData.Rows.Count > 0 Then
                                            myResponse = myDeleteData.Rows(0).Item("DELETE").ToString.Trim
                                        Else
                                            myResponse = "No Response"
                                        End If

                                        If myResponse = "." Then
                                            AppendText("Delete Successful")
                                        Else
                                            AppendText("Delete Not Successful")
                                        End If
                                    End If
                                End If
                            Else
                                pDS = Nothing
                            End If

                            'End If

                            pRecords = pRecords + 1

                        Catch ex As Exception
                            'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                            WriteToErrorEmail(ex.Message.ToString)
                        End Try

                    End With

                Next
            End If
        Catch ex As Exception
            pDS = Nothing
        End Try
    End Sub
    Function ATH(ByVal hex As String) As String
        ATH = ""
        Try
            Dim text As New System.Text.StringBuilder(hex.Length \ 2)
            For i As Integer = 0 To hex.Length - 2 Step 2
                text.Append(Chr(Convert.ToByte(hex.Substring(i, 2), 16)))
            Next
            Return text.ToString
        Catch ex As Exception
        End Try
    End Function
    Public Function HTA(ByVal Data As String) As String
        Return BitConverter.ToString(System.Text.Encoding.ASCII.GetBytes(Data)).Replace("-", "")
    End Function



    Private Sub ReadAppUserPasswordChange()
        Dim myReadCue As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 1

        'Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCUEPC+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""
        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing

        Dim myAppUserId As String = ""
        Dim myDateTimechanged As Date
        Dim mySequenceNumber As Integer = 0
        Dim myOldPassword As String = ""
        Dim myNewPassword As String = ""
        Dim mySSOUUID As String = ""
        Dim myErrorMessage As String = ""
        Dim myProxyError As Boolean = False
        While myContinue = True And myCount < myLimit

            mySql = "exec AppUserPasswordChangeGet"
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myAdapter
                Try
                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using
                    'If myData.Rows.Count < 1 Then Exit Sub 'nothing to process

                    Try
                        If myData.Rows(0).ItemArray.Count <= 1 Then Exit Sub 'NOTHING TO PROCESS
                        ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                        If ReturnMsg = "" Then 'nothing to process
                            Exit Sub
                        End If

                        If ReturnMsg <> "." Then
                            AppendText("AppUserPasswordChangeGet returned: " & ReturnMsg.ToString)
                            WriteToErrorEmail("AppUserPasswordChangeGet returned: " & ReturnMsg.ToString)
                            GoTo NEXTRECORD
                        End If
                    Catch ex As Exception

                        WriteToErrorEmail(ex.Message.ToString)
                        Exit Sub
                    End Try

                    'Keep the key to send back to delete if update to Abacus is completed properly
                    myAppUserId = myData.Rows(0).Item("AppUserId")
                    myDateTimechanged = myData.Rows(0).Item("DateTimeChanged")
                    mySequenceNumber = myData.Rows(0).Item("SequenceNumber")
                    myOldPassword = myData.Rows(0).Item("OldPassword")
                    myNewPassword = myData.Rows(0).Item("NewPassword")
                    mySSOUUID = myData.Rows(0).Item("SSOUUID")

                    AppendText("Processing SSO Change Password: AppUserId: " & myAppUserId & "  DateTimeChanged: " & myDateTimechanged.ToString & _
                                    "  Sequence Number: " & mySequenceNumber)
                    Trace.WriteLine("Processing SSO Change Password: AppUserId: " & myAppUserId & "  DateTimeChanged: " & myDateTimechanged.ToString & _
                                    "  Sequence Number: " & mySequenceNumber)
                    myData = Nothing
                    myAdapter = Nothing
                Catch ex As Exception
                    myData = Nothing
                    myAdapter = Nothing

                    WriteToErrorEmail(ex.Message.ToString)
                End Try
            End With

            'Update SSO
            myProxyError = False
            Trace.WriteLine("AppUser: " & myAppUserId & "  myNewPassword: " & myNewPassword & " myOldPassword: " & myOldPassword)
            Dim myreturn As String = ChangePassword(myAppUserId, myNewPassword, myOldPassword, myAppUserId)

            If myreturn.IndexOf("Proxy Error") > 0 Then 'use later
                myProxyError = True
            End If

            'problem with password change skip update and report.
            If myreturn <> "." Then
                Try
                    WriteToErrorEmail("SSO - Change Password failed for : AppUserId: " & _
                                          myAppUserId & "  DateTimeChanged: " & myDateTimechanged.ToString & _
                                          "  Sequence Number: " & mySequenceNumber & vbCrLf & _
                                          " Returned: " & myreturn)
                    myErrorMessage = "SSO - Change Password failed for : AppUserId: " & _
                                          myAppUserId & "  DateTimeChanged: " & myDateTimechanged.ToString & _
                                          "  Sequence Number: " & mySequenceNumber & vbCrLf & _
                                          " Returned: " & myreturn

                    Trace.WriteLine("SSO - Change Password failed for : AppUserId: " & _
                                         myAppUserId & "  DateTimeChanged: " & myDateTimechanged.ToString & _
                                         "  Sequence Number: " & mySequenceNumber & vbCrLf & _
                                         " Returned: " & myreturn)

                    myErrorMessage = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "|" & myErrorMessage
                    If myErrorMessage.Length > 255 Then 'max len of field
                        myErrorMessage = myErrorMessage.Substring(0, 255)
                    End If

                    mySql = "execute dbo.AppUserPasswordChangeErrorUpdate @AppUserId, @DateTimeChanged, @SequenceNumber, @ErrorMessage"
                    Dim myErrorAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                    With myErrorAdapter
                        .SelectCommand.Parameters.Add("@AppUserId", SqlDbType.VarChar).Value = myAppUserId
                        .SelectCommand.Parameters.Add("@DateTimeChanged", SqlDbType.VarChar).Value = myDateTimechanged.ToString
                        .SelectCommand.Parameters.Add("@SequenceNumber", SqlDbType.Int).Value = mySequenceNumber
                        .SelectCommand.Parameters.Add("@ErrorMessage", SqlDbType.VarChar).Value = myErrorMessage
                        myData = New DataTable("Response")
                        Using myErrorAdapter
                            .Fill(myData)
                        End Using
                        ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                        myData = Nothing
                        myErrorAdapter = Nothing
                        If ReturnMsg <> "." Then
                            WriteToErrorEmail("AppUserPasswordChangeErrorUpdate failed for : AppUserId: " & _
                                                  myAppUserId & "  DateTimeChanged: " & myDateTimechanged.ToString & _
                                                  "  Sequence Number: " & mySequenceNumber & vbCrLf & _
                                                  " Returned: " & myreturn)
                        End If
                    End With

                    'starts getting timeouts, so move on to other processes when this happens
                    If myProxyError Then
                        Exit Sub
                    End If
                Catch ex As Exception
                    WriteToErrorEmail(ex.Message.ToString)
                    GoTo NEXTRECORD
                End Try
                GoTo NEXTRECORD
            End If

            mySql = "execute dbo.AppUserPasswordChangeHistoryUpdate @AppUserId, @DateTimeChanged, @SequenceNumber"
            Dim myDoneAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myDoneAdapter
                Try
                    .SelectCommand.Parameters.Add("@AppUserId", SqlDbType.VarChar).Value = myAppUserId
                    .SelectCommand.Parameters.Add("@DateTimeChanged", SqlDbType.VarChar).Value = myDateTimechanged.ToString
                    .SelectCommand.Parameters.Add("@SequenceNumber", SqlDbType.Int).Value = mySequenceNumber
                    myData = New DataTable("Response")
                    Using myDoneAdapter
                        .Fill(myData)
                    End Using
                    ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                    myData = Nothing
                    myDoneAdapter = Nothing
                    If myreturn <> "." Then
                        WriteToErrorEmail("AppUserPasswordChangeHistoryUpdate failed for : AppUserId: " & _
                                              myAppUserId & "  DateTimeChanged: " & myDateTimechanged.ToString & _
                                              "  Sequence Number: " & mySequenceNumber.ToString & vbCrLf & _
                                                  " Returned: " & ReturnMsg)
                        GoTo NEXTRECORD
                    End If

                Catch ex As Exception
                    myData = Nothing
                    myDoneAdapter = Nothing

                    WriteToErrorEmail(ex.Message.ToString)
                End Try
            End With

            'Process small groups at a time
NEXTRECORD:
            myCount = myCount + 1
            If myCount = myLimit Then Exit Sub 'Getting out of loop, need to delete last transaction
        End While

    End Sub
    Public Function ChangePassword(sAppUserID As String, sNewPassword As String, sCurrentPassword As String, sLoggedInAppUserID As String) As String
        'Dim ctx = HttpContext.Current
        'Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim jsonResponse As String = ""
        Dim retVal As String = ""
        Dim jdo As SSO.SSO_Authenticate
        Dim jdo1 As SSO.SSO_Subject
        Dim retVal1 As String = ""
        Dim retVal2 As String = ""
        Dim sTokenID As String = ""

        Dim sSSOLoginURL As String = UCase(ModMain.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))
        Dim sApiKey As String = ModMain.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
        Dim sSSO_URL As String = ModMain.GetAppControlCharacter("AMS", "MM", "SSO_URL")
        Dim sStatus As String = ""
        Dim sURI As String = ""
        Try
            'sSSO_URL = "https://staging.security.americanmessaging.net"
            'sAppUserID = "reyna@dev.intellimsg.net"
            'sCurrentPassword = "Test1234"

            sURI = sSSO_URL & "/authenticate?apiKey=" & sApiKey & "&loginId=" & sAppUserID & "&password=" & HttpUtility.UrlEncode(sCurrentPassword)
            Trace.WriteLine("ChangePassword: " & " URL:" & sURI)

            Try
                jsonResponse = SSO.GetJSONDownloadString(sURI)
                If IsvalidJson(jsonResponse) = False Then
                    retVal = "ChangePassword: URL: " & sURI & vbCrLf & "Response: " & jsonResponse.ToString
                    SSOAPICallLog("ChangePassword: " & "URL: " & sURI, jsonResponse, False) 'Logging
                    Return retVal
                End If

                jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Authenticate)(jsonResponse)
                sStatus = jdo.status
            Catch ex As Exception
                retVal = "ChangePassword: URL: " & sURI & vbCrLf & "Response: " & jsonResponse.ToString
                SSOAPICallLog("ChangePassword: " & "URL: " & sURI, jsonResponse, False) 'Logging
                Return retVal
            End Try


            If sStatus = "SUCCESS" Then
                sTokenID = jdo.token.tokenId

            Else
                retVal1 = jdo.error.message.ToString 'save old password error message
                'Try new password incase SSO generated the change password and not MM or devices
                sURI = sSSO_URL & "/authenticate?apiKey=" & sApiKey & "&loginId=" & sAppUserID & "&password=" & HttpUtility.UrlEncode(sNewPassword)
                Try
                    jsonResponse = SSO.GetJSONDownloadString(sURI)
                    retVal = jsonResponse 'in case its junk

                    jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Authenticate)(jsonResponse)
                Catch ex As Exception
                    retVal = "ChangePassword: URL: " & sURI & vbCrLf & "Response: " & jsonResponse.ToString
                    SSOAPICallLog("ChangePassword: " & "URL: " & sURI, jsonResponse, False) 'Logging
                    Return retVal
                End Try

                If jdo.status Is Nothing Then 'non json response
                    WriteToErrorEmail("ChangePassword: URL: " & sURI & vbCrLf & "Response: " & jsonResponse.ToString)
                    Return retVal
                End If
                sStatus = jdo.status
                If sStatus = "SUCCESS" Then
                    'already set to new password, nothing left to do
                    'Password change must have been generated by SSO
                    retVal = "."
                    Return retVal
                Else
                    retVal = retVal1 'return orig error message
                    Return retVal
                End If

            End If


            Dim modifyUser As New SSOModifyUser()
            With modifyUser
                .firstName = ""
                .lastName = ""
                .email = ""
                .loginId = sAppUserID
                .password = sCurrentPassword
                .newPassword = sNewPassword
                .active = "true"
            End With


            Dim json As String = JsonConvert.SerializeObject(modifyUser)


            sURI = sSSO_URL & "/subject?apiKey=" & sApiKey & "&tokenId=" & sTokenID

            Dim sysURI As New System.Uri(sURI)
            Dim data = Encoding.UTF8.GetBytes(json)
            jsonResponse = ""

            'API may return junk
            Try
                jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "PUT")
                retVal = jsonResponse 'in case its junk
                jdo1 = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            Catch ex As Exception
                retVal = "ChangePassword: URL: " & sURI & vbCrLf & "Response: " & jsonResponse.ToString
                SSOAPICallLog("ChangePassword: " & "URL: " & sURI, jsonResponse, False) 'Logging
                Return retVal
            End Try

            If jdo1.status Is Nothing Then 'non json response
                WriteToErrorEmail("SSO - jsonResponse: " & jsonResponse)
                Exit Try
            End If
            sStatus = jdo1.status

            If sStatus = "SUCCESS" Then
                retVal = "."
            Else
                If jdo1.error.message IsNot Nothing Then 'sometimes the server return proxy web error instead of formatted JSON
                    retVal = jdo1.error.message
                Else
                    retVal = jsonResponse
                End If


            End If
        Catch ex As Exception
            If jsonResponse <> "" Then
                WriteToErrorEmail("SSO - jsonResponse: " & jsonResponse & vbCrLf & ex.Message.ToString)
            Else
                WriteToErrorEmail("SSO -" & ex.Message.ToString)
            End If
        End Try

        Return retVal

    End Function



    Private Sub ReadWIFCLFWD()
        'call forward transactions
        Dim myReadCue As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 10

        'Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCUEPC+"
        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMCLFWD+"

        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Dim myDateTimeReceived As Date
        Dim myDateTimeReceivedStr As String = ""
        Dim myAppUserId As String = ""
        Dim myMessageNumber As String = ""
        Dim myPagerNumber As String = ""
        Dim myMessageReceivedText As String = ""
        Dim myCallForwardTransType As String = ""
        Dim myCallForwardPhoneNumber As String = ""
        Dim myProcessErrorMessage As String = ""

        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing


        'READ ================
        While myContinue = True And myCount < myLimit

            mySql = "execute dbo.WIFCLFWDGet "
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myAdapter
                Try

                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using

                    Try
                        If myData.Rows(0).ItemArray.Count <= 1 Then Exit Sub 'NOTHING TO PROCESS

                    Catch ex As Exception
                        WriteToErrorEmail(ex.Message.ToString)
                        GoTo NEXTLOOP
                    End Try

                    'Keep the key to send back to delete if update to Abacus is completed properly
                    ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                    If ReturnMsg = "norecords" Then
                        GoTo NEXTLOOP
                    End If

                    If ReturnMsg <> "." Then
                        WriteToErrorEmail("Error Stored Procedure WIFCLFWDGet " & ReturnMsg)
                        GoTo NEXTLOOP
                    End If

                    myDateTimeReceived = myData.Rows(0).Item("DateTimeReceived")
                    myAppUserId = myData.Rows(0).Item("AppUserId")
                    myMessageNumber = myData.Rows(0).Item("MessageNumber")
                    myPagerNumber = myData.Rows(0).Item("PagerNumber")
                    myMessageReceivedText = myData.Rows(0).Item("MessageReceivedText")
                    myCallForwardTransType = myData.Rows(0).Item("CallForwardTransType")
                    myCallForwardPhoneNumber = myData.Rows(0).Item("CallForwardPhoneNumber")

                    'format date to string 2013-11-07 10:11:16.453
                    myDateTimeReceivedStr = myDateTimeReceived.ToString("yyyyMMddHHmmss")

                    myData = Nothing
                    myAdapter = Nothing
                Catch ex As Exception
                    myData = Nothing
                    myAdapter = Nothing
                    WriteToErrorEmail(ex.Message.ToString)
                    GoTo NEXTLOOP
                End Try
            End With

            'PROCESS =============
            myURIRequest = myURIBase & "<MSG>" & HTA(myDateTimeReceivedStr) & "/" _
                                                & HTA(myAppUserId) & "/" _
                                                & HTA(myMessageNumber) & "/" _
                                                & HTA(myPagerNumber) & "/" _
                                                & HTA(myMessageReceivedText) & "/" _
                                                & HTA(myCallForwardTransType) & "/" _
                                                & HTA(myCallForwardPhoneNumber) & "</MSG>"

            ReturnMsg = ""
            myReadCue = SendHTTPRequest(myURIRequest, ReturnMsg)

            If myReadCue = True Then

                ReturnMsg = ReturnMsg.Trim.ToString
                If ReturnMsg <> "." Then
                    myProcessErrorMessage = ReturnMsg
                    WriteToErrorEmail("WIFCLFWDUpdate process failed." & vbCrLf & _
                                        "Return Message: " & ReturnMsg & vbCrLf & _
                                        "DateTimeReceived: " & myDateTimeReceivedStr & vbCrLf & _
                                         "AppUserId: " & myAppUserId & vbCrLf & _
                                         "PagerNumber: " & myPagerNumber & vbCrLf & _
                                         "MessageReceivedText: " & myMessageReceivedText & vbCrLf & _
                                         "CallForwardTransType: " & myCallForwardTransType & vbCrLf & _
                                         "CallForwardPhoneNumber: " & myCallForwardPhoneNumber)
                End If

                'UPDATE ====================

                If myProcessErrorMessage.Trim.Length > 50 Then 'table max length is 50 char
                    myProcessErrorMessage = myProcessErrorMessage.Substring(0, 50)
                End If
                mySql = "execute dbo.WIFCLFWDUpdate @DateTimeReceived, @AppUserId, @MessageNumber, @ProcessErrorMessage"
                Dim myDoneAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                With myDoneAdapter
                    Try
                        .SelectCommand.Parameters.Add("@DateTimeReceived", SqlDbType.DateTime).Value = myDateTimeReceived
                        .SelectCommand.Parameters.Add("@AppUserId", SqlDbType.Char).Value = myAppUserId
                        .SelectCommand.Parameters.Add("@MessageNumber", SqlDbType.Char).Value = myMessageNumber.ToString
                        .SelectCommand.Parameters.Add("@ProcessErrorMessage", SqlDbType.Char).Value = myProcessErrorMessage.ToString


                        myData = New DataTable("Response")
                        Using myDoneAdapter
                            .Fill(myData)
                        End Using

                        'verify update successful check if data row exists
                        Try
                            ReturnMsg = ""
                            ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                        Catch ex As Exception
                            WriteToErrorEmail("WIFCLFWDUpdate process failed,  No record returned for " & vbCrLf & _
                                       "DateTimeReceived: " & myDateTimeReceivedStr & vbCrLf & _
                                       "AppUserId: " & myAppUserId & vbCrLf & _
                                       "MessageNumber: " & myMessageNumber & vbCrLf & _
                                       "ProcessErrorMessage: " & myProcessErrorMessage)
                            GoTo NEXTLOOP
                        End Try

                        If ReturnMsg <> "." Then
                            WriteToErrorEmail("WIFCLFWDUpdate process failed." & vbCrLf & _
                                           "DateTimeReceived: " & myDateTimeReceivedStr & vbCrLf & _
                                           "AppUserId: " & myAppUserId & vbCrLf & _
                                           "MessageNumber: " & myMessageNumber & vbCrLf & _
                                           "ProcessErrorMessage: " & myProcessErrorMessage & vbCrLf & _
                                           "Return Message: " & ReturnMsg.ToString)

                        End If

                        myData = Nothing
                        myDoneAdapter = Nothing

                    Catch ex As Exception
                        myData = Nothing
                        myDoneAdapter = Nothing

                        WriteToErrorEmail(ex.Message.ToString)
                        GoTo NEXTLOOP
                    End Try
                End With
            Else 'myReadCue = FALSE
                'myURIRequest
                WriteToErrorEmail("WIFCLFWDUpdate SendHTTPRequest process failed." & vbCrLf & _
                               "DateTimeReceived: " & myDateTimeReceivedStr & vbCrLf & _
                               "AppUserId: " & myAppUserId & vbCrLf & _
                               "MessageNumber: " & myMessageNumber & vbCrLf & _
                               "ProcessErrorMessage: " & myProcessErrorMessage & vbCrLf & _
                               "uriRequest: " & myURIRequest.ToString)

            End If 'If myReadCue = True Then END

            'Continue Processing
NEXTLOOP:
            myCount = myCount + 1
            If myCount = myLimit Then Exit While 'Getting out of loop, need to delete last transaction

        End While ' looping

    End Sub

    Private Sub ReadWIFMCDOO()
        Dim myReadPager As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 10

        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMMCDOO+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Dim YYYYMMDDHHMISS As String = ""
        Dim DATA_ACTION As String = ""
        Dim CO_CODE As String = ""
        Dim ACCT_NO As String = ""
        Dim SERVICE_TYPE As String = ""
        Dim MC_DOMAIN As String = ""
        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing
        Dim YMDWork As String = ""

        While myContinue = True And myCount < myLimit

            myURIRequest = myURIBase & "<MSG>" & YYYYMMDDHHMISS & "/" & DATA_ACTION & "/" & CO_CODE & "/" & ACCT_NO & "/" & SERVICE_TYPE & "</MSG>"

            myReadPager = SendHTTPRequest(myURIRequest, myURIResponse)

            If myReadPager = True Then
                If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                    Dim P1 As Integer = 0
                    Dim P2 As Integer = 0
                    Dim myError As String = myURIResponse
                    While myError.IndexOf("<ERROR>") >= 0
                        P1 = myError.IndexOf("<ERROR>") + 7
                        P2 = myError.LastIndexOf("</ERROR>")
                        myError = myError.Substring(P1, P2 - P1)
                    End While
                    AppendText("XML Error: " & myError)
                    myContinue = False
                Else
                    Dim dstXMLData As New DataSet()
                    dstXMLData = ConvertXMLToDataSet(myURIResponse)
                    YYYYMMDDHHMISS = dstXMLData.Tables(0).Rows(0).Item("YYYYMMDDHHMISS")
                    DATA_ACTION = dstXMLData.Tables(0).Rows(0).Item("DATA_ACTION")
                    CO_CODE = dstXMLData.Tables(0).Rows(0).Item("CO_CODE")
                    ACCT_NO = dstXMLData.Tables(0).Rows(0).Item("ACCT_NO")
                    SERVICE_TYPE = dstXMLData.Tables(0).Rows(0).Item("SERVICE_TYPE")
                    MC_DOMAIN = dstXMLData.Tables(0).Rows(0).Item("MC_DOMAIN")

                    YMDWork = ATH(YYYYMMDDHHMISS)
                    If YMDWork.Trim = "" Then
                        myContinue = False
                    Else
                        mySql = "execute dbo.SaveAbacusWIFMCDOO @YYYYMMDDHHMISS, @DATA_ACTION, @CO_CODE, @ACCT_NO, @SERVICE_TYPE, @MC_DOMAIN"
                        Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
                        With myAdapter
                            Try
                                .SelectCommand.Parameters.Add("@YYYYMMDDHHMISS", SqlDbType.Char).Value = ATH(YYYYMMDDHHMISS)
                                .SelectCommand.Parameters.Add("@DATA_ACTION", SqlDbType.Char).Value = ATH(DATA_ACTION)
                                .SelectCommand.Parameters.Add("@CO_CODE", SqlDbType.Char).Value = ATH(CO_CODE)
                                .SelectCommand.Parameters.Add("@ACCT_NO", SqlDbType.Char).Value = ATH(ACCT_NO)
                                .SelectCommand.Parameters.Add("@SERVICE_TYPE", SqlDbType.Char).Value = ATH(SERVICE_TYPE)
                                .SelectCommand.Parameters.Add("@MC_DOMAIN", SqlDbType.Char).Value = ATH(MC_DOMAIN)

                                myData = New DataTable("Response")
                                Using myAdapter
                                    .Fill(myData)
                                End Using
                                ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                                If ReturnMsg <> "." Then
                                    AppendText("WIFMCDOO " & ReturnMsg)
                                    WriteToErrorEmail("WIFMCDOO " & ReturnMsg)
                                    myContinue = False
                                End If

                                myData = Nothing
                                myAdapter = Nothing
                            Catch ex As Exception

                                WriteToErrorEmail(ex.Message.ToString)
                                'MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
                            End Try
                        End With
                        myCount = myCount + 1
                        If myCount = myLimit Then 'Getting out of loop, need to delete last transaction
                            myURIRequest = myURIBase & "<MSG>" & YYYYMMDDHHMISS & "/" & DATA_ACTION & "/" & CO_CODE & "/" & ACCT_NO & "/" & SERVICE_TYPE & "</MSG>"
                            myReadPager = SendHTTPRequest(myURIRequest, myURIResponse)
                        End If
                    End If
                End If
            Else
                myContinue = False
            End If

        End While

        If myCount > 0 Then
            myActivity = True
        End If

    End Sub

    Private Sub ReadWIFIMSGUCreated()
        Dim TransactionDate As String = Date.Now.ToString("yyyyMMdd")
        Dim CO_CODE As String = "H1"
        Dim ACCT_NO As String = "767248"
        Dim PROCESS_ID As String = "CreateIntelliMessageWIFIMSGU"
        Dim RUN_YYYYMMDD As String = ""

        Dim mySql As String = ""
        Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing
        RUN_YYYYMMDD = TransactionDate

        mySql = "execute dbo.CheckWIFIMSGUCreated @CO_CODE, @ACCT_NO, @PROCESS_ID, @RUN_YYYYMMDD"
        Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
        With myAdapter
            Try
                .SelectCommand.Parameters.Add("@CO_CODE", SqlDbType.VarChar).Value = CO_CODE
                .SelectCommand.Parameters.Add("@ACCT_NO", SqlDbType.Char).Value = ACCT_NO
                .SelectCommand.Parameters.Add("@PROCESS_ID", SqlDbType.Char).Value = PROCESS_ID
                .SelectCommand.Parameters.Add("@RUN_YYYYMMDD", SqlDbType.Char).Value = RUN_YYYYMMDD

                myData = New DataTable("Response")
                Using myAdapter
                    .Fill(myData)
                End Using
                ReturnMsg = myData.Rows(0).Item("ReturnMsg")

                If ReturnMsg <> "." Then
                    AppendText("ReadWIFIMSGUCreated " & ReturnMsg)
                    gIntelliMessageWIFIMSGUCreated = False
                Else
                    gIntelliMessageWIFIMSGUCreated = True
                End If

                myData = Nothing
                myAdapter = Nothing
            Catch ex As Exception
                ' MsgBox("The following unexpected error has occurred: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)

                WriteToErrorEmail(ex.Message.ToString)
            End Try
        End With
    End Sub

    Private Sub ReadWIFIMSGUCleared()
        Dim TransactionDate As String = Date.Now.ToString("yyyyMMdd")
        Dim myReadWIFIMSGU As Boolean = False
        Dim myURIData As String = ""
        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMIMSGU+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""
        Dim CO_CODE As String = "H1"
        Dim ACCT_NO As String = "767248"
        Dim PROCESS_ID As String = "AAMIMSGU"
        Dim RUN_YYYYMMDD As String = ""
        Dim RUN_HHMISS As String = ""
        Dim READ_OR_UPDATE As String = "R"
        RUN_YYYYMMDD = TransactionDate
        gRUN_HHMISS = ""

        myURIRequest = myURIBase & "<MSG>" & CO_CODE & "/" & ACCT_NO & "/" & PROCESS_ID & "/" & RUN_YYYYMMDD & "/" & RUN_HHMISS & "/" & READ_OR_UPDATE & "</MSG>"

        myReadWIFIMSGU = SendHTTPRequest(myURIRequest, myURIResponse)

        If myReadWIFIMSGU = True Then
            If myURIResponse.IndexOf("input parameters") >= 0 Or myURIResponse.IndexOf("NotComplete") >= 0 Then
                AppendText("WIMIMSGU: " & myURIResponse)
            Else
                Dim dstXMLData As New DataSet()
                dstXMLData = ConvertXMLToDataSet(myURIResponse)
                RUN_HHMISS = dstXMLData.Tables(0).Rows(0).Item("RUN_HHMISS")

                If RUN_HHMISS.Trim <> "" Then
                    gAbacusWIFIMSGUCleared = True
                    gRUN_HHMISS = RUN_HHMISS
                    gRUN_YYYYMMDD = RUN_YYYYMMDD
                End If
            End If
        Else
            AppendText("WIMIMSGU did not execute properly")
            gAbacusWIFIMSGUCleared = False
        End If
    End Sub

    Private Sub ReadWIFIMSGUToCopy(ByRef pWIFIMSGU As DataSet)
        Dim mySql As String = ""
        Dim myFlag As Boolean = False

        Try
            mySql = "exec GetWIFIMSGUToCopy"
            myFlag = FillSqlDataSet(mySql, pWIFIMSGU, gErrorMessage)
            If pWIFIMSGU.Tables(0).Rows.Count = 0 Then
                myFlag = False
            End If
            myIMSGUsers = myFlag
        Catch ex As Exception
            WriteToErrorEmail(ex.Message.ToString)
        End Try
    End Sub

    Private Sub SendWIFIMSGUToAdd(ByRef pWIFIMSGU As DataSet, ByRef pRecords As Integer)
        Dim myDataTable As DataTable = Nothing
        Dim myWIFIMSGU As DataTable = Nothing
        Dim myWIFIMSGUDTRow As DataRow = Nothing
        Dim myWIFIMSGUDS As DataSet = Nothing
        Dim DESTINATION As String = ""
        Dim PAGER_PHONE_NO As String = ""
        Dim myWIFIMSGUXMLData As String = ""
        Dim mySQL As String = ""
        Dim myFlag As Boolean = False
        Dim myData As DataTable = Nothing
        Dim myResponse As String = ""
        Dim P1 As Integer = 0
        Dim p2 As Integer = 0
        Dim mySendSuccess As Boolean = False


        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIUIMSGU+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""

        Try
            If pWIFIMSGU.Tables.Count > 0 Then
                myDataTable = pWIFIMSGU.Tables(0)
                myWIFIMSGU = New DataTable("WIFIMSGU")
                myWIFIMSGU.Columns.Add("DESTINATION", Type.GetType("System.String"))
                myWIFIMSGU.Columns.Add("PAGER_PHONE_NO", Type.GetType("System.String"))
                For Each myRow As DataRow In myDataTable.Rows
                    DESTINATION = myRow.Item("DESTINATION").ToString
                    PAGER_PHONE_NO = myRow.Item("PAGER_PHONE_NO").ToString
                    myWIFIMSGUDTRow = myWIFIMSGU.NewRow
                    myWIFIMSGUDTRow.Item("DESTINATION") = HTA(DESTINATION)
                    myWIFIMSGUDTRow.Item("PAGER_PHONE_NO") = HTA(PAGER_PHONE_NO)
                    myWIFIMSGU.Rows.Add(myWIFIMSGUDTRow)
                Next

                myWIFIMSGUDS = New DataSet("WIFIMSGU")
                myWIFIMSGUDS.Tables.Add(myWIFIMSGU)

                myWIFIMSGUXMLData = ConvertDataSetToXML(myWIFIMSGUDS)

                myURIRequest = myURIBase & "<msg>" & myWIFIMSGUXMLData & "</msg>"

                mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)

                If mySendSuccess = True Then
                    If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                        Dim P3 As Integer = 0
                        Dim P4 As Integer = 0
                        Dim myError As String = myURIResponse
                        While myError.IndexOf("<ERROR>") >= 0
                            P3 = myError.IndexOf("<ERROR>") + 7
                            P4 = myError.LastIndexOf("</ERROR>")
                            myError = myError.Substring(P3, P4 - P3)
                        End While
                        AppendText("XML Error: " & myError)
                        pWIFIMSGU = Nothing
                    Else
                        If myURIResponse = "." Then
                            AppendText("Add Successful")

                            mySQL = "exec DeleteWIFIMSGUCopied"
                            Dim myAdapter As New SqlClient.SqlDataAdapter(mySQL, gDatabaseConnectionString)
                            With myAdapter
                                Try
                                    myData = New DataTable("Response")
                                    Using myAdapter
                                        .Fill(myData)
                                    End Using

                                    If myData.Rows.Count > 0 Then
                                        myResponse = myData.Rows(0).Item("ReturnMsg").ToString.Trim
                                    Else
                                        myResponse = "No Response"
                                    End If
                                    AppendText("HarkNumber Response: " & myResponse)

                                    If myResponse = "." Then
                                        AppendText("DeleteWIFIMSGUCopied Successful")
                                    Else
                                        AppendText("DeleteWIFIMSGUCopied Not Successful")
                                    End If
                                Catch ex As Exception
                                    WriteToErrorEmail(ex.Message.ToString)
                                End Try
                            End With
                        End If
                    End If
                Else
                    pWIFIMSGU = Nothing
                End If

                pRecords = pRecords + 1
            End If
        Catch ex As Exception
            pWIFIMSGU = Nothing
        End Try
    End Sub

    Private Sub UpdateWIFIMSGUCreated()
        Dim TransactionDate As String = Date.Now.ToString("yyyyMMdd")
        Dim myReadWIFIMSGU As Boolean = False
        Dim myURIData As String = ""
        Dim myURIBase As String = gURIBase & "xmlrequest=METHOD+WIMIMSGU+"
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""
        Dim CO_CODE As String = "H1"
        Dim ACCT_NO As String = "767248"
        Dim PROCESS_ID As String = "AAMIMSGU"
        Dim RUN_YYYYMMDD As String = gRUN_YYYYMMDD
        Dim RUN_HHMISS As String = gRUN_HHMISS
        Dim READ_OR_UPDATE As String = "U"
        RUN_YYYYMMDD = TransactionDate

        myURIRequest = myURIBase & "<MSG>" & CO_CODE & "/" & ACCT_NO & "/" & PROCESS_ID & "/" & RUN_YYYYMMDD & "/" & RUN_HHMISS & "/" & READ_OR_UPDATE & "</MSG>"

        myReadWIFIMSGU = SendHTTPRequest(myURIRequest, myURIResponse)

        If myReadWIFIMSGU = True Then
            If myURIResponse.IndexOf("<ERROR>") >= 0 Then
                Dim P1 As Integer = 0
                Dim P2 As Integer = 0
                Dim myError As String = myURIResponse
                While myError.IndexOf("<ERROR>") >= 0
                    P1 = myError.IndexOf("<ERROR>") + 7
                    P2 = myError.LastIndexOf("</ERROR>")
                    myError = myError.Substring(P1, P2 - P1)
                End While
                AppendText("XML Error: " & myError)
            Else
                Dim dstXMLData As New DataSet()
                dstXMLData = ConvertXMLToDataSet(myURIResponse)
                RUN_HHMISS = dstXMLData.Tables(0).Rows(0).Item("RUN_HHMISS")

                If RUN_HHMISS.Trim <> "" Then
                    AppendText("Abacus WIFIMSGU Create was Successful")
                    gIMSGUsersLastUpdate = Now
                    gAbacusWIFIMSGUCleared = False
                    gIntelliMessageWIFIMSGUCreated = False
                End If
            End If
        End If
    End Sub



    Private Sub Create_Update_SSO_Accounts()

        'Process(outline) ================================================
        'NEW API REQUEST SSO, AM, AND MM CREATE,UPDATE, AND SYNC

        '	Create SSO IM user from MM AppUser ID
        '	Create SSO AM if does not exist
        '	Link SSO IM and SSO AM accounts

        'APP USER ACCOUNT RENAME
        '	When APP User account is renamed, MM process will update table AppUser  field OldAppUserID with the Prior Account ID
        '	After creating new SSO IM account will need to set old SSO IM Account Inactive with API Modify Subject, field Active  = false
        '	After creating new SSO AM account, check the App User field Primary Email for the new and old accounts.  Create the new account. 
        '   If the Primary Phone number exists on a different active account in table App User then do not deactivate the old account.

        'MM ACCOUNT PASSWORD CHANGE
        '	Change the password for both SSO IM account if Account and SSO AM accounts.
        '===================================================================


        Dim myReadCue As Boolean = False
        Dim myURIData As String = ""
        Dim myContinue As Boolean = True
        Dim myCount As Integer = 0
        Dim myLimit As Integer = 1
        Dim myURIRequest As String = ""
        Dim myURIResponse As String = ""
        Dim mySSOUUID As String = ""
        Dim mySql As String = ""
        'Dim ReturnMsg As String = ""
        Dim myData As DataTable = Nothing

        Dim myAppUserId As String = ""
        Dim myOldAppUserId As String = ""
        Dim myOldPrimaryEmail As String = ""
        Dim mySequenceNumber As Integer = 0
        Dim myPassword As String = ""
        Dim mySSO_IM_UUID As String = ""
        Dim mySSO_AM_UUID As String = ""
        Dim mySSO_IM_Active As String = ""
        Dim mySSO_AM_Active As String = ""
        Dim mySSO_AM_Token As String = ""
        Dim mySSO_IM_Token As String = ""
        Dim myErrorMessage As String = ""
        Dim myPrimaryEmail As String = ""
        Dim myFirstName As String = ""
        Dim myLastName As String = ""
        Dim myCall As String = ""
        Dim myResult As String = ""
        Dim myPrimaryPhone As String = ""
        Dim myMessage As String = ""

        Dim myAMAdministrator As String = ""
        Dim myreturn As String = ""
        Dim myAMUUID As String = ""

        Dim myAdminToken As String = ""
        Dim myAcctStatus As String = "" 'valid values "okay", "invalidpassword", "invalidaccount", "accountinactive"



        Try
            'gApiKey = GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            gSSO_IM_APIKey = GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            gSSO_AM_APIKey = GetAppControlCharacter("AMS", "MM", "SSO_AMApiKey")


            gSSO_URL = GetAppControlCharacter("AMS", "MM", "SSO_URL")

            If gSSO_IM_APIKey.Trim = "" Then
                myMessage = "SSO_ApiKey control record not found"
                AppendText(myMessage)
                WriteToErrorEmail(myMessage)
                Exit Sub
            End If
            If gSSO_AM_APIKey.Trim = "" Then
                myMessage = "SSO_AMApiKey control record not found"
                AppendText(myMessage)
                WriteToErrorEmail(myMessage)
                Exit Sub
            End If


            If gSSO_URL.Trim = "" Then
                myMessage = "gSSO_URL control record not found"
                AppendText(myMessage)
                WriteToErrorEmail(myMessage)
            End If
        Catch ex As Exception
            myMessage = "Sub: Create_Update_SSO_Accounts, Get Control Records Failed " & ex.Message.ToString
            AppendText(myMessage)
            WriteToErrorEmail(myMessage)
        End Try
        Trace.WriteLine("Running " & DateTime.Now.ToString)
        While myContinue = True And myCount < myLimit
            'Get account ===================================================================================
            'Clear fields
            mySSO_IM_UUID = "" : mySSO_IM_Active = "" : mySSO_AM_UUID = "" : mySSO_AM_Active = ""

            gSSOInterfaceFailing = ""
            mySql = "exec AppUserSSOUserCreate"
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myAdapter
                Try
                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using
                    'If myData.Rows.Count < 1 Then Exit Sub 'nothing to process


                    Try
                        If myData.Rows(0).ItemArray.Count <= 1 Then Exit Sub 'NOTHING TO PROCESS
                        myreturn = myData.Rows(0).Item("ReturnMsg")
                        myAppUserId = myData.Rows(0).Item("AppUserId")
                        myPrimaryEmail = myData.Rows(0).Item("PrimaryEmail")
                        myFirstName = myData.Rows(0).Item("FirstName")
                        myLastName = myData.Rows(0).Item("LastName")
                        myPassword = myData.Rows(0).Item("Password")
                        myPrimaryPhone = myData.Rows(0).Item("PrimaryPhone")
                        myAMAdministrator = myData.Rows(0).Item("AMAdministrator")
                        mySSOUUID = myData.Rows(0).Item("SSOUUID")
                        myOldAppUserId = myData.Rows(0).Item("OldAppUserId")
                        myOldPrimaryEmail = myData.Rows(0).Item("OldPrimaryEmail")
                        'SSO API wants everything in lower case
                        myAppUserId = myAppUserId.ToLower
                        myPrimaryEmail = myPrimaryEmail.ToLower.Trim

                        'Spaces in SSO API will cause it to explode
                        If myPrimaryEmail.IndexOf(" ") > 0 Then
                            myMessage = "Sub: Create_Update_SSO_Accounts, Bad Primary Email Address for account " & myAppUserId & _
                                vbCrLf & " PrimaryEmail: " & myPrimaryEmail
                            AppendText(myMessage)
                            WriteToErrorEmail(myMessage)
                            GoTo NEXTRECORD

                        End If

                        'TESTING ============================================================
                        'myAMAdministrator = "1"
                        'myAppUserId = "dcb20140818@dev.intellimsg.net"
                        'myPassword = "dcb20140818"
                        'myPrimaryEmail = "dcb20140818@abw.com"
                        'myreturn = "."
                        
                        'myreturn = "."
                        'myreturn = "" : myreturn = AuthenticateUser(myAppUserId, myPassword, mySSO_IM_UUID, mySSO_IM_Token, myAcctStatus) 'AuthenticateUser returns Token
                        'myAdminToken = GetAdminToken()
                        ''myreturn = "" : myreturn = AuthenticateUser(myAppUserId, myPassword, mySSO_IM_UUID, mySSO_IM_Token, myAcctStatus) 'AuthenticateUser returns UUID, Active Flag,Token
                        'myreturn = ModifySSOUser(myAdminToken, "", "", "", myAppUserId, "", "", "0")
                        'myreturn = ModifySSOUser(myAdminToken, "", "", "", myAppUserId, "", "", "1")
                        '====================================================================
                        'Check if there are Open Records
                        If myreturn = "" Then GoTo NEXTRECORD 'Nope
                        'Check for Error
                        If myreturn <> "." Then
                            myMessage = "Sub: Create_Update_SSO_Accounts, SP: AppUserSSOUserCreate Returned: " & myreturn.ToString
                            AppendText(myMessage)
                            WriteToErrorEmail(myMessage)
                            GoTo NEXTRECORD
                        End If
                    Catch ex As Exception
                        WriteToErrorEmail("Sub: Create_Update_SSO_Accounts, SP: AppUserSSOUserCreate Returned: " + ex.Message.ToString)
                        Exit Sub
                    End Try
                    myData = Nothing
                    myAdapter = Nothing
                Catch ex As Exception
                    myData = Nothing
                    myAdapter = Nothing
                    WriteToErrorEmail("Sub: Create_Update_SSO_Accounts, SP: AppUserSSOUserCreate Returned: " + ex.Message.ToString)
                End Try
            End With
            Trace.WriteLine("CURRENT USER: " & myAppUserId)
            'Get INTELLIMSG_ADMIN token
            If myAdminToken = "" Then
                myAdminToken = GetAdminToken()
            End If
            If myAdminToken = "" Then
                myMessage = "Sub: Create_Update_SSO_Accounts, Process GetAdminToken Failed: IM AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                gSSOInterfaceFailing = myMessage
                AppendText(myMessage)
                WriteToErrorEmail(myMessage)
                GoTo PROCESSFAILED
            End If

            'Create SSO IM User account if it does not exist =======================================================           
            myreturn = "" : myreturn = AuthenticateUser(gSSO_IM_APIKey, myAppUserId, myPassword, mySSO_IM_UUID, mySSO_IM_Token, myAcctStatus) 'AuthenticateUser returns UUID, Active Flag,Token
            If myAcctStatus.ToLower = "invalidpassword" Then 'invalid password
                myMessage = "Sub: Create_Update_SSO_Accounts, Process CreateSSOUser Failed MISSMATCHED PASSWORDS: SSO_IM AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                gSSOInterfaceFailing = myreturn
                AppendText(myMessage)
                WriteToErrorEmail(myMessage)
                GoTo PROCESSFAILED
            End If

            If myreturn <> "." And myAcctStatus = "invalidaccount" Then 'failed to validate account
                'Create SSO User 
                myreturn = "" : myreturn = CreateSSOUser(gSSO_IM_APIKey, myFirstName, myLastName, myPrimaryEmail, myAppUserId, myPassword, myPrimaryPhone, mySSO_IM_UUID) 'returns UUID
                If myreturn <> "." Then 'success
                    gSSOInterfaceFailing = myreturn
                    AppendText(myreturn)
                    WriteToErrorEmail(myreturn)
                    GoTo PROCESSFAILED
                End If

                'get Token needed for Link Accounts Add Principle  Process after adding account
                myreturn = "" : myreturn = AuthenticateUser(gSSO_IM_APIKey, myAppUserId, myPassword, mySSO_IM_UUID, mySSO_IM_Token, myAcctStatus) 'returns Token
                If myreturn <> "." Then 'success
                    gSSOInterfaceFailing = myreturn
                    myMessage = "Sub: Create_Update_SSO_Accounts, Process CreateSSOUser/AuthenticateUser After CreateSSOUser : SSO_IM AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                    AppendText(myMessage)
                    WriteToErrorEmail(myMessage)
                    GoTo PROCESSFAILED
                End If

            Else  'Make SSO account active 
                If myAcctStatus = "accountinactive" Then
                    myreturn = ModifySSOUser(gSSO_IM_APIKey, myAdminToken, "", "", "", myAppUserId, "", "", "1")
                    If myreturn <> "." Then 'success
                        gSSOInterfaceFailing = myreturn
                        myMessage = "Sub: Create_Update_SSO_Accounts, Process ModifySSOUser Failed: SSO AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                        AppendText(myMessage)
                        WriteToErrorEmail(myMessage)
                        GoTo PROCESSFAILED
                    End If
                End If
            End If

            'SSO_IM Update MM SSOUUID field if blank or if the 2 UUID don't match ====================================================================
            If (mySSOUUID = "" And mySSO_IM_UUID.Trim <> "") Or _
                (mySSOUUID <> "" And mySSO_IM_UUID.Trim <> "" And mySSO_IM_UUID.Trim <> mySSOUUID) Then 'From appUser table

                myreturn = SaveAppUserSSOUUID(myAppUserId, mySSO_IM_UUID, "Y")
                mySSOUUID = mySSO_IM_UUID
                If myreturn <> "." Then
                    myMessage = "Sub: Create_Update_SSO_Accounts, Process SaveAppUserSSOUUID Failed: AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                    gSSOInterfaceFailing = myMessage
                    AppendText(myMessage)
                    WriteToErrorEmail(myMessage)
                    GoTo PROCESSFAILED
                End If
            End If

            'Create SSO AM Account if not Exists========================================================================================================================
            'Account exists and/or password is different.  Still try add principle
            If myAMAdministrator = True And myPrimaryEmail.Trim <> "" Then 'Does AM Account already exist
                myreturn = "" : myreturn = AuthenticateUser(gSSO_IM_APIKey, myPrimaryEmail, myPassword, mySSO_AM_UUID, mySSO_AM_Token, myAcctStatus) 'returns UUID, Active, Token
                If myAcctStatus = "invalidpassword" Then 'Account exists with different password. 
                    'Go ahead and add Principle
                    GoTo ADD_PRINCIPLE
                End If

                'Try and Create SSO_AM account
                If myreturn <> "." And myAcctStatus = "invalidaccount" Then 'not exist 
                    myreturn = "" : myreturn = CreateSSOUser(gSSO_AM_APIKey, myFirstName, myLastName, myPrimaryEmail, myPrimaryEmail, myPassword, myPrimaryPhone, mySSO_AM_UUID) 'returns UUID
                    If myreturn <> "." Then 'success
                        gSSOInterfaceFailing = myMessage
                        AppendText(myreturn)
                        WriteToErrorEmail(myreturn)
                        GoTo PROCESSFAILED
                    End If
                Else 'SET SSO AM account active incase
                    'This process may not be needed because the SSO_AM account should be linked to the SSO_IM account LEAVE FOR NOW
                    If myAcctStatus = "accountinactive" Then
                        myreturn = ModifySSOUser(gSSO_AM_APIKey, myAdminToken, "", "", "", myPrimaryEmail, "", "", "1")
                        If myreturn <> "." Then 'success
                            gSSOInterfaceFailing = myreturn
                            myMessage = "Sub: Create_Update_SSO_AM_Accounts, Process ModifySSOUser Failed: SSO AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                            AppendText(myMessage)
                            WriteToErrorEmail(myMessage)
                            GoTo PROCESSFAILED
                        End If
                    End If
                End If
            End If


            'Try and add priciple anyway
ADD_PRINCIPLE:

            'Link Accounts Add Principle =======================================================================================================
            If myAMAdministrator = True Then


                'removed 8/18/2014 think it is wrong
                'If mySSO_IM_UUID = "" Then
                '    myMessage = "Sub: Create_Update_SSO_Accounts, Process AddPrincipal Failed: SSO AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                '    gSSOInterfaceFailing = myMessage
                '    AppendText(myMessage)
                '    WriteToErrorEmail(myMessage)
                '    GoTo PROCESSFAILED
                'End If
                'check if Accounts are not already linked
                'Link Accounts
                myreturn = "" : myreturn = AddPrincipal(myAdminToken, myPrimaryEmail, mySSO_IM_UUID)
                If myreturn <> "." Then 'failure
                    myMessage = "Sub: Create_Update_SSO_Accounts, Process AddPrincipal Failed: IM AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                    gSSOInterfaceFailing = myMessage
                    AppendText(myMessage)
                    WriteToErrorEmail(myMessage)
                    GoTo PROCESSFAILED
                End If

                'BUG API SSO_IM UUID is changed, Need to get new UUID and correct MM
                myreturn = "" : myreturn = AuthenticateUser(gSSO_IM_APIKey, myAppUserId, myPassword, mySSO_IM_UUID, mySSO_IM_Token, myAcctStatus) 'AuthenticateUser returns UUID, Active Flag,Token
                If myreturn <> "." Then 'failure
                    myMessage = "Sub: Create_Update_SSO_Accounts, Process AuthenticateUser After Link Accts; IM AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                    gSSOInterfaceFailing = myMessage
                    AppendText(myMessage)
                    WriteToErrorEmail(myMessage)
                    GoTo PROCESSFAILED
                End If

                'Update MM UUID If Needed
                If mySSO_IM_UUID.Trim <> mySSOUUID.Trim Then
                    myreturn = SaveAppUserSSOUUID(myAppUserId, mySSO_IM_UUID, "Y")
                    If myreturn <> "." Then
                        myMessage = "Sub: Create_Update_SSO_Accounts After Link Accts, Process SaveAppUserSSOUUID Failed: AppUserId:" & myAppUserId & " ReturnMsg: " & myreturn.ToString
                        gSSOInterfaceFailing = myMessage
                        AppendText(myMessage)
                        WriteToErrorEmail(myMessage)
                        GoTo PROCESSFAILED
                    End If
                End If

            End If 'Link Accounts Add Principle END


            'Check if OldAppUserId account exists and set accounts in-active ===============================================
            'Account Manager IntelliMessage Client ID Screen Rename
            '1.	When MM IM account is renamed, MM process will update table AppUser  field Old App User ID with the Prior Account ID 
            '2.	After creating new SSO IM account will need to set old SSO IM Account Inactive
            '3.	After creating new SSO AM account, check the App User field Primary Email for the account and see if it exists on any other accounts.  
            '4.	 If the Primary Phone number exists on a different active account in table App User then do not deactivate the old account, otherwise, set old SSO AM Account Inactive.
            If myOldAppUserId.Trim <> "" Then 'SSO_IM
                'set old SSO_IM account inactive
                myreturn = ModifySSOUser(gSSO_IM_APIKey, myAdminToken, "", "", "", myOldAppUserId, "", "", "0")
                If myreturn <> "." Then
                    myMessage = "Sub: Create_Update_SSO_Accounts Set Old SSO_IM Acct In-Active, AppUserId:" & myOldAppUserId & " ReturnMsg: " & myreturn.ToString
                    gSSOInterfaceFailing = myMessage
                    AppendText(myMessage)
                    WriteToErrorEmail(myMessage)
                    GoTo PROCESSFAILED
                End If

                'SSO_AM
                If myOldPrimaryEmail <> "" And CueSSOAAMAccountExists(myOldAppUserId, myOldPrimaryEmail) <> "." Then 'SSO_AM
                    myreturn = ModifySSOUser(gSSO_AM_APIKey, myAdminToken, "", "", "", myOldPrimaryEmail, "", "", "0")
                    If myreturn <> "." Then
                        myMessage = "Sub: Create_Update_SSO_Accounts Set Old SSO_AM Acct In-Active, AppUserId:" & myOldPrimaryEmail & " ReturnMsg: " & myreturn.ToString
                        gSSOInterfaceFailing = myMessage
                        AppendText(myMessage)
                        WriteToErrorEmail(myMessage)
                        GoTo PROCESSFAILED
                    End If

                End If

            End If

PROCESSFAILED:

            'Return results ================================================================================================
            ' if gSSOInterfaceFailing = "" then successful
            If gSSOInterfaceFailing.Length > 254 Then
                gSSOInterfaceFailing = gSSOInterfaceFailing.Substring(0, 254)
            End If
            mySql = "exec AppUserSSOUserUpdate @AppUserId , @SSOInterfaceFailing "
            Dim myDoneAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myDoneAdapter
                Try
                    .SelectCommand.Parameters.Add("@AppUserId", SqlDbType.VarChar).Value = myAppUserId
                    .SelectCommand.Parameters.Add("@SSOInterfaceFailing", SqlDbType.VarChar).Value = gSSOInterfaceFailing
                    myData = New DataTable("Response")
                    Using myDoneAdapter
                        .Fill(myData)
                    End Using
                    Try
                        If myData.Rows(0).ItemArray.Count <= 1 Then 'Process failed

                        End If
                        myreturn = myData.Rows(0).Item("ReturnMsg")
                        If myreturn <> "." Then 'nothing to process
                            myMessage = "SP AppUserSSOUserUpdate: Failed for : AppUserId: " & _
                                              myAppUserId & vbCrLf & " Returned: " & myreturn
                            AppendText(myMessage)
                            WriteToErrorEmail(myMessage)
                            GoTo NEXTRECORD
                        End If
                    Catch ex As Exception
                        WriteToErrorEmail("Sub: Create_Update_SSO_Accounts: " + ex.Message.ToString)
                        Exit Sub
                    End Try
                    myData = Nothing
                    myAdapter = Nothing
                Catch ex As Exception
                    myData = Nothing
                    myDoneAdapter = Nothing
                    WriteToErrorEmail("Sub: Create_Update_SSO_Accounts: " + ex.Message.ToString)
                End Try
            End With
NEXTRECORD:
            myCount = myCount + 1
            If myCount = myLimit Then Exit Sub 'Getting out of loop, need to delete last transaction
        End While

    End Sub
    Function GetAdminToken() As String
        Dim retVal As String = ""
        Try
            Dim myAccount As String = ""
            Dim myPassword As String = ""
            Dim myUUID As String = ""
            Dim myActive As String = ""
            Dim myToken As String = ""

            Dim j As Integer = 0
            Dim myreturn As String = GetAppControlCharacter("AMS", "MM", "AbacusAdminAccount")
            If myreturn <> "" Then
                Dim Array() As String = myreturn.Split("|")
                myAccount = Array(0)
                myPassword = Array(1)
            End If
            myreturn = "" : myreturn = AuthenticateUser(gSSO_IM_APIKey, myAccount, myPassword, myUUID, myToken) 'returns UUID, Active, Token
            If myreturn = "." Then
                retVal = myToken
            End If
        Catch ex As Exception
            retVal = ""
        End Try
        Return retVal
    End Function





    Sub SSOAPICallLog(ByVal sCall As String, ByVal sResult As String, ByVal sSuccess As Boolean)
        'History of all calls to SSO.
        Try
            Dim myData As DataTable = Nothing
            Dim ReturnMsg As String = ""
            Dim myreturn As String = ""
            Dim mySql = "execute dbo.SSOAPICallLogWrite @Call, @Result, @Success"
            Dim myDoneAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)

            If sCall.Length > 1000 Then sCall = sCall.Substring(0, 1000)
            If sResult.Length > 1000 Then sResult = sResult.Substring(0, 1000)
            With myDoneAdapter
                Try
                    .SelectCommand.Parameters.Add("@Call", SqlDbType.VarChar).Value = sCall
                    .SelectCommand.Parameters.Add("@Result", SqlDbType.VarChar).Value = sResult
                    .SelectCommand.Parameters.Add("@Success", SqlDbType.Bit).Value = sSuccess
                    myData = New DataTable("Response")
                    Using myDoneAdapter
                        .Fill(myData)
                    End Using
                    ReturnMsg = myData.Rows(0).Item("ReturnMsg")
                    myData = Nothing
                    myDoneAdapter = Nothing
                    If ReturnMsg <> "." Then
                        WriteToErrorEmail("SSOAPICallLog: " & vbCrLf & _
                                          "Call: " & sCall & vbCrLf & _
                                         "Result: " & sResult & vbCrLf & _
                                         "Return Msg: " & ReturnMsg)
                    End If
                Catch ex As Exception
                    myData = Nothing
                    myDoneAdapter = Nothing
                    WriteToErrorEmail("SSOAPICallLog: " & ex.Message.ToString & vbCrLf & _
                                          "Call: " & sCall & vbCrLf & _
                                        "Result: " & sResult & vbCrLf & _
                                         "Return Msg: " & ReturnMsg)
                End Try
            End With
        Catch ex As Exception
            WriteToErrorEmail("SSOAPICallLog: " & ex.Message.ToString)
        End Try

    End Sub

    Function CreateSSOUser(ByVal sApiKey As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sEmail As String, ByVal sClientID As String, ByVal sPassword As String, ByVal sPrimaryPhone As String, ByRef sUUID As String) As String
        'stype = SSO Or IM
        Trace.WriteLine("Start CreateSSOUser")
        Dim retVal As String = ""
        Dim myCall As String = ""
        Dim myResult As String = ""
        Dim jdo As AbacusInterface.SSO.SSO_Subject
        Dim sStatus As String = ""
        Dim jsonResponse As String = ""
        Try
            'See if Account already exits
            sUUID = ""
            Dim createUser As New SSOCreateUser()
            With createUser
                .firstName = sFirstName
                .lastName = sLastName
                .email = sEmail
                .loginId = sClientID
                .password = sPassword
                .phoneNumber = sPrimaryPhone
            End With

            Dim json As String = JsonConvert.SerializeObject(createUser)
            Dim sURI As String = gSSO_URL & "/subject?apiKey=" & sApiKey
            Dim sysURI As New System.Uri(sURI)
            Dim data = Encoding.UTF8.GetBytes(json)
            'API can fail to return valid data error catch
            Try
                jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "POST")
                If IsvalidJson(jsonResponse) = False Then
                    retVal = "CreateSSOUser: URL: " & sURI & vbCrLf & " Body: " & json & vbCrLf & "Response: " & jsonResponse.ToString
                    SSOAPICallLog("CreateSSOUser: " & "URL: " & sURI & " Body: " & json, jsonResponse, False) 'Logging
                    Return retVal
                End If

                jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
                sStatus = jdo.status
            Catch ex As Exception
                'try a catch SSO server hang
                retVal = "CreateSSOUser: URL: " & sURI & vbCrLf & " Body: " & json & vbCrLf & "Response: " & jsonResponse.ToString
                SSOAPICallLog("CreateSSOUser: " & "URL: " & sURI & " Body: " & json, jsonResponse, False) 'Logging
                Return retVal
            End Try

            Dim sSSOUUID As String = ""
            myCall = "Json: " & json.ToString & vbCrLf & "sURI: " & sURI
            myResult = jsonResponse
            If sStatus = "SUCCESS" Then
                sUUID = jdo.subject.uuid
                SSOAPICallLog("CreateSSOUser: " & sURI, jsonResponse, True) 'Logging
                retVal = "."
            Else 'failed
                retVal = "CreateSSOUser: URL: " & sURI & vbCrLf & " Body: " & json & vbCrLf & "Response: " & jsonResponse.ToString
                SSOAPICallLog("CreateSSOUser: " & "URL: " & sURI & " Body: " & json, jsonResponse, False) 'Logging
                Return retVal
            End If

            jdo = Nothing
            sysURI = Nothing
            jsonResponse = Nothing
        Catch ex As System.Data.SqlClient.SqlException
            retVal = ex.Message.ToString
            SSOAPICallLog("CreateSSOUser: " & "AppUserId: " & sClientID, retVal, False) 'Logging

        End Try

        Trace.WriteLine("END CreateSSOUser")
        Return retVal

    End Function


    Function ModifySSOUser(ByVal sApiKey As String, ByVal sToken As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sEmail As String, ByVal sClientID As String, ByVal sPassword As String, ByVal sNewPassword As String, ByVal sActive As String) As String
        Trace.WriteLine("Start ModifySSOUser")
        Dim retVal As String = "."
        Dim sUUID As String = ""
        Dim jdo As SSO.SSO_Subject
        Dim jsonResponse As String = ""
        Dim sString As String = ""
        Dim sStatus As String = ""
        Try

            If sActive = "1" Then
                sActive = "true"
            End If
            If sActive = "0" Then
                sActive = "false"
            End If
            sClientID = LCase(sClientID)
            Dim sURI As String = gSSO_URL & "/subject?apiKey=" & sApiKey & "&tokenId=" & sToken & "&loginId=" & sClientID   'dcb

            Dim modifyUser As New SSOModifyUser()
            'Blank fields are not updated or reguired
            With modifyUser
                .firstName = sFirstName
                .lastName = sLastName
                .email = sEmail
                .loginId = sClientID
                .password = sPassword
                .newPassword = sNewPassword
                .active = sActive
            End With

            Dim json As String = JsonConvert.SerializeObject(modifyUser)
            Dim sysURI As New System.Uri(sURI)
            Dim data = Encoding.UTF8.GetBytes(json)



            'API can fail to return valid data error catch
            Try
                jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "PUT")
                If IsvalidJson(jsonResponse) = False Then
                    retVal = "ModifySSOUser: URL: " & sURI & vbCrLf & " Body: " & json & vbCrLf & "Response: " & jsonResponse.ToString
                    SSOAPICallLog("ModifySSOUser: " & "URL: " & sURI & " Body: " & json, jsonResponse, False) 'Logging
                    Return retVal
                End If

                jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
                sStatus = jdo.status
            Catch ex As Exception
                retVal = "ModifySSOUser: URL: " & sURI & vbCrLf & " Body: " & json & vbCrLf & "Response: " & jsonResponse.ToString
                SSOAPICallLog("ModifySSOUser: " & "URL: " & sURI & " Body: " & json, jsonResponse, False) 'Logging
                Return retVal
            End Try


            Dim myResult As String = jsonResponse
            If sStatus = "SUCCESS" Then
                SSOAPICallLog("ModifiySSOUser: " & sURI, jsonResponse, True) 'Logging
                retVal = "."
            Else 'failed
                retVal = myResult
                SSOAPICallLog("ModifySSOUser: " & sURI, jsonResponse, True) 'Logging
            End If

        Catch ex As Exception
            retVal = ex.Message.ToString
            SSOAPICallLog("ModifySSOUser: " & "AppUserId: " & sClientID, retVal, False) 'Logging
        End Try
        Trace.WriteLine("End ModifySSOUser")
        Return retVal

    End Function

    Function SaveAppUserSSOUUID(ByVal sAppUserID As String, ByVal sSSOUUID As String, Optional ByVal sReplace As String = "") As String
        ' Returns '.' or 'An error occurred saving the Application User Info.'
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUserSSOUUID"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@SSOUUID", sSSOUUID)
            .Parameters.AddWithValue("@Replace", sReplace)
        End With

        cn.Open()
        retValue = cmdSQL.ExecuteScalar()

        If Not cmdSQL Is Nothing Then
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End If

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception
            WriteToErrorEmail("SaveAppUserSSOUUID: " & ex.Message.ToString)
        End Try

        Return retValue
    End Function
    Function AuthenticateUser(ByVal sApiKey As String, ByVal sAppUserID As String, ByVal sPassword As String, ByRef sUUID As String, ByRef sToken As String, Optional ByRef sAcctStatus As String = "") As String
        Trace.WriteLine("Start AuthenticateUser")

        sToken = ""
        sUUID = ""
        sAcctStatus = ""

        Dim myreturn As String = ""
        Dim myCall As String = ""
        Dim myResult As String = ""
        Dim sURI As String = ""

        Dim myAccountActive As String = ""
        Dim myRoles(20) As String
        Dim myalternateId(20) As String
        Dim i As Integer = 0
        Dim myInvalidPassword As Boolean = False
        Dim sStatus As String = ""
        Dim jdo As AbacusInterface.SSO.SSO_Authenticate
        Dim jsonResponse As String = ""
        Try
            'sAppUserID = "akausel.mshs@dev.intellimsg.net" : sPassword = "Test1234"
            'sSSO_URL = "http://qa.security.americanmessaging.net"
            'Authentication (Single Sign On) API
            sURI = gSSO_URL & "/authenticate?apiKey=" & sApiKey & "&loginId=" & sAppUserID & "&password=" & HttpUtility.UrlEncode(sPassword)

            Try
                jsonResponse = SSO.GetJSONDownloadString(sURI)
                If IsvalidJson(jsonResponse) = False Then
                    myreturn = "AuthenticateUser: URL: " & sURI & vbCrLf & "Response: " & jsonResponse.ToString
                    SSOAPICallLog("AuthenticateUser: " & "URL: " & sURI, jsonResponse, False) 'Logging
                    Return myreturn
                End If

                jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Authenticate)(jsonResponse)
                sStatus = jdo.status
            Catch ex As Exception
                myreturn = "AuthenticateUser failed: URL: " & gSSO_URL & vbCrLf & "Response: " & jsonResponse.ToString
                SSOAPICallLog("AuthenticateUser failed: " & "URL: " & gSSO_URL, jsonResponse, False) 'Logging
                Return myreturn
            End Try

            myCall = "sURI: " & sURI
            myResult = jsonResponse

            If sStatus = "SUCCESS" Then
                myreturn = "."
                SSOAPICallLog("AuthenticateUser: " & myCall, myResult, True) 'Logging
                'Active = jdo.subject.active [does not work]
                sToken = jdo.token.tokenId
                sAcctStatus = "Okay"

                ' myRoles = jdo.subject.roles.ToString
                sUUID = jdo.subject.uuid.ToString
                'For i = 0 To jdo.subject.alternateIds.Count - 1
                '    myalternateId(i) = jdo.subject.alternateIds(i).alternateId.ToString
                'Next
                'For i = 0 To jdo.subject.roles.Count - 1
                '    myRoles(i) = jdo.subject.roles(i).name.ToString
                'Next
            Else

                If myResult.ToLower.IndexOf("invalid 'loginid'") >= 0 Then
                    sAcctStatus = "invalidaccount"
                End If
                If myResult.ToLower.IndexOf("invalid 'password'") >= 0 Then
                    sAcctStatus = "invalidpassword"
                End If

                If myResult.ToLower.IndexOf("account is inactive") >= 0 Then
                    sAcctStatus = "accountinactive"
                End If

                myreturn = jdo.error.message
                SSOAPICallLog("AuthenticateUser: " & myCall, myResult, False) 'Logging
            End If
            jdo = Nothing
        Catch ex As System.Net.WebException
            WriteToErrorEmail("AuthenticateUser: " & ex.Message.ToString & vbCrLf & _
               "Call: " & myCall & vbCrLf & _
               "Result: " & myResult)

        Catch ex As Exception
            WriteToErrorEmail("AuthenticateUser: " & ex.Message.ToString & vbCrLf & _
               "Call: " & myCall & vbCrLf & _
               "Result: " & myResult)
        End Try

        Trace.WriteLine("END AuthenticateUser")
        Return myreturn
    End Function
    Function AddPrincipal(ByVal mySSOAdminToken As String, ByVal sEmailAddress As String, ByVal mySSO_IM_UUID As String)
        Dim retVal As String = ""
        Dim myCall As String = ""
        Dim myResult As String = ""
        Dim jsonResponse As String = ""
        ' Dim myappUuid As String = "8b3f771c-70d5-4318-a250-506135e51ba4" 'per shawn 7/22/2014 per email

        ' 'Goal: Link SSO_IM and SSO_AM accounts
        '  Properties()
        'Token              =  SSO_IM Account token admin user token  INTELLIMSG_ADMIN = abacus
        'Principle          =  SSO_AM Account email
        'Uuid               =  SSO_AM  Account UUID
        'appUuid            =  "33fea87d-9c26-4dcd-8b91-c2e31bd3a92f"                                                       OLD: SSO_IM Account UUID

        Try

            'Dim AddPrincipalUpdate As New SSOAddPrincipal()
            'With AddPrincipalUpdate
            '    .principal = sEmailAddress
            '    .uuid = mySSO_IM_UUID
            '    .appUuid = "33fea87d-9c26-4dcd-8b91-c2e31bd3a92f" 'mySSO_AM_UUID
            'End With
            Dim json As String = "" ' JsonConvert.SerializeObject(AddPrincipalUpdate)
            'Dim sURI As String = gSSO_URL & "/subject/principal?apiKey=" & gApiKey & "&tokenId=" & mySSOAdminToken
            'Dim sURI As String = gSSO_URL & "/subject/principal?apiKey=" & gApiKey & "&tokenId=" & _
            '    mySSOAdminToken & "&principal=" & sEmailAddress & "&uuid=" & mySSO_IM_UUID & "&appUuid=" & gApiKey
            '08/18/2014 change ApiKey to SSO_AM in code

            Dim sURI As String = gSSO_URL & "/subject/principal?apiKey=" & gSSO_AM_APIKey & "&tokenId=" & _
                mySSOAdminToken & "&principal=" & sEmailAddress & "&uuid=" & mySSO_IM_UUID & "&appUuid=" & gSSO_AM_APIKey
            Dim sysURI As New System.Uri(sURI)
            Dim data = Encoding.UTF8.GetBytes(json)
            jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "POST")
            Dim jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            myCall = "Json: " & json.ToString & vbCrLf & "sURI: " & sURI
            myResult = jsonResponse
            If jdo.status = "SUCCESS" Then
                SSOAPICallLog("AddPrincipal: " & sURI, jsonResponse, True) 'Logging
                'mySSO_IM_UUID = jdo.subject.uuid 'it's been changed
                retVal = "."
            Else 'failed
                retVal = myResult
                SSOAPICallLog("AddPrincipal: " & sURI, jsonResponse, False) 'Logging
            End If

        Catch ex As Exception
            retVal = ex.Message.ToString
            SSOAPICallLog("AddPrincipal: " & "AppUserId: " & sEmailAddress, jsonResponse, False) 'Logging

        End Try
        Return retVal
    End Function

    Function CueSSOAAMAccountExists(ByVal sOldAppUserID As String, ByRef sOldPrimaryEmail As String) As String
        Dim retVal As String = ""
        Dim myMessage As String = ""
        Dim myData As DataTable
        Try
            'Returns . if none found or returns first active account found with same PrimaryEmailAddress
            Dim mySql As String = "exec CueSSOAMAccountExists @AppUserId , @PrimaryEmail "
            Dim myAdapter As New SqlClient.SqlDataAdapter(mySql, gDatabaseConnectionString)
            With myAdapter
                Try
                    .SelectCommand.Parameters.Add("@AppUserID", SqlDbType.VarChar).Value = sOldAppUserID
                    .SelectCommand.Parameters.Add("@PrimaryEmail", SqlDbType.VarChar).Value = sOldPrimaryEmail
                    myData = New DataTable("Response")
                    Using myAdapter
                        .Fill(myData)
                    End Using
                    Try
                        If myData.Rows(0).ItemArray.Count < 1 Then 'Process failed
                            myMessage = "Process: CueSSOAAMAccountExists, SP: CueSSOAMAccountExists for : AppUserId: " & sOldAppUserID
                            retVal = myMessage
                            AppendText(myMessage)
                            WriteToErrorEmail(myMessage)
                            Exit Try
                        End If
                        retVal = myData.Rows(0).Item("ReturnMsg")
                    Catch ex As Exception
                        myMessage = "Process: CueSSOAAMAccountExists, SP: CueSSOAMAccountExists for : AppUserId: " & sOldAppUserID & "   " & ex.Message
                        WriteToErrorEmail(myMessage)
                    End Try
                    myData = Nothing
                    myAdapter = Nothing
                Catch ex As Exception
                    myData = Nothing
                    myAdapter = Nothing
                    myMessage = "Process: CueSSOAAMAccountExists, SP: CueSSOAMAccountExists for : AppUserId: " & sOldAppUserID & "   " & ex.Message
                    WriteToErrorEmail(myMessage)
                End Try
            End With
        Catch ex As Exception

        End Try
        Return retVal
    End Function

    Function IsvalidJson(ByVal sjson As String) As Boolean
        'Needs Imports Newtonsoft.Json.Linq
        'pass in json string
        Dim retval As Boolean = False
        Try
            'if throws an error not valid json
            'JsonSchema.Parse(sjson)
            Dim stoken As Newtonsoft.Json.Linq.JToken = JContainer.Parse(sjson)
            retval = True
        Catch ex As Exception
            Return retval
        End Try
        Return retval
    End Function

    Sub PurgeOldCueLogFiles()
        Dim myCUEPurgeLogsDays As String = ""
        Dim myPath As String = ""
        Dim myPathFilename As String = ""
        Dim myFileNameParts() As String
        Dim myFileName As String = ""
        Dim myDate As Date
        Dim myDiff As Long = 0
        Dim myResult As Boolean = 0
        Try
            myCUEPurgeLogsDays = GetAppControlCharacter("AMS", "MM", "CUEPurgeLogs")

            myPath = GetAppControlCharacter("AMS", "MM", "CueLogsUploadPath")
            If My.Computer.Name.ToLower = "lrw7" Then
                myPath = "\\devrest\c\inetpub\wwwroot\Admin\downloads\cueLogs\"
            End If

            Dim Logfiles() As String = Directory.GetFiles(myPath, "*.*")

            For i = 0 To Logfiles.Length - 1
                myPathFilename = Logfiles(i)
                myFileNameParts = myPathFilename.Split("\")
                myFileName = myFileNameParts(myFileNameParts.Length - 1) 'last part
                Try
                    myDate = FileDateTime(myPathFilename)
                Catch ex As Exception
                    GoTo NEXTFILE
                End Try
                myDiff = DateDiff(DateInterval.Day, myDate, DateTime.Now)


                'Purge file and SQL table record
                If myDiff > myCUEPurgeLogsDays Then
                    File.Delete(myPathFilename)
                    myResult = DeleteCueLogfiles(myFileName)
                End If

                ' Trace.WriteLine(myPathFilename)

NEXTFILE:
            Next


        Catch ex As Exception
            WriteToErrorEmail("PurgeOldCueLogFiles: " & ex.Message.ToString)
        End Try
    End Sub
    Public Function DeleteCueLogfiles(ByVal sFileName As String) As Boolean
        Dim myConn As openAppCn = New openAppCn
        Dim myReturn As Boolean = 0
        Dim cn As New SqlConnection(myConn.cnString)
        'Dim cn As New SqlConnection(gDatabaseConnectionString)
        Dim cmdSQL As New SqlCommand
        Dim da As New SqlDataAdapter
        Dim dt As New DataTable
        Try

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.Text
            .CommandText = "Delete CueLogFiles where FileName = '" & sFileName & "'"

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)
            myReturn = True


        If Not cmdSQL Is Nothing Then
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End If

        If Not da Is Nothing Then
            da.Dispose()
            da = Nothing
        End If

        If Not dt Is Nothing Then
            dt.Dispose()
            dt = Nothing
        End If


        cn.Close()
        cn.Dispose()
        cn = Nothing
        Catch ex As Exception

        End Try
        Return myReturn
    End Function

    Private Sub CureatrAlert()
        Try
            'gCureatrAlertCount
            'gCureatrAlertSuccess
            'gCureatrAlertLastEmailSent
            Dim sURI As String = gCureatrAlert
            Dim myURIRequest As String = gCureatrAlert
            Dim mySendSuccess As String = ""
            Dim myURIResponse As String = ""

            Dim pFrom As String = gEmailFromAddress     'If blank will be set SendMailMessage
            Dim pRecipient As String = ""  'If blank will be set SendMailMessage
            Dim pBCC As String = ""
            Dim pCC As String = ""
            Dim pSubject As String = "Cureatr Alert Failed"
            Dim pBody As String = "Cureatr Alert "
            Dim myTimeNow As Date = Now

            'myURIRequest = My.Settings.CureatrAlertPROD

            Dim wMin As Long = DateDiff(DateInterval.Minute, gCureatrAlertLastEmailSent, myTimeNow)
            mySendSuccess = SendHTTPRequest(myURIRequest, myURIResponse)
            If myURIResponse.IndexOf("Web Services not available") > 0 Then
                mySendSuccess = False

            End If
            'mySendSuccess = True
            If mySendSuccess = False Then
                If gCureatrAlertCount < 3 Then                      'Send emails 3 times on fail
                    gCureatrAlertSuccess = False
                    gCureatrAlertCount = gCureatrAlertCount + 1
                    gCureatrAlertLastEmailSent = Now
                    pBody = "Web Services not available"
                    AppendText("Cureatr Alert Failed " & Now)
                    GoTo SENDEMAIL
                End If
                If gCureatrAlertCount >= 3 And wMin > 15 Then         'Wait 15 min and resend warning emails
                    gCureatrAlertSuccess = False
                    gCureatrAlertCount = 1
                    gCureatrAlertLastEmailSent = Now
                    pBody = "Web Services not available"
                    AppendText("Cureatr Alert Failed " & Now)
                    GoTo SENDEMAIL
                End If
            Else
                If gCureatrAlertSuccess = False Then                'Send one successful email and reset counts
                    gCureatrAlertSuccess = True
                    gCureatrAlertCount = 0
                    gCureatrAlertLastEmailSent = Nothing
                    pBody = "Cureatr Alert Successful"
                    pSubject = "Cureatr Alert Successful"
                    AppendText("Cureatr Alert Successful " & Now)
                    GoTo SENDEMAIL
                End If
            End If

            Exit Sub

SENDEMAIL:
            SendMailMessage(pFrom, pRecipient, pBCC, pCC, pSubject, pBody)


        Catch ex As Exception
            WriteToErrorEmail("PurgeOldCueLogFiles: " & ex.Message.ToString)
        End Try





    End Sub
End Class
