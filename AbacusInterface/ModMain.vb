Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Imports System.Net.Mail
Module ModMain


    Public Function AdminChangePasswordAppUser2(ByVal sAppUserID As String, ByVal sNewPassword As String, ByVal sLoggedInUser As String) As String
        Dim retValue As String = ""
        Try
            Dim myConn As openAppCn = New openAppCn
            Dim cn As New SqlConnection(myConn.cnString)
            Dim cmdSQL As New SqlCommand

            With cmdSQL
                .Connection = cn
                .CommandType = Data.CommandType.StoredProcedure
                .CommandText = "AdminChangePasswordAppUser2"
                .Parameters.AddWithValue("@AppUserId", sAppUserID)
                .Parameters.AddWithValue("@NewPassword", sNewPassword)
                .Parameters.AddWithValue("@LoggedInAppUserId", sLoggedInUser)
            End With

            cn.Open()
            retValue = cmdSQL.ExecuteScalar()
            cn.Close()

        Catch ex As Exception

            WriteToErrorEmail(ex.Message.ToString)
        End Try
        Return retValue
    End Function

    Public Function AdminChangePasswordAppUser(ByVal sAppUserID As String, ByVal sNewPassword As String, ByVal sNewPassword2 As String, ByVal sCurrentPassword As String, ByVal sLoggedInUser As String) As String
        Dim retValue As String = ""
        Try
            Dim myConn As openAppCn = New openAppCn
            Dim cn As New SqlConnection(myConn.cnString)
            Dim cmdSQL As New SqlCommand

            With cmdSQL
                .Connection = cn
                .CommandType = Data.CommandType.StoredProcedure
                .CommandText = "AdminChangePasswordAppUser"
                .Parameters.AddWithValue("@AppUserId", sAppUserID)
                .Parameters.AddWithValue("@NewPassword", sNewPassword)
                .Parameters.AddWithValue("@NewPassword2", sNewPassword2)
                .Parameters.AddWithValue("@CurrentPassword", sCurrentPassword)
                .Parameters.AddWithValue("@LoggedInAppUserId", sLoggedInUser)
            End With

            cn.Open()
            retValue = cmdSQL.ExecuteScalar()
            cn.Close()
        Catch ex As Exception

            WriteToErrorEmail(ex.Message.ToString)
        End Try
        Return retValue
    End Function

    Public Function GetAppControlCharacter(sCompanyCode As String, sModuleCode As String, sApplicationControlCode As String) As String
        Dim sCharacterParameter As String = ""
        Try
            Dim dt As DataTable
            dt = ModMain.GetApplicationControl(sCompanyCode, sModuleCode, sApplicationControlCode)
            If dt.Rows.Count > 0 Then
                sCharacterParameter = dt.Rows(0)("CharacterParameter").ToString()
            End If

            If Not dt Is Nothing Then
                dt.Dispose()
                dt = Nothing
            End If
        Catch ex As Exception

            WriteToErrorEmail(ex.Message.ToString)
        End Try
        Return sCharacterParameter
    End Function

    Public Function GetApplicationControl(ByVal sCompanyCode As String, ByVal sModuleCode As String, ByVal sApplicationControlCode As String) As DataTable
        Dim myConn As openAppCn = New openAppCn

        Dim cn As New SqlConnection(myConn.cnString)
        'Dim cn As New SqlConnection(gDatabaseConnectionString)
        Dim cmdSQL As New SqlCommand
        Dim da As New SqlDataAdapter
        Dim dt As New DataTable

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterApplicationControls"
            .Parameters.AddWithValue("@CompanyCode", sCompanyCode)
            .Parameters.AddWithValue("@ModuleCode", sModuleCode)
            .Parameters.AddWithValue("@ApplicationControlCode", sApplicationControlCode)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetApplicationControl = dt

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


    End Function
  
    Public Function SendEmail(ByVal pText As String) As Boolean
        Dim myreturn As Boolean = False
        Dim mySubject As String = ""
        Dim myBody As String = ""
        Dim myToAddress As String = ""
        Dim myBCC As String = ""
        Dim myCC As String = ""
        Dim myFrom As String = ""
        Dim myHost As String = ""
        Dim myTest As Boolean = False
        Dim myData As DataTable = Nothing
        Dim mysql As String = ""
        Try
            mySubject = Environment.MachineName & " ABACUSInterface Requires Attention!"
            myBody = pText
            myFrom = frmMain.gEmailFromAddress

            If myFrom.Length = 0 Then
                Exit Function
            End If

            'Build temp table
            myData = New DataTable
            myData.Columns.Add("EmailAddress", Type.GetType("System.String"))

            'add rows
            Dim myNewRow As DataRow = myData.NewRow
            myData.Rows.Add("daleb@abw.com")
            myData.Rows.Add("larryr@abw.com")
            myData.Rows.Add("zerrialb@abw.com")
            If pText.IndexOf("SSO - ") > 0 Then
                myData.Rows.Add("Peter.Barnett@americanmessaging.net")
                myData.Rows.Add("Amy.Williams@americanmessaging.net")

            End If

            For Each myRow As DataRow In myData.Rows
                myToAddress = myRow.Item("EmailAddress").ToString.Trim
                mysql = "execute dbo.SendEmail @From, @To, @Subject, @Body, @ReturnMsg, @StoredProc"
                Dim myAdapter As New SqlClient.SqlDataAdapter(mysql, frmMain.gDatabaseConnectionString)
                With myAdapter
                    Try
                        .SelectCommand.Parameters.Add("@From", SqlDbType.VarChar).Value = myFrom
                        .SelectCommand.Parameters.Add("@To", SqlDbType.Char).Value = myToAddress
                        .SelectCommand.Parameters.Add("@Subject", SqlDbType.Char).Value = mySubject
                        .SelectCommand.Parameters.Add("@Body", SqlDbType.Char).Value = "The ABACUS Interface has failed for the following reason: " & vbCrLf & vbCrLf & pText
                        .SelectCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar).Value = ""
                        .SelectCommand.Parameters.Add("@StoredProc", SqlDbType.VarChar).Value = "0"

                        myData = New DataTable("Response")
                        Using myAdapter
                            .Fill(myData)
                        End Using

                        If myData.Rows(0) Is Nothing Then
                            Exit For 'process failed exit and try again after frmMain.gWaitTimeBetweenErrorEmails
                        End If
                        If myData.Rows(0).ItemArray.Count < 1 Then
                            Exit For 'process failed exit and try again after frmMain.gWaitTimeBetweenErrorEmails
                        Else
                            If myData.Rows(0).Item("ReturnMsg") <> "." Then 'SQl Email failed
                                Exit For 'process failed exit and try again after frmMain.gWaitTimeBetweenErrorEmails
                            End If
                        End If
                    Catch ex As Exception
                        AppendText(ex.Message.ToString)
                        Exit For
                    End Try
                End With

            Next
            myreturn = True ' success
            myData = Nothing

        Catch ex As Exception
            AppendText(ex.Message.ToString)
        End Try
        Return myreturn
    End Function
    Public Sub SendMailMessage_OLD(ByVal pFrom As String, ByVal pRecipient As String, ByVal pBCC As String, ByVal pCC As String, ByVal pSubject As String, ByVal pBody As String)
        Dim myMailMessage As MailMessage

        Try
            ' Instantiate a new instance of MailMessage
            myMailMessage = New MailMessage()

            ' Set the sender address of the mail message
            myMailMessage.From = New MailAddress(pFrom)

            ' Set the recipient address of the mail message
            myMailMessage.To.Add(New MailAddress(pRecipient))

            ' Check if the bcc value is nothing or an empty string
            If Not pBCC Is Nothing And pBCC <> String.Empty Then
                ' Set the Bcc address of the mail message
                myMailMessage.Bcc.Add(New MailAddress(pBCC))
            End If

            ' Check if the cc value is nothing or an empty value
            If Not pCC Is Nothing And pCC <> String.Empty Then
                ' Set the CC address of the mail message
                myMailMessage.CC.Add(New MailAddress(pCC))
            End If

            ' Set the subject of the mail message
            myMailMessage.Subject = pSubject

            ' Set the body of the mail message
            myMailMessage.Body = pBody

            ' Set the format of the mail message body as HTML
            myMailMessage.IsBodyHtml = False

            ' Set the priority of the mail message to normal
            myMailMessage.Priority = MailPriority.Normal

            ' Instantiate a new instance of SmtpClient
            Dim mSmtpClient As New SmtpClient()
            'mSmtpClient.Host = pHost

            ' Send the mail message
            ' setup for failover
            Try

                mSmtpClient.Host = frmMain.gEmailSmtpHost1.ToString
                mSmtpClient.Send(myMailMessage)
            Catch ex As Exception
                AppendText(ex.Message.ToString)
                mSmtpClient.Host = frmMain.gEmailSmtpHost2.ToString
                mSmtpClient.Send(myMailMessage)
            End Try

        Catch ex As Exception
            AppendText(ex.Message.ToString)
        End Try
    End Sub
    Public Function SendMailMessage(ByVal pFrom As String, ByVal pRecipient As String, ByVal pBCC As String, ByVal pCC As String, ByVal pSubject As String, ByVal pBody As String)
        Dim myreturn As Boolean = False
        Dim mySubject As String = pSubject
        Dim myBody As String = pBody
        Dim myToAddress As String = pRecipient
        Dim myBCC As String = pBCC
        Dim myCC As String = pCC
        Dim myFrom As String = pFrom
        Dim myHost As String = ""
        Dim myTest As Boolean = False
        Dim myData As DataTable = Nothing
        Dim mysql As String = ""
        Try
            If mySubject = "" Then
                mySubject = Environment.MachineName & " ABACUSInterface Requires Attention!"
            End If

            If myFrom = "" Then
                myFrom = frmMain.gEmailFromAddress
            End If

            If myFrom.Length = 0 Then
                Return myreturn
            End If

            'Build temp table
            myData = New DataTable
            myData.Columns.Add("EmailAddress", Type.GetType("System.String"))

            'add rows
            Dim myNewRow As DataRow = myData.NewRow
            If pRecipient = "" Then
                myData.Rows.Add("daleb@abw.com")
                myData.Rows.Add("larryr@abw.com")
                myData.Rows.Add("zerrialb@abw.com")
            Else
                myData.Rows.Add(pRecipient)
            End If

            For Each myRow As DataRow In myData.Rows
                myToAddress = myRow.Item("EmailAddress").ToString.Trim
                mysql = "execute dbo.SendEmail @From, @To, @Subject, @Body, @ReturnMsg, @StoredProc"
                Dim myAdapter As New SqlClient.SqlDataAdapter(mysql, frmMain.gDatabaseConnectionString)
                With myAdapter
                    Try
                        .SelectCommand.Parameters.Add("@From", SqlDbType.VarChar).Value = myFrom
                        .SelectCommand.Parameters.Add("@To", SqlDbType.Char).Value = myToAddress
                        .SelectCommand.Parameters.Add("@Subject", SqlDbType.Char).Value = mySubject
                        .SelectCommand.Parameters.Add("@Body", SqlDbType.Char).Value = myBody
                        .SelectCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar).Value = ""
                        .SelectCommand.Parameters.Add("@StoredProc", SqlDbType.VarChar).Value = "0"

                        myData = New DataTable("Response")
                        Using myAdapter
                            .Fill(myData)
                        End Using

                        If myData.Rows(0) Is Nothing Then
                            Exit For 'process failed exit and try again after frmMain.gWaitTimeBetweenErrorEmails
                        End If
                        If myData.Rows(0).ItemArray.Count < 1 Then
                            Exit For 'process failed exit and try again after frmMain.gWaitTimeBetweenErrorEmails
                        Else
                            If myData.Rows(0).Item("ReturnMsg") <> "." Then 'SQl Email failed
                                Exit For 'process failed exit and try again after frmMain.gWaitTimeBetweenErrorEmails
                            End If
                        End If
                    Catch ex As Exception
                        AppendText(ex.Message.ToString)
                        Exit For
                    End Try
                End With

            Next
            myreturn = True ' success
            myData = Nothing

        Catch ex As Exception
            AppendText(ex.Message.ToString)
        End Try
        Return myreturn
    End Function
    Public Sub AppendText(ByVal pText As String)
        frmMain.txtText.AppendText(Format(DateTime.Now, "MM/dd/yy HH:mm:ss") & " " & pText & vbCrLf)
        Application.DoEvents()
    End Sub

    Public Sub LogDeviceToDesktop(ByVal Comment As String)
        'If gErrorLogWrite = False Then Exit Sub
        Dim myFileName As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & "ABACUSInterface_Log.txt"
        Dim sw As System.IO.StreamWriter
        Dim myErrorNo As String = ""
        Dim myErrorContent As String = ""
        Dim datastring As String = ""
        Try
            datastring = Comment & vbCr
            sw = New System.IO.StreamWriter(myFileName, True)
            sw.WriteLine(datastring)
            sw.Flush()
            sw.Close()
            sw = Nothing
        Catch ex As Exception
          
        End Try
    End Sub
    Public Function WriteToErrorEmail(ByRef data As String) As String
        Dim myreturn As String = "failed"
        Try
            data = Date.Now & " | " & data & vbCrLf
            LogDeviceToDesktop(data)
            AppendText(data)
            Dim N As Integer = frmMain.gDBErrorEmails.Columns("auto").AutoIncrement
            frmMain.gDBErrorEmails.Rows.Add(N, data)
            myreturn = "."
        Catch ex As Exception
            AppendText(ex.Message.ToString)
        End Try
        Return myreturn
    End Function
    Public Sub ProcessErrorMessages()

        Dim myCurrentTime As Date = System.DateTime.Now
        ' gWaitTimeBetweenErrorEmails = 5 min
        If CInt(DateDiff(DateInterval.Minute, frmMain.gTimelastErrorEmail, myCurrentTime)) < frmMain.gWaitTimeBetweenErrorEmails Then
            Exit Sub
        End If
        frmMain.gTimelastErrorEmail = System.DateTime.Now
        Dim j As Integer = 0
        Dim i As Integer = 0
        Dim toDelete As New List(Of DataRow)

        For Each Row As DataRow In frmMain.gDBErrorEmails.Rows
            Try
                If SendEmail(Row.Item("ErrorMessage").ToString.Trim) = True Then
                    toDelete.Add(Row) ''Write to list FOR LATER DELETE
                Else
                    AppendText("Not Able to send email via SQL dbo.SendEmail ")
                    Exit For
                End If
                i = i + 1
                If i >= 10 Then 'MAX # TO PROCESS BEFORE R
                    Exit For
                End If
            Catch ex As Exception
                AppendText(ex.Message.ToString)
            End Try
           
        Next

        'delete processed rows
        Try
            For Each row As DataRow In toDelete
                Trace.WriteLine(frmMain.gDBErrorEmails.Rows.Count)
                row.Delete()
            Next
            frmMain.gDBErrorEmails.AcceptChanges()
            toDelete = Nothing
            Trace.WriteLine(frmMain.gDBErrorEmails.Rows.Count)
        Catch ex As Exception
            AppendText("ProcessErrorMesssages delete processed records failed.")
        End Try



    End Sub
End Module
