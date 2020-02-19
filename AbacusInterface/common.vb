Imports Microsoft.VisualBasic
Imports System.Data
Imports Newtonsoft.Json

Public Class common

    Public Shared Function LoginUser(Optional ByVal sUserID As String = "", Optional ByVal sPassword As String = "") As String
        Dim sAppCode As String = common.GetAppControlCharacter("AMS", "MM", "AppCode")
        Dim sUserIsProgrammer As String = System.Configuration.ConfigurationManager.AppSettings("UserIsProgrammer")
        Dim sDevTestProd As String = common.GetAppControlCharacter("AMS", "MM", "DevTestProd")

        LoginUser = ""
        Dim ctx = HttpContext.Current
        ctx.Session("AppCode") = common.GetAppControlCharacter("AMS", "MM", "AppCode")
        Dim dt As DataTable

        Dim sReturnMsg As String = ""
        Dim sRedirectURL As String = ""
        Dim sRedirPath As String = ""
        Dim sCurrentPageName As String = LCase(curPageName)
        Dim queryFromPage As String = Trim(ctx.Request.QueryString("fromPage") & "")

        sRedirPath = Trim(ctx.Request.QueryString("redirPath")) & ""


        'ctx.Response.Write("sCurrentPageName: " & sCurrentPageName)
        'ctx.Response.End()

        'Check login credentials                     
        ctx.Session("LoggedInAtLeastOnce") = False 'If 0 then this is their first login to the site
        'ctx.Response.Write("<br>UseSSO: " & ctx.Session("UseSSO"))
        'ctx.Response.Write("<br>SSOUUID: " & ctx.Session("SSOUUID"))
        'ctx.Response.Write("<br>sDevTestProd: " & sDevTestProd)
        'ctx.Response.End()

        If ctx.Session("UseSSO") = "" Then
            ctx.Session("UseSSO") = System.Configuration.ConfigurationManager.AppSettings("UseSSO")
        End If

        If ctx.Session("UseSSO") = "1" Then
            If sDevTestProd = "DEV" And ctx.Session("SSOUUID") = "" Then
                ctx.Session("UseSSO") = "0"
                dt = DataCalls.LoginAppUser(sAppCode, LCase(sUserID), sPassword, ctx.Session("fromMPA"))
                'ctx.Response.Write("yLoginAppUser '" & sAppCode & "', '" & LCase(sUserID) & "', '" & sPassword & "', '" & ctx.Session("fromMPA") & "'")

            Else
                dt = DataCalls.LoginSSOUser(sAppCode, LCase(sUserID), ctx.Session("SSOUUID"))
                'ctx.Response.Write("xLoginSSOUser '" & sAppCode & "', '" & LCase(sUserID) & "', '" & ctx.Session("SSOUUID") & "'")
                'ctx.response.end()
            End If
        Else
            dt = DataCalls.LoginAppUser(sAppCode, LCase(sUserID), sPassword, ctx.Session("fromMPA"))
            'ctx.Response.Write("xLoginAppUser '" & sAppCode & "', '" & LCase(sUserID) & "', '" & sPassword & "', '" & ctx.Session("fromMPA") & "'")

        End If

        'ctx.Response.Write("LoginAppUser '" & sAppCode & "', '" & LCase(sUserID) & "', '" & sPassword & "', '" & ctx.Session("fromMPA") & "'")
        ' ctx.Response.End()
        If dt.Rows.Count > 0 Then
            sReturnMsg = dt.Rows(0)("ReturnMsg").ToString()
            If sReturnMsg = "." Then

                ctx.Session("AppUserID") = LCase(dt.Rows(0)("AppUserID").ToString())
                ctx.Session("LoggedInAppUserID") = LCase(dt.Rows(0)("AppUserID").ToString())
                ctx.Session("AppGroupID") = LCase(dt.Rows(0)("AppGroupID").ToString())
                ctx.Session("AppGroupIDDisplay") = dt.Rows(0)("AppGroupID").ToString()
                ctx.Session("PrimaryGroup") = LCase(dt.Rows(0)("AppGroupID").ToString())
                ctx.Session("PrimaryGroupDisplay") = dt.Rows(0)("AppGroupID").ToString()
                ctx.Session("PrimaryEmail") = dt.Rows(0)("PrimaryEmail").ToString()
                ctx.Session("AppUserName") = dt.Rows(0)("AppUserName").ToString()
                ctx.Session("TimeZoneOffsetHours") = dt.Rows(0)("TimeZoneOffsetHours").ToString()
                ctx.Session("IT") = CBool(dt.Rows(0)("IT").ToString())
                ctx.Session("LoggedInAtLeastOnce") = dt.Rows(0)("LoggedInAtLeastOnce").ToString()
                ctx.Session("EULA") = dt.Rows(0)("EULAAccepted").ToString()
                ctx.Session("UserType") = dt.Rows(0)("UserType").ToString()

            End If
        Else
            sReturnMsg = "Invalid Client ID and/or Password."

        End If

        If Not dt Is Nothing Then
            dt.Dispose()
            dt = Nothing
        End If

        If sReturnMsg = "." Then

            ctx.Session("AppUserSecurityGroup") = DataCalls.GetUserSecurityGroup(ctx.Session("LoggedInAppUserID"))
            ctx.Session("LoggedInAppUserSecurityGroup") = ctx.Session("AppUserSecurityGroup")
            ctx.Session("ShowMainSelectedTabForSubMenu") = True
            ctx.Session("AppUserSecurityGroupIT") = ctx.Session("IT")
            If Trim(ctx.Session("TimeZoneOffsetHours")) & "" <> "" Then
                ctx.Session("TimeZoneOffsetHours") = CLng(ctx.Session("TimeZoneOffsetHours"))
            Else
                ctx.Session("TimeZoneOffsetHours") = 0
            End If

            If Trim(ctx.Session("IT")) & "" <> "" Then
                ctx.Session("IT") = CBool(ctx.Session("IT"))
            Else
                ctx.Session("IT") = False
            End If

            If Trim(ctx.Session("LoggedInAtLeastOnce")) & "" <> "" Then
                ctx.Session("LoggedInAtLeastOnce") = CBool(ctx.Session("LoggedInAtLeastOnce"))
            Else
                ctx.Session("LoggedInAtLeastOnce") = False
            End If

            If Trim(ctx.Session("EULA")) & "" <> "" Then
                ctx.Session("EULA") = CBool(ctx.Session("EULA"))
            Else
                ctx.Session("EULA") = False
            End If

            If ctx.Session("UserType") = "Group" Then
                ctx.Session("AppUserSecurityGroup") = "User"
                ctx.Session("LoggedInAppUserSecurityGroup") = "User"
                ctx.Session("EULA") = True
            End If

            'Log login into database!
            Dim LogID As String = Trim(ctx.Session("LogID")) & ""

            'ctx.response.write("LogID: " & LogID)
            'ctx.response.end()
            If LogID = "" Or LogID = "0" Then
                'They probably established a session and are coming back in from Account Manager
                Dim IPAddress As String = Trim(ctx.Request.ServerVariables("Remote_addr"))
                Dim BrowserType As String = common.GetBrowserType()
                Dim BrowserUserAgent As String = Trim(ctx.Request.ServerVariables("HTTP_User_Agent"))
                Dim AspNetSessionID As String = System.Web.HttpContext.Current.Session.SessionID
                Dim ComputerName As String = common.GetComputerName
                Dim dtl As DataTable
                Dim retVal As String = ""

                dtl = DataCalls.LogUserSessionLogin2(ctx.Session("LoggedInAppUserID"), ctx.Session("AppUserID"), ctx.Session("AppGroupID"), "PCM", IPAddress, BrowserType, BrowserUserAgent, ctx.Session("UseSSO"), ComputerName, AspNetSessionID)
                If dtl.Rows.Count > 0 Then
                    retVal = dtl.Rows(0)("ReturnMsg").ToString()
                    LogID = dtl.Rows(0)("LogID").ToString()
                Else
                    LogID = "0"
                End If

                If Not dtl Is Nothing Then
                    dtl.Dispose()
                    dtl = Nothing
                End If

                ctx.Session("LogID") = LogID
                ctx.Session("ComputerName") = ComputerName
            End If


            ctx.Session("ShowMainSelectedTabForSubMenu") = ctx.Session("AppCode")


            If sCurrentPageName = "default.aspx" Or sCurrentPageName = "loginx.aspx" Then
                If ctx.Session("LoggedInAppUserSecurityGroup") = "Basic" Then
                    ctx.Session("HomePage") = "SendMessageBasic.aspx"
                Else
                    Dim sGetLastSavedCurrentPage As String = UserPreferences.GetCharacterParameter(ctx.Session("LoggedInAppUserID"), "SY", "CurrentPage")
                    If sGetLastSavedCurrentPage <> "" Then
                        ctx.Session("HomePage") = sGetLastSavedCurrentPage
                    Else
                        ctx.Session("HomePage") = "MyMessages.aspx"

                    End If
                End If

            ElseIf ctx.Session("HomePage") <> "" Then
                'leave it
            Else
                ctx.Session("HomePage") = sCurrentPageName

            End If


            If sUserIsProgrammer = "1" Then
                'Do not redirect because I am programming on my local machine!
            Else
                Dim sBetaPath As String = System.Configuration.ConfigurationManager.AppSettings("BetaPath")
                If ctx.Session("AppUserSecurityGroup") = "User" Or ctx.Session("AppUserSecurityGroup") = "Basic" Then
                    sRedirectURL = common.GetAppControlCharacter("AMS", "MM", "MyAccountRedirectURL")
                Else
                    sRedirectURL = common.GetAppControlCharacter("AMS", "MM", "MessageManagerRedirectURL")
                End If
                sRedirectURL = sRedirectURL & sBetaPath
            End If

            sRedirectURL = Trim(sRedirectURL) & ""

            If sRedirectURL <> "" Then
                If Right(sRedirectURL, 1) <> "/" Then
                    sRedirectURL = sRedirectURL & "/"
                End If
            End If

            If ctx.Session("EULA") = False Then

                sRedirectURL = sRedirectURL & "EULA.aspx"

                If InStr(sRedirectURL, "?") = 0 Then
                    sRedirectURL = sRedirectURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedirectURL = sRedirectURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If
                If queryFromPage <> "" Then
                    sRedirectURL = sRedirectURL & "&fromPage=" & queryFromPage
                End If

                ctx.Response.Redirect(sRedirectURL, False)
                ctx.Response.End()

            ElseIf ctx.Session("LoggedInAtLeastOnce") = True Then
                If sRedirPath <> "" Then

                    sRedirectURL = sRedirectURL & sRedirPath

                    If InStr(sRedirectURL, "?") = 0 Then
                        sRedirectURL = sRedirectURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirectURL = sRedirectURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If

                    If queryFromPage <> "" Then
                        sRedirectURL = sRedirectURL & "&fromPage=" & queryFromPage
                    End If
                    ctx.Response.Redirect(sRedirectURL, False)
                    ctx.Response.End()

                Else

                    sRedirectURL = sRedirectURL & Trim(ctx.Session("HomePage") & "")

                    If InStr(sRedirectURL, "?") = 0 Then
                        sRedirectURL = sRedirectURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirectURL = sRedirectURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If
                    If queryFromPage <> "" Then
                        sRedirectURL = sRedirectURL & "&fromPage=" & queryFromPage
                    End If

                    ctx.Response.Redirect(sRedirectURL, False)
                    ctx.Response.End()

                End If
            Else

                sRedirectURL = sRedirectURL & Trim(ctx.Session("HomePage") & "")

                If InStr(sRedirectURL, "?") = 0 Then
                    sRedirectURL = sRedirectURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedirectURL = sRedirectURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If
                If queryFromPage <> "" Then
                    sRedirectURL = sRedirectURL & "&fromPage=" & queryFromPage
                End If

                ctx.Response.Redirect(sRedirectURL, False)
                ctx.Response.End()

            End If
        End If

        LoginUser = sReturnMsg

    End Function

    Public Shared Function SetUserInfo() As Boolean
        'Get User Info

        SetUserInfo = False
        Dim ctx = HttpContext.Current

        ctx.Session("FromMPA") = "1"
        ctx.Session("INTELLIMSG_USER") = False
        ctx.Session("ACCOUNT_MANAGER_USER") = False
        ctx.Session("SSOUUID") = ""
        ctx.Session("UserID") = ""

        Dim sApiKey = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
        Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))
        Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""
        Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
        Dim sURI As String = sSSO_URL & "/subject?apiKey=" & sApiKey & "&tokenId=" & ctx.Session("TokenID")
        Dim jsonResponse = SSO.GetJSONDownloadString(sURI)
        Dim jdo As Object
        Dim sStatus As String = ""

        Try
            jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            sStatus = jdo.status
        Catch ex As Exception
            'sso server fail/down
            'Logged in successfully so just set this to true and move on
            ctx.Session("INTELLIMSG_USER") = True
            SetUserInfo = True
            Exit Function
        End Try

        'ctx.Response.Write("sstatus in setuserinfo: " & sStatus)
        If sStatus = "SUCCESS" Then
            'Successful SSO Get Subject
            ctx.Session("SSOUUID") = jdo.subject.uuid
            ctx.Session("UserID") = jdo.subject.principal
            For i As Integer = 0 To jdo.subject.roles.Count - 1
                ' ctx.Response.Write(jdo.subject.roles.Item(i).name)
                Select Case jdo.subject.roles.Item(i).name
                    Case "INTELLIMSG_USER"
                        ctx.Session("INTELLIMSG_USER") = True
                    Case "ACCOUNT_MANAGER_USER"
                        ctx.Session("ACCOUNT_MANAGER_USER") = True
                    Case "ACCOUNT_MANAGER_ADMIN"
                        ctx.Session("ACCOUNT_MANAGER_USER") = True

                End Select
            Next
            ' ctx.Response.End()
            If ctx.Session("INTELLIMSG_USER") = True Then
                SetUserInfo = True
                ctx.Session("GetUserInfo") = False
            Else
                SetUserInfo = False
            End If
        Else
            'Failure
            SetUserInfo = False
        End If
        'ctx.Response.Write("here in setuserinfo: " & SetUserInfo)
        'ctx.Response.End()
    End Function

    Public Shared Sub TokenRenew()
        Dim ctx = HttpContext.Current
        Dim TheCurrentDateAndTime As DateTime = Now()

        If Len(ctx.Session("DateTimeToRenewToken")) = 0 Then
            ctx.Session("DateTimeToRenewToken") = TheCurrentDateAndTime
        End If

        'ctx.Response.Write("s: " & ctx.Session("DateTimeToRenewToken"))
        'ctx.Response.Write("<br>n: " & TheCurrentDateAndTime)
        ' ctx.Response.End()
        If ctx.Session("DateTimeToRenewToken") <= TheCurrentDateAndTime Then
            'If we're past time to validate the token then re-validate it, otherwise return true and exit
            Dim sTokenID As String = Trim(HttpContext.Current.Session("TokenID")) & ""
            If sTokenID = "" Then
                sTokenID = Trim(HttpContext.Current.Request("tokenId")) & ""
            End If
            If sTokenID <> "" Then
                'ctx.Response.Write("<br>Renew Token")
                Try
                    Dim sApiKey = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
                    Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
                    Dim sURI As String = sSSO_URL & "/token/renew?apiKey=" & sApiKey & "&tokenId=" & sTokenID
                    Dim jsonResponse = SSO.GetJSONDownloadString(sURI)
                    Dim jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Token)(jsonResponse)
                    Dim sStatus As String = jdo.status
                    If sStatus = "SUCCESS" Then
                        'Successful Token Renewal
                        Dim NumMinutesHitSSO As Long = 5
                        If Len(ctx.Session("NumMinutesHitSSO")) = 0 Then
                            NumMinutesHitSSO = GetAppControlNumeric("AMS", "MM", "NumMinutesHitSSO")
                            ctx.Session("NumMinutesHitSSO") = NumMinutesHitSSO
                        Else
                            NumMinutesHitSSO = CLng(ctx.Session("NumMinutesHitSSO"))
                        End If
                        ctx.Session("DateTimeToRenewToken") = DateAdd(DateInterval.Minute, NumMinutesHitSSO, TheCurrentDateAndTime)
                    Else
                        'Failure on Token Renewal
                        ctx.Session("DateTimeToRenewToken") = TheCurrentDateAndTime
                    End If
                Catch ex As Exception

                End Try



            End If
        End If

    End Sub

    Public Shared Function TokenValidate() As Boolean

        TokenValidate = False
        Dim ctx = HttpContext.Current
        Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim TheCurrentDateAndTime As DateTime = Now()
        
        If Len(ctx.Session("DateTimeToValidateToken")) = 0 Then
            ctx.Session("DateTimeToValidateToken") = TheCurrentDateAndTime
        End If

        If ctx.Session("DateTimeToValidateToken") <= TheCurrentDateAndTime Then
            'If we're past time to validate the token then re-validate it, otherwise return true and exit
            If sUseSSO = "1" Then
                'ctx.response.write("<br>use")

                Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))
                Dim sTokenIDSession As String = Trim(ctx.Session("TokenID")) & ""
                Dim sTokenID As String = ""
                Dim sTokenIDRequest As String = Trim(ctx.Request("tokenId")) & ""

                If sTokenIDRequest <> "" And sTokenIDSession <> "" Then
                    If sTokenIDRequest <> sTokenIDSession Then
                        'changed users, re-login
                        sTokenIDSession = ""
                    End If

                End If

                'ctx.response.write("<br>t1")

                If sTokenIDRequest <> "" Then
                    sTokenID = sTokenIDRequest
                    ctx.Session("TokenID") = sTokenID
                    ctx.Session("GetUserInfo") = True
                Else
                    sTokenID = sTokenIDSession
                    ctx.Session("GetUserInfo") = False
                End If

                'ctx.response.write("<br>t2")


                If sTokenID = "" Then
                    'Logout
                    TokenValidate = False
                    'If sDevTestProd = "DEV" Then
                    'ctx.Session("UseSSO") = "0"
                    'TokenValidate = True
                    'End If
                    'ctx.response.write("<br>t3")

                Else
                    'ctx.Response.Write("<br>TokenValidate")

                    'Check for valid token
                    Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
                    Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
                    Dim sURI As String = sSSO_URL & "/token?apiKey=" & sApiKey & "&tokenId=" & ctx.Session("TokenID")
                    Dim jsonResponse = SSO.GetJSONDownloadString(sURI)
                    Dim jdo As Object
                    Dim sStatus As String = ""

                    Try
                        jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Token)(jsonResponse)
                        sStatus = jdo.status
                    Catch ex As Exception
                        'sso server fail/down
                        TokenValidate = False
                        Exit Function
                    End Try


                    'ctx.Response.Write("<br>sTokenIDRequest: " & sTokenIDRequest)
                    'ctx.Response.Write("<br>sTokenIDSession: " & sTokenIDSession)
                    'ctx.Response.Write("<br>sTokenID: " & sTokenID)
                    'ctx.Response.Write("<br>sURI: " & sURI)
                    'ctx.Response.Write("<br>sStatus: " & sStatus)

                    'ctx.Response.End()

                    If sStatus = "SUCCESS" Then
                        'Valid token, continue
                        ctx.Session("TokenID") = jdo.token.tokenId

                        Dim NumMinutesHitSSO As Long = 5
                        If Len(ctx.Session("NumMinutesHitSSO")) = 0 Then
                            NumMinutesHitSSO = GetAppControlNumeric("AMS", "MM", "NumMinutesHitSSO")
                            ctx.Session("NumMinutesHitSSO") = NumMinutesHitSSO
                        Else
                            NumMinutesHitSSO = CLng(ctx.Session("NumMinutesHitSSO"))
                        End If
                        ctx.Session("DateTimeToValidateToken") = DateAdd(DateInterval.Minute, NumMinutesHitSSO, TheCurrentDateAndTime)
                        TokenValidate = True
                    Else
                        'Invalid Token
                        ctx.Session("DateTimeToValidateToken") = TheCurrentDateAndTime
                        TokenValidate = False
                    End If

                End If

            End If
        Else
            'Already validated in the last 5 minutes, do not do it again yet
            TokenValidate = True
        End If

        'ctx.Response.Write("<br>TokenValidate: " & TokenValidate)

    End Function

    Public Shared Sub CheckSecurity(ByVal sAppScreenName As String, Optional ByVal bScreenIsPopup As Boolean = False, Optional ByVal bPopupIsModal As Boolean = False)
        CheckSecurityDual(sAppScreenName, bScreenIsPopup, bPopupIsModal)
        If 1 = 2 Then
            Dim ctx = HttpContext.Current
            Dim sTokenQstring As String = Trim(ctx.Request("TokenID")) & ""
            Dim sLogID As String = Trim(ctx.Session("LogID") & "")

            If sLogID = "" Then
                sLogID = "0"
            End If

            Dim sCurrentPageName As String = curPageName()
            If sCurrentPageName = "" Then
                sCurrentPageName = "."
            End If
            Dim sQstring As String = Trim(ctx.Request.ServerVariables("QUERY_STRING")) & ""

            If sTokenQstring <> "" And sLogID <> "0" Then



                'skip the logging because they have not logged in yet
                Dim retval As String = DataCalls.LogUserSessionDetail(sLogID, sCurrentPageName, sQstring)

                If sTokenQstring = "" Then
                    'ok to log, but skip if they are coming directly from SSO because they will never be passing a LOGID in the string
                    If Len(sQstring) <> 0 Then
                        Dim sLogIDQString As String = Trim(ctx.Request.QueryString("LogID")) & ""
                        If sLogIDQString = "" Then
                            sLogIDQString = "0"
                        End If
                        If sLogID <> sLogIDQString Then
                            Dim retval2 As String = DataCalls.LogUserSessionLogIDMismatch(sLogIDQString, sLogID, sCurrentPageName, sQstring)
                        End If
                    End If
                End If


            End If



            Dim sAppCode As String = common.GetAppControlCharacter("AMS", "MM", "AppCode")
            Dim sUserIsProgrammer As String = System.Configuration.ConfigurationManager.AppSettings("UserIsProgrammer")
            Dim sDevTestProd As String = common.GetAppControlCharacter("AMS", "MM", "DevTestProd")


            ctx.Session("AppCode") = sAppCode

            Dim sAppUserID As String = ctx.Session("LoggedInAppUserID")
            Dim ScreenAllowed As Long = 0
            Dim isSystemDown As Boolean = CheckSystemDown()
            Dim sCurrentPath As String = ""
            Dim sUseSSO As String = System.Configuration.ConfigurationManager.AppSettings("UseSSO")
            Dim sSSOLoginURL As String = LCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))
            Dim redir As String = "Default.aspx"
            Dim sReturnMsg As String = ""

            If ctx.Session("GetUserInfo") = True Then
                ctx.Session("UseSSO") = sUseSSO
                'ctx.Response.Write("<br>x1")

            ElseIf ctx.Session("UseSSO") = "0" And sDevTestProd = "DEV" Then
                sUseSSO = "0"
                'ctx.Response.Write("<br>x2")
            Else
                ctx.Session("UseSSO") = sUseSSO
                'ctx.Response.Write("<br>x3: " & sUseSSO)
            End If




            If sUserIsProgrammer = "1" Then
                sSSOLoginURL = "Default.aspx"
            End If

            If sUseSSO = "1" Then

                Dim bValidToken As Boolean = common.TokenValidate()
                If bValidToken Then
                    'Valid token, continue
                    If ctx.Session("GetUserInfo") = True Then
                        'ctx.Response.Write("here2")
                        ctx.Session("HomePage") = ""
                        Dim t As String = ctx.Session("TokenID")
                        ' ctx.Session.RemoveAll()
                        ctx.Session("TokenID") = t
                        'ctx.Response.End()
                        If SetUserInfo() = True Then
                            'Can use message manager
                            sReturnMsg = common.LoginUser(ctx.Session("UserID"), "")
                            If sReturnMsg <> "" And sReturnMsg <> "." Then

                                sSSOLoginURL = sSSOLoginURL & "?ErrMsg=" & sReturnMsg

                                If InStr(sSSOLoginURL, "?") = 0 Then
                                    sSSOLoginURL = sSSOLoginURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                                Else
                                    sSSOLoginURL = sSSOLoginURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                                End If

                                ctx.Response.Redirect(sSSOLoginURL, False)
                                ctx.Response.End()

                            End If
                        Else

                            sSSOLoginURL = sSSOLoginURL & "?ErrMsg=Could Not get user info."

                            If InStr(sSSOLoginURL, "?") = 0 Then
                                sSSOLoginURL = sSSOLoginURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                            Else
                                sSSOLoginURL = sSSOLoginURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                            End If

                            ctx.Response.Redirect(sSSOLoginURL, False)
                            ctx.Response.End()

                        End If
                    Else
                        ' ctx.Response.Write("here")
                        'ctx.Response.End()
                        TokenRenew()
                    End If

                Else
                    'Invalid token
                    sCurrentPath = IO.Path.GetFileName(ctx.Request.PhysicalPath)

                    Dim sRedURL As String = "Logout.aspx?redirPath=" & sCurrentPath & "&p=" & bScreenIsPopup & "&m=" & bPopupIsModal

                    If InStr(sRedURL, "?") = 0 Then
                        sRedURL = sRedURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedURL = sRedURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If

                    ctx.Response.Redirect(sRedURL, False)
                    ctx.Response.End()

                    ' If sUserIsProgrammer = "1" Then
                    'ctx.Response.Write("bValidTokenx: " & bValidToken)
                    'ctx.Response.End()
                    'End If
                    '''''sSSOLoginURL = LCase(sSSOLoginURL) & "?referrer=" & ctx.Request.Url.Scheme + "://" + ctx.Request.Url.Authority + ctx.Request.ApplicationPath & IO.Path.GetFileName(ctx.Request.PhysicalPath)


                    ' ctx.Response.Write("sSSOLoginURL: " & sSSOLoginURL)
                    'ctx.Response.End()

                    'sSSOLoginURL = sSSOLoginURL & "&ErrMsg=Invalid Token-a."


                End If


            End If



            If isSystemDown Then
                If ctx.Session("AppUserSecurityGroup") = "Admin" And ctx.Session("IT") = True Then
                    'Allowed to use system while shutdown
                Else

                    Dim sShutURL As String = "SystemShutdown.aspx"

                    If InStr(sShutURL, "?") = 0 Then
                        sShutURL = sShutURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sShutURL = sShutURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If


                    ctx.Response.Redirect(sShutURL, False)
                    ctx.Response.End()

                End If
            End If


            'ctx.Response.Write("sAppUserID: " & sAppUserID & "<br>")


            If sAppUserID = "" Then


                'ctx.Response.Write("useSSO: " & sUseSSO)

                'ctx.Response.End()
                If sUseSSO = "1" Then
                    If InStr(sSSOLoginURL, "?referrer=") = 0 Then
                        sSSOLoginURL = LCase(sSSOLoginURL) & "?referrer=" & ctx.Request.Url.Scheme + "://" + ctx.Request.Url.Authority + ctx.Request.ApplicationPath & IO.Path.GetFileName(ctx.Request.PhysicalPath)

                        'ctx.Response.Write("sSSOLoginURL: " & sSSOLoginURL)
                        'ctx.Response.End()

                        'sSSOLoginURL = sSSOLoginURL & "&ErrMsg=Invalid Token-a."
                    End If

                    If InStr(sSSOLoginURL, "?") = 0 Then
                        sSSOLoginURL = sSSOLoginURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sSSOLoginURL = sSSOLoginURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If

                    ctx.Response.Redirect(sSSOLoginURL, False)
                    ctx.Response.End()

                Else


                    sCurrentPath = IO.Path.GetFileName(ctx.Request.PhysicalPath)

                    Dim sLUrl As String = "Logout.aspx?redirPath=" & sCurrentPath & "&p=" & bScreenIsPopup & "&m=" & bPopupIsModal

                    If InStr(sLUrl, "?") = 0 Then
                        sLUrl = sLUrl & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sLUrl = sLUrl & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If

                    ctx.Response.Redirect(sLUrl, False)
                    ctx.Response.End()


                End If
                Exit Sub
            End If

            Dim UserGroup As String = Trim(LCase(DataCalls.GetUserGroup(ctx.Session("AppUserID"), True)))
            Dim AppScreenName As String = Trim(ctx.Session("HelpScreen"))

            If AppScreenName = "" Then
            Else
                Dim retvallog As String = DataCalls.LogViewAnotherUser(0, ctx.Session("LoggedInAppUserID"), ctx.Session("AppGroupID"), ctx.Session("AppUserID"), UserGroup, AppScreenName)

            End If

            Dim sRedirURL As String = ""

            ScreenAllowed = DataCalls.CheckSecurity(sAppUserID, sAppScreenName)
            If ScreenAllowed = 1 Then
                'Ok to view
                'Log the change

            ElseIf ScreenAllowed = 99 Then 'Inactive user or groups, log out
                sRedirURL = "Logout.aspx?p=" & bScreenIsPopup & "&m=" & bPopupIsModal

                If InStr(sRedirURL, "?") = 0 Then
                    sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If

                ctx.Response.Redirect(sRedirURL, False)
                ctx.Response.End()


            ElseIf ScreenAllowed = 97 Then 'Need EULA Acceptance
                If sAppScreenName = "EULA" Then 'do nothing, already on EULA screen
                Else
                    sRedirURL = "EULA.aspx"

                    If InStr(sRedirURL, "?") = 0 Then
                        sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If

                    ctx.Response.Redirect(sRedirURL, False)
                    ctx.Response.End()

                End If
            ElseIf ScreenAllowed = 98 Then 'User must change password
                If sAppScreenName = "ChangePasswordFirst" Then 'do nothing, already on change password screen
                    Exit Sub
                Else
                    ctx.Session("ForcePasswordChange") = True

                    sRedirURL = "AppUserChangePasswordFirst.aspx"

                    If InStr(sRedirURL, "?") = 0 Then
                        sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If

                    ctx.Response.Redirect(sRedirURL, False)
                    ctx.Response.End()
                End If
            ElseIf ScreenAllowed = 0 Then
                If sAppScreenName = "MessageHistory" Or sAppScreenName = "MyMessages" Then
                    'logout, otherwise it gets in a redirect loop
                    If sUseSSO = "1" And sUserIsProgrammer <> "1" Then

                        sRedirURL = sSSOLoginURL

                        If InStr(sRedirURL, "?") = 0 Then
                            sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        Else
                            sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        End If

                        ctx.Response.Redirect(sRedirURL, False)
                        ctx.Response.End()

                    Else

                        sRedirURL = redir & "?ErrMsg=Access Denied."

                        If InStr(sRedirURL, "?") = 0 Then
                            sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        Else
                            sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        End If

                        ctx.Response.Redirect(sRedirURL, False)
                        ctx.Response.End()

                    End If
                Else
                    If sAppScreenName = "Loginx" Then

                        sRedirURL = Trim(ctx.Session("HomePage") & "")

                        If InStr(sRedirURL, "?") = 0 Then
                            sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        Else
                            sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        End If

                        ctx.Response.Redirect(sRedirURL, False)
                        ctx.Response.End()

                    Else

                        sRedirURL = Trim(ctx.Session("HomePage") & "") & "?Error=Access Denied."

                        If InStr(sRedirURL, "?") = 0 Then
                            sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        Else
                            sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        End If

                        ctx.Response.Redirect(sRedirURL, False)
                        ctx.Response.End()

                    End If
                End If

            Else

                sRedirURL = Trim(ctx.Session("HomePage") & "") & "?Error=Unable to Check Security Access."

                If InStr(sRedirURL, "?") = 0 Then
                    sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If

                ctx.Response.Redirect(sRedirURL, False)
                ctx.Response.End()

            End If
        End If

    End Sub

    Public Shared Function CheckSystemDown() As Boolean
        Dim sSystemShutdown As String = System.Configuration.ConfigurationManager.AppSettings("SystemShutdown").ToString.Trim

        If CBool(sSystemShutdown) Then
            Return True
            Exit Function
        End If

        Dim sSystemCode As String = ""
        Try
            sSystemCode = common.GetAppControlCharacter("AMS", "MM", "SystemCode").ToString.Trim()
        Catch ex As Exception

        End Try

        If sSystemCode = "" Then
            'Database down
            CheckSystemDown = False
        Else
            If CBool(DataCalls.IsSystemShutdown(sSystemCode)) Then
                CheckSystemDown = True
            Else
                CheckSystemDown = False
            End If

        End If


    End Function

    Public Shared Function GetPageTitle(Optional ByVal sSubTitle As String = "") As String
        Dim sPageTitle As String = "Message Manager"

        If sSubTitle <> "" Then
            sPageTitle = sSubTitle & " - " & sPageTitle
        End If

        GetPageTitle = sPageTitle
    End Function

    Public Shared Function PhoneCheck(ByVal phone As String) As Boolean
        PhoneCheck = False
        '((\(\d{3}\) ?)|(\d{3}-))?\d{3}-\d{4}  -- (000) 000-0000 or 000-000-0000
        '((\(\d{3}\) ?)|(\d{3}-))?\d{3}-\d{4}( x\d{0,})? -- (000) 000-0000 x1234
        '[0-9][0-9][0-9]\-[0-9][0-9][0-9]\-[0-9][0-9][0-9][0-9] 000-000-0000 only 
        Dim pattern As String = "[0-9][0-9][0-9]\-[0-9][0-9][0-9]\-[0-9][0-9][0-9][0-9]"
        Dim phoneMatch As Match = Regex.Match(phone, pattern)
        If phoneMatch.Success Then
            PhoneCheck = True
        Else
            PhoneCheck = False
        End If

    End Function


    Public Shared Function EmailAddressCheck(ByVal emailAddress As String) As Boolean

        EmailAddressCheck = False

        Dim atPosition As Long = InStr(emailAddress, "@")
        Dim lastAtPosition As Long = InStrRev(emailAddress, "@")
        Dim lastDotPosition As Long = InStrRev(emailAddress, ".")
        Dim spacePosition As Long = InStr(emailAddress, " ")

        'If spacePosition <> 0 Then
        'EmailAddressCheck = False
        'Exit Function
        'End If

        If atPosition <> lastAtPosition Then
            EmailAddressCheck = False
            Exit Function
        End If

        If atPosition = 0 Or lastDotPosition = 0 Then
            EmailAddressCheck = False
            Exit Function
        End If

        If atPosition > lastDotPosition Then
            EmailAddressCheck = False
            Exit Function
        End If

        'If InStr(emailAddress, " ") <> 0 Then
        ' EmailAddressCheck = False
        ' Exit Function
        'End If

        EmailAddressCheck = True

        Return EmailAddressCheck

    End Function

    Public Shared Function EmailAddressCheckAppUserID(ByVal emailAddress As String) As Boolean

        'Allow only . - _ in email address

        EmailAddressCheckAppUserID = False
        Dim pattern As String = "^[a-zA-Z0-9][\w\._\-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"

        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheckAppUserID = True
        Else
            EmailAddressCheckAppUserID = False
        End If

        Return EmailAddressCheckAppUserID

    End Function

    Public Shared Sub SetError(ByVal lblError As Label, Optional ByVal sErrorText As String = "")
        lblError.Text = sErrorText
        If lblError.Text = "" Then
            lblError.Visible = False
        Else
            lblError.Visible = True
        End If

    End Sub

    Public Shared Function UserIsAdmin(ByVal sAppUserID As String) As Boolean
        Dim sRetVal As String = "0"
        Dim retValue As Boolean = False
        sRetVal = Trim(DataCalls.UserIsAdmin(sAppUserID) & "")
        If sRetVal = "" Then
            sRetVal = "0"
        End If
        retValue = CBool(sRetVal)
        Return retValue

    End Function

    Public Shared Sub EnableButton(ByVal lb As LinkButton)
        Try
            lb.Enabled = True
        Catch ex As Exception

        End Try

    End Sub

    Public Shared Sub DisableButton(ByVal lb As LinkButton)
        Try
            lb.Enabled = False
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Function GetBreadcrumbs(ByVal sAppScreen As String) As String
        Dim sAppCode As String = common.GetAppControlCharacter("AMS", "MM", "AppCode")

        Dim sBreadcrumbs As String = ""
        Dim sCrumb As String = ""
        Dim ctx = HttpContext.Current
        Dim sSecurityGroupName As String = Trim(HttpContext.Current.Session("LoggedInAppUserSecurityGroup")) & ""
        Dim sCrumbSpacer As String = "&nbsp;&rsaquo;&nbsp;"
        Dim dt As DataTable = DataCalls.GetPageBreadcrumbs(sAppCode, sSecurityGroupName, sAppScreen)
        Dim t As String = ""
        Dim u As String = ""
        Dim rc As Long = dt.Rows.Count
        Dim counter As Long = 0
        Dim lastCrumb As Boolean = False
        For Each dr In dt.Rows
            counter = counter + 1
            t = dr("BreadcrumbText")
            u = dr("BreadcrumbURL")

            If counter = rc Then
                'change text to blue
                lastCrumb = True
            Else
                lastCrumb = False
            End If

            sCrumb = BuildCrumb(t, u, lastCrumb)

            If sBreadcrumbs = "" Then
                sBreadcrumbs = sCrumb
            Else
                sBreadcrumbs = sBreadcrumbs & sCrumbSpacer & sCrumb
            End If
        Next

        If Not dt Is Nothing Then
            dt.Dispose()
            dt = Nothing
        End If

        GetBreadcrumbs = sBreadcrumbs
    End Function

    Public Shared Function BuildCrumb(ByVal t As String, ByVal u As String, ByVal lastCrumb As Boolean) As String
        Dim ctx = HttpContext.Current
        Dim sClass As String = "breadcrumbsText"
        If lastCrumb Then
            sClass = "breadcrumbsTextBlue"
        End If
        If u = "" Then
            sClass = sClass & "NoDeco"
            BuildCrumb = "<span class='" & sClass & "'>" & t & "</span>"
        Else

            If InStr(u, "?") = 0 Then
                u = u & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
            Else
                u = u & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
            End If

            BuildCrumb = "<a class='" & sClass & "' href='" & u & "'>" & t & "</a>"
        End If

    End Function

    Public Shared Function GetOptionsText() As String

        Dim sHiddenWelcomeText As String = ""
        Dim UserName As String = HttpContext.Current.Session("AppUserName")
        sHiddenWelcomeText = "<br>User: " & HttpContext.Current.Session("LoggedInAppUserID")
        If HttpContext.Current.Session("PrimaryGroup") <> "" Then
            sHiddenWelcomeText = sHiddenWelcomeText & "<br/>Primary Account: " & HttpContext.Current.Session("PrimaryGroup")
        End If
        If HttpContext.Current.Session("AppGroupID") <> "" Then
            sHiddenWelcomeText = sHiddenWelcomeText & "&nbsp;&nbsp;&#8226;&nbsp;&nbsp;Current Account: " & HttpContext.Current.Session("AppGroupID")
        End If

        GetOptionsText = sHiddenWelcomeText
    End Function

    Public Shared Function GetUserStatus(ByVal sAppUserID As String, ByVal bIncludeLabel As Boolean) As String
        Dim UserStatus As String = DataCalls.GetUserCurrentStatus(sAppUserID)
        If UserStatus <> "" And bIncludeLabel Then
            UserStatus = "<b>Status:</b> " & UserStatus
        End If
        GetUserStatus = UserStatus
    End Function

    Public Shared Function FormatAppDateTime(ByVal sDateTime As String, Optional ByVal bIncludeTime As Boolean = True, Optional ByVal bIncludeSeconds As Boolean = True) As String

        If sDateTime <> "" Then
            If Year(sDateTime) = "1900" Then
                sDateTime = ""
            Else
                If bIncludeTime Then
                    If bIncludeSeconds Then
                        sDateTime = Format(CDate(sDateTime), "MM/dd/yyyy hh:mm:sstt")
                    Else
                        sDateTime = Format(CDate(sDateTime), "MM/dd/yyyy hh:mmtt")
                    End If

                Else
                    sDateTime = Format(CDate(sDateTime), "MM/dd/yyyy")
                End If

            End If
        End If
        FormatAppDateTime = sDateTime

    End Function

    Public Shared Sub ResetCurrentlyViewing(Optional ByVal bAllClients As Boolean = False, Optional ByVal bHide As Boolean = False)
        Dim ctx = HttpContext.Current
        Dim cvString As String = ""
        Dim AppUser As String = Trim(ctx.Session("AppUserID"))
        Dim LoggedInUser As String = Trim(ctx.Session("LoggedInAppUserID"))
        Dim LoggedInUserGroup As String = Trim(ctx.Session("AppGroupID"))
        Dim UserGroup As String = Trim(LCase(DataCalls.GetUserGroup(AppUser, True)))
        Dim AppScreenName As String = Trim(ctx.Session("HelpScreen"))
        Dim LogID As String = ""

        'This may not work here because of auto refresh... come back to it after the big move on Saturday
        'Dim retval As String = DataCalls.LogViewAnotherUser(LogID, LoggedInUser, LoggedInUserGroup, AppUser, UserGroup, AppScreenName)

        If LCase(Trim(AppUser)) = LCase(Trim(LoggedInUser)) Or bHide Then
            cvString = ""
        ElseIf LCase(Trim(AppUser)) = "" Then
            cvString = ""
        ElseIf bAllClients Then
            cvString = "Currently Viewing All Clients Account: " & UserGroup
        Else
            cvString = "Currently Viewing Client: " & AppUser
        End If

        If cvString = "" Or bAllClients Then
            Dim LoggedInUserPrimaryGroup As String = DataCalls.GetUserGroup(LoggedInUser, True)
            If Trim(LCase(LoggedInUserPrimaryGroup)) <> Trim(LCase(ctx.Session("AppGroupID"))) Then
                cvString = "Currently Viewing Account: " & ctx.Session("AppGroupID")
            End If
        Else
            cvString = cvString & "&nbsp;Account: " & UserGroup
        End If
        Dim ScreenName As String = ctx.Session("HelpScreen")

        If ScreenName = "" Then
        Else
            Dim retval As String = DataCalls.LogViewAnotherUser(0, LoggedInUser, LoggedInUserGroup, AppUser, UserGroup, ScreenName)

        End If

        Dim s2 As String = "<script language='javascript'>"
        s2 = s2 & " try { document.getElementById('ctl00_cMenu_divBreadcrumbsRight').innerHTML = '" & cvString & "'; } catch (err) {} "
        s2 = s2 & " </script>"

        Dim http = Web.HttpContext.Current
        If Not http Is Nothing Then
            Dim page = TryCast(http.CurrentHandler, Web.UI.Page)
            If Not page Is Nothing Then
                Try
                    Dim sc = System.Web.UI.ScriptManager.GetCurrent(page)
                    ''sc.RegisterStartupScript(Me, Me.GetType, "savedt", s2, False)
                    ''ScriptManager.RegisterStartupScript(page, page.GetType, "savedt", s2, False)
                    ScriptManager.RegisterClientScriptBlock(page, page.GetType, "savedt", s2, False)
                Catch ex As Exception

                End Try

            End If
        End If



    End Sub

    Public Shared Function curPageURL() As String
        Dim ctx = HttpContext.Current

        Dim s As String
        Dim protocol As String
        Dim port As String

        If ctx.Request.ServerVariables("HTTPS") = "on" Then
            s = "s"
        Else
            s = ""
        End If

        protocol = strLeft(LCase(ctx.Request.ServerVariables("SERVER_PROTOCOL")), "/") & s

        If ctx.Request.ServerVariables("SERVER_PORT") = "80" Then
            port = ""
        Else
            port = ":" & ctx.Request.ServerVariables("SERVER_PORT")
        End If

        curPageURL = protocol & "://" & ctx.Request.ServerVariables("SERVER_NAME") & port & ctx.Request.ServerVariables("SCRIPT_NAME")
    End Function

    Public Shared Function strLeft(ByVal str1 As String, ByVal str2 As String) As String
        strLeft = Left(str1, InStr(str1, str2) - 1)
    End Function

    Public Shared Function curPageName() As String
        Dim ctx = HttpContext.Current

        Dim pagename As String

        pagename = ctx.Request.ServerVariables("SCRIPT_NAME")

        If InStr(pagename, "/") > 0 Then
            pagename = Right(pagename, Len(pagename) - InStrRev(pagename, "/"))
        End If

        curPageName = pagename
    End Function


    Public Shared Function ForgotLogin(ByVal sEmail As String, ByVal sFirstName As String, ByVal sLastName As String) As String
        Dim ctx = HttpContext.Current
        Dim sDevTestProd As String = common.GetAppControlCharacter("AMS", "MM", "DevTestProd")

        Dim sUseSSO As String = System.Configuration.ConfigurationManager.AppSettings("UseSSO")
        Dim retVal As String = ""
        Dim retVal2 As String = ""

        If sUseSSO = "1" Then

            Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
            Dim sURI As String = sSSO_URL & "/forgot/loginId?apiKey=" & sApiKey & "&email=" & sEmail & "&firstName=" & sFirstName & "&lastName=" & sLastName
            Dim jsonResponse = SSO.GetJSONDownloadString(sURI)
            Dim jdo As Object
            Dim sStatus As String = ""

            Try
                jdo = JsonConvert.DeserializeObject(Of SSO.SSO_ForgotLogin)(jsonResponse)
                sStatus = jdo.status
            Catch ex As Exception
                'sso server fail/down
                ForgotLogin = "Could not connect to SSO"
                Exit Function
            End Try


            If sStatus = "SUCCESS" Then
                retVal = "."
            Else
                retVal = jdo.error.message

            End If

            If sDevTestProd = "DEV" Or sDevTestProd = "TEST" Then
                'Try to Send regular password too because SSO could differ
                'retVal2 = DataCalls.ForgotLogin(sEmail)
                'If retVal2 = "." Then
                'If retVal = "." Then
                'Else
                '   retVal = retVal & "Login sent successfully."
                ' End If
                'Else
                '   If retVal = "." Then
                'retVal = "SSO- Login sent successfully.  "
                'Else
                '   retVal = retVal & ".  IntelliMessage " & retVal2
                'End If
                'End If
            End If
        Else
            retVal = DataCalls.ForgotLogin(sEmail)

        End If

        Return retVal

    End Function

    Public Shared Function ForgotPassword(ByVal sUserID As String) As String
        Dim ctx = HttpContext.Current
        Dim sDevTestProd As String = common.GetAppControlCharacter("AMS", "MM", "DevTestProd")

        Dim sUseSSO As String = System.Configuration.ConfigurationManager.AppSettings("UseSSO")
        Dim retVal As String = ""
        Dim retVal2 As String = ""

        If sUseSSO = "1" Then

            Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
            Dim sURI As String = sSSO_URL & "/forgot/password?apiKey=" & sApiKey & "&loginId=" & sUserID
            Dim jsonResponse = SSO.GetJSONDownloadString(sURI)
            Dim jdo As Object
            Dim sStatus As String = ""

            Try
                jdo = JsonConvert.DeserializeObject(Of SSO.SSO_ForgotPassword)(jsonResponse)
                sStatus = jdo.status
            Catch ex As Exception
                'sso server fail/down
                ForgotPassword = "Could not connect to SSO"
                Exit Function
            End Try


            If sStatus = "SUCCESS" Then
                retVal = "."
            Else
                retVal = jdo.error.message

            End If

            If sDevTestProd = "DEV" Or sDevTestProd = "TEST" Then
                'Try to Send regular password too because SSO could differ
                'retVal2 = DataCalls.ForgotPassword(sUserID)
                'If retVal2 = "." Then
                'If retVal = "." Then
                'Else
                '   retVal = retVal & ".  IntelliMessage Password sent successfully."
                'End If
                'Else
                '   If retVal = "." Then
                'retVal = "SSO- Password sent successfully.  "
                'Else
                '   retVal = retVal & ".  IntelliMessage " & retVal2
                'End If
                '  End If
            End If

        Else
            retVal = DataCalls.ForgotPassword(sUserID)

        End If


        Return retVal

    End Function

    Public Shared Function ChangePassword(ByVal sAppUserID As String, ByVal sNewPassword As String, ByVal sNewPassword2 As String, ByVal sCurrentPassword As String, ByVal sLoggedInAppUserID As String) As String
        Dim ctx = HttpContext.Current
        Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim retVal As String = ""
        Dim retVal2 As String = ""
        Dim SSOPasswordChangeAllowedByMM As Boolean = False

        If sUseSSO = "1" Then
            'we are skipping the change password SSO Process here because another service
            ' will now be handling it

            SSOPasswordChangeAllowedByMM = common.GetAppControlTrueFalse("AMS", "MM", "AllowChangePasswordSSO")

        End If

        If sUseSSO = "1" And SSOPasswordChangeAllowedByMM Then

            Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""

            Dim modifyUser As New SSOModifyUser()
            With modifyUser
                .firstName = ""
                .lastName = ""
                .email = ""
                .loginId = sAppUserID
                .password = sCurrentPassword
                .newPassword = sNewPassword
                .active = ""
            End With

            'product.Expiry = New DateTime(2008, 12, 28)
            'product.Price = 3.99D
            'product.Sizes = New String() {"Small", "Medium", "Large"}

            Dim json As String = JsonConvert.SerializeObject(modifyUser)

            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
            Dim sURI As String = sSSO_URL & "/subject?apiKey=" & sApiKey & "&tokenId=" & sTokenID

            Dim sysURI As New System.Uri(sURI)
            Dim data = Encoding.UTF8.GetBytes(json)
            Dim jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "PUT")
            Dim jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            Dim sStatus As String = jdo.status

            If sStatus = "SUCCESS" Then
                retVal = "."
            Else
                retVal = jdo.error.message
            End If

            If retVal = "." Then
                'SSO Password Change Successful, now change (overwrite) in IntelliMessage Database
                retVal2 = DataCalls.AdminChangePasswordAppUser2(sAppUserID, sNewPassword, sLoggedInAppUserID)
                If retVal2 = "." Then

                Else
                    If retVal = "." Then
                        retVal = "Password change successful."
                    Else
                        retVal = retVal & "IntelliMessage " & retVal2
                    End If
                End If
            End If


        Else
            retVal = DataCalls.AdminChangePasswordAppUser(sAppUserID, sNewPassword, sNewPassword2, sCurrentPassword, sLoggedInAppUserID)

        End If


        Return retVal

    End Function

    Public Shared Function SendPassword(ByVal sUserID As String, Optional ByVal sTo As String = "") As String
        Dim ctx = HttpContext.Current
        Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim retVal As String = ""
        Dim retVal2 As String = ""

        If sUseSSO = "1" Then

            Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
            Dim sURI As String = sSSO_URL
            If sTo <> "" And LCase(sTo) <> LCase(sUserID) Then
                'Only send the token ID if we are sending the User's password to the Administrator
                sURI = sURI & "/send/password?apiKey=" & sApiKey & "&loginId=" & sUserID & "&tokenId=" & sTokenID
            Else
                sURI = sURI & "/forgot/password?apiKey=" & sApiKey & "&loginId=" & sUserID
            End If
            Dim jsonResponse = SSO.GetJSONDownloadString(sURI)
            Dim jdo As Object
            Dim sStatus As String = ""

            Try
                jdo = JsonConvert.DeserializeObject(Of SSO.SSO_ForgotPassword)(jsonResponse)
                sStatus = jdo.status
            Catch ex As Exception
                'sso server fail/down
                SendPassword = "Could not connect to SSO"
                Exit Function
            End Try


            If sStatus = "SUCCESS" Then
                retVal = "."
            Else
                retVal = jdo.error.message


            End If

            'If sDevTestProd = "DEV" Or sDevTestProd = "TEST" Then
            'Try to Send regular password too because SSO could differ
            'retVal2 = DataCalls.ForgotPassword(sUserID, sTo)
            'If retVal2 = "." Then
            ' If retVal = "." Then
            'Else
            '   retVal = retVal & ".  IntelliMessage Password sent successfully."
            'End If
            'Else
            '   If retVal = "." Then
            'retVal = "SSO- Password sent successfully.  "
            'Else
            '   retVal = retVal & ".  IntelliMessage " & retVal2
            'End If
            'End If
            'End If

        Else
            retVal = DataCalls.ForgotPassword(sUserID, sTo)

        End If


        Return retVal

    End Function

    Public Shared Function CreateSSOUser(ByVal sFirstName As String, ByVal sLastName As String, ByVal sEmail As String, ByVal sClientID As String, ByVal sPassword As String) As String
        Dim ctx = HttpContext.Current
        Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim retVal As String = ""

        If sUseSSO = "1" Then

            Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")

            Dim createUser As New SSOCreateUser()
            With createUser
                .firstName = sFirstName
                .lastName = sLastName
                .email = sEmail
                .loginId = sClientID
                .password = sPassword
            End With

            Dim json As String = JsonConvert.SerializeObject(createUser)
            Dim sURI As String = sSSO_URL & "/subject?apiKey=" & sApiKey
            Dim sysURI As New System.Uri(sURI)
            Dim data = Encoding.UTF8.GetBytes(json)
            Dim jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "POST")
            Dim jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            Dim sString As String = ""
            Dim sStatus As String = jdo.status
            Dim sSSOUUID As String = ""
            Dim retVal2 As String = ""
            Dim retVal3 As String = ""
            Dim retVal4 As String = ""

            If sStatus = "SUCCESS" Then
                retVal = "."
                sSSOUUID = jdo.subject.uuid
                'Update user's ssouuid
                retVal2 = DataCalls.SaveAppUserSSOUUID(sClientID, sSSOUUID)
                If retVal2 = "." Then
                Else
                    retVal = retVal2
                End If


            Else
                retVal = jdo.error.message


            End If
        Else
            retVal = "."
        End If

        Return retVal

    End Function

    Public Shared Function AddSSORole(ByVal RoleName As String, ByVal ClientID As String) As String
        Dim ctx = HttpContext.Current
        Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim retVal As String = ""

        If sUseSSO = "1" Then

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")

            Dim createRole As New SSOCreateRole
            With createRole
                .role = RoleName
            End With

            Dim json As String = JsonConvert.SerializeObject(createRole)
            Dim sURI As String = sSSO_URL & "/subject/role?apiKey=" & sApiKey & "&tokenId=" & sTokenID & "&loginId=" & ClientID
            Dim sysURI As New System.Uri(sURI)
            Dim data = Encoding.UTF8.GetBytes(json)
            Dim jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "POST")
            Dim jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            Dim sString As String = ""
            Dim sStatus As String = jdo.status
            Dim sSSOUUID As String = ""
            If sStatus = "SUCCESS" Then
                retVal = "."
            Else
                retVal = jdo.error.message
            End If
        Else
            retVal = "."
        End If
        Return retVal
    End Function

    Public Shared Function ModifySSOUser(ByVal sFirstName As String, ByVal sLastName As String, ByVal sEmail As String, ByVal sClientID As String, ByVal sPassword As String, ByVal sNewPassword As String, ByVal sActive As String) As String
        Dim ctx = HttpContext.Current
        Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim retVal As String = "."

        If sActive = "1" Then
            sActive = "true"
        End If
        If sActive = "0" Then
            sActive = "false"
        End If

        If sUseSSO = "1" Then
            Dim sLoggedinAppUserID As String = LCase(ctx.Session("LoggedInAppUserID"))
            Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
            sClientID = LCase(sClientID)
            Dim sClientIDPreserve As String = sClientID
            Dim sURI As String = sSSO_URL & "/subject?apiKey=" & sApiKey & "&tokenId=" & sTokenID

            If sLoggedinAppUserID <> sClientID Then
                'we are changing a different user
                'Add to end of string instead of posting in class
                sURI = sURI & "&loginId=" & sClientID
                sClientID = ""
                sPassword = ""
                sNewPassword = ""
            End If

            Dim modifyUser As New SSOModifyUser()
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
            Dim jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "PUT")
            Dim jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            Dim sString As String = ""
            Dim sStatus As String = jdo.status
            Dim sSSOUUID As String = ""
            Dim retval2 As String = ""

            If sStatus = "SUCCESS" Then
                retVal = "."
                sSSOUUID = jdo.subject.uuid

                'Update user's ssouuid
                retVal2 = DataCalls.SaveAppUserSSOUUID(sClientIDPreserve, sSSOUUID)
                If retVal2 = "." Then
                Else
                    retVal = retVal2
                End If

            Else
                retVal = jdo.error.message


            End If
        Else
            retVal = "."
        End If

        Return retVal

    End Function

    Function StripDisplayName(ByVal s As String) As String

        Dim openParenPosition As Long = InStr(s, "(")
        Dim returnString As String = s
        If openParenPosition > 0 Then
            returnString = Left(s, openParenPosition - 1)
        End If
        returnString = Trim(returnString)
        returnString = Replace(returnString, ";", "")
        returnString = Replace(returnString, ",", "")
        returnString = Trim(returnString)
        Return Trim(returnString)

    End Function

    Public Shared Function GetSSOUser(ByVal sClientID As String) As String
        Dim ctx = HttpContext.Current
        Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim retVal As String = "."
        
        If sUseSSO = "1" Then

            Dim sLoggedinAppUserID As String = LCase(ctx.Session("LoggedInAppUserID"))
            Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
            sClientID = LCase(sClientID)
            Dim sURI As String = sSSO_URL & "/subject?apiKey=" & sApiKey & "&tokenId=" & sTokenID & "&loginId=" & sClientID

            Dim jsonResponse As String = SSO.GetJSONDownloadString(sURI)
            Dim jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            Dim sString As String = ""

            sString = sString & "<div>URL passed: " & sURI & "<BR><BR></div>"

            Dim sStatus As String = jdo.status

            If sStatus = "SUCCESS" Then
                Dim sActive As String = jdo.subject.active
                Dim sEmail As String = jdo.subject.email
                Dim sFirstName As String = jdo.subject.firstName
                Dim sLastName As String = jdo.subject.lastName
                Dim sSSOUUID As String = jdo.subject.uuid
                Dim RoleName As String = ""
                Dim retValRoles As String = ""
                Dim retVal3 As String = ""
                Dim retVal4 As String = ""

                'Update our database with stuff from SSO
                retVal = DataCalls.SaveAppUserFromSSO(sClientID, sEmail, sFirstName, sLastName, sActive, sSSOUUID)
                If retVal = "." Or retVal = "" Then
                    If LCase(Trim(ctx.Session("LoggedInAppUserID"))) = LCase(Trim(sClientID)) Then
                        ctx.Session("AppUserName") = sFirstName & " " & sLastName
                    End If

                    For i As Integer = 0 To jdo.subject.roles.Count - 1
                        RoleName = Trim(jdo.subject.roles.Item(i).name)
                        ctx.Session(RoleName) = True
                    Next

                    If retValRoles = "." Or retValRoles = "" Then
                    Else
                        retVal = retValRoles
                    End If

                End If
            End If

        End If

        GetSSOUser = retVal
    End Function


    Public Shared Function GetAppControlCharacter(ByVal sCompanyCode As String, ByVal sModuleCode As String, ByVal sApplicationControlCode As String) As String
        Dim sCharacterParameter As String = ""
        Dim dt As DataTable
        dt = DataCalls.GetApplicationControl(sCompanyCode, sModuleCode, sApplicationControlCode)
        If dt.Rows.Count > 0 Then
            sCharacterParameter = dt.Rows(0)("CharacterParameter").ToString()
        End If

        If Not dt Is Nothing Then
            dt.Dispose()
            dt = Nothing
        End If

        Return sCharacterParameter

    End Function

    Public Shared Function GetAppControlNumeric(ByVal sCompanyCode As String, ByVal sModuleCode As String, ByVal sApplicationControlCode As String) As Long
        Dim lNumericParameter As Long = 0
        Dim dt As DataTable
        dt = DataCalls.GetApplicationControl(sCompanyCode, sModuleCode, sApplicationControlCode)
        If dt.Rows.Count > 0 Then
            lNumericParameter = dt.Rows(0)("NumericParameter").ToString()
        End If

        If Not dt Is Nothing Then
            dt.Dispose()
            dt = Nothing
        End If

        Return lNumericParameter

    End Function

    Public Shared Function GetAppControlTrueFalse(ByVal sCompanyCode As String, ByVal sModuleCode As String, ByVal sApplicationControlCode As String) As Boolean
        Dim bTrueFalseParameter As Boolean = False
        Dim dt As DataTable
        dt = DataCalls.GetApplicationControl(sCompanyCode, sModuleCode, sApplicationControlCode)
        If dt.Rows.Count > 0 Then
            bTrueFalseParameter = CBool(dt.Rows(0)("TrueFalse").ToString())
        End If

        If Not dt Is Nothing Then
            dt.Dispose()
            dt = Nothing
        End If

        Return bTrueFalseParameter

    End Function

    Public Shared Function DeleteSSORole(ByVal RoleName As String, ByVal ClientID As String) As String
        Dim ctx = HttpContext.Current
        Dim sUseSSO As String = ctx.Session("UseSSO")
        Dim retVal As String = ""

        If sUseSSO = "1" Then

            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")

            Dim deleteRole As New SSODeleteRole()
            With deleteRole
                .role = RoleName
            End With

            Dim json As String = JsonConvert.SerializeObject(deleteRole)
            Dim sURI As String = sSSO_URL & "/subject/role?apiKey=" & sApiKey & "&tokenId=" & sTokenID & "&loginId=" & ClientID
            Dim sysURI As New System.Uri(sURI)
            Dim data = Encoding.UTF8.GetBytes(json)
            Dim jsonResponse = SSO.SendJSONRequest(sysURI, data, "application/json", "DELETE")
            Dim jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            Dim sString As String = ""
            Dim sStatus As String = jdo.status
            Dim sSSOUUID As String = ""
            If sStatus = "SUCCESS" Then
                retVal = "."
            Else
                retVal = jdo.error.message
            End If
        Else
            retVal = "."
        End If
        Return retVal
    End Function

    Public Shared Function GetUserHasRole(ByVal sUserID As String, ByVal sRoleName As String) As Boolean
        'Find out if user has role or not
        GetUserHasRole = False
        Dim ctx = HttpContext.Current

        Dim sApiKey = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
        Dim sSSOLoginURL As String = UCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))
        Dim sTokenID As String = Trim(ctx.Session("TokenID")) & ""
        Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
        Dim sURI As String = sSSO_URL & "/subject?apiKey=" & sApiKey & "&tokenId=" & ctx.Session("TokenID") & "&loginId=" & sUserID
        Dim jsonResponse = SSO.GetJSONDownloadString(sURI)
        Dim jdo As Object
        Dim sStatus As String = ""

        Try
            jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Subject)(jsonResponse)
            sStatus = jdo.status
        Catch ex As Exception
            'sso server fail/down
            GetUserHasRole = False
            Exit Function
        End Try


        If sStatus = "SUCCESS" Then
            'Successful SSO Get Subject
            For i As Integer = 0 To jdo.subject.roles.Count - 1

                If UCase(jdo.subject.roles.Item(i).name) = UCase(sRoleName) Then
                    GetUserHasRole = True
                    Exit Function
                End If

            Next
        Else
            'Failure
            GetUserHasRole = False
        End If

    End Function

    Public Shared Function FormatPhoneNumber(ByVal myNumber As String) As String
        Dim mynewNumber As String
        mynewNumber = ""
        myNumber = myNumber.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "")
        If myNumber.Length = 10 Then
            mynewNumber = myNumber.Substring(0, 3) & "-" &
                    myNumber.Substring(3, 3) & "-" & myNumber.Substring(6, 4)

        Else
            mynewNumber = myNumber
        End If
        Return mynewNumber
    End Function

    Public Shared Function GetBrowserType() As String
        Dim ctx = HttpContext.Current
        Dim br As System.Web.HttpBrowserCapabilities = ctx.Request.Browser
        Dim sBrowserUserAgent As String = Trim(ctx.Request.ServerVariables("HTTP_User_Agent"))
        Dim sBrowserType As String = br.Browser

        If sBrowserType = "Mozilla" And InStr(sBrowserUserAgent, "Trident/") <> 0 Then
            sBrowserType = "IE"
        End If

        GetBrowserType = sBrowserType

    End Function

    Public Shared Function GetSystemDays(ByVal sLoggedInAppUserID As String) As Long

        Dim dt As DataTable
        dt = DataCalls.GetMasterSystemControl(sLoggedInAppUserID)
        If dt.Rows.Count > 0 Then
            GetSystemDays = dt.Rows(0)("ForcePasswordDays").ToString()
        Else
            GetSystemDays = 0
        End If

        If Not dt Is Nothing Then
            dt.Dispose()
            dt = Nothing
        End If

        Return GetSystemDays

    End Function

    Public Shared Function GetComputerName() As String
        Dim ctx = HttpContext.Current
        Dim sComputerName As String = ""


        'Try
        'sComputerName = System.Net.Dns.GetHostEntry(ctx.Request.UserHostAddress).HostName
        'Catch ex As Exception

        'End Try

        If sComputerName = "" Then
            Try
                sComputerName = ctx.Request.UserHostAddress
            Catch ex As Exception

            End Try

        End If

        GetComputerName = sComputerName
    End Function

    Public Shared Function GetUserStatus2(ByVal sAppUserID As String, ByVal bIncludeLabel As Boolean, ByVal bIncludeForwardTo As Boolean) As String
        Dim UserStatus As String = DataCalls.GetUserCurrentStatus(sAppUserID, bIncludeForwardTo)
        If UserStatus <> "" And bIncludeLabel Then
            UserStatus = "<b>Status:</b> " & UserStatus
        End If
        GetUserStatus2 = UserStatus
    End Function


    Public Shared Function LoginUserDual(Optional ByVal sUserID As String = "", Optional ByVal sPassword As String = "") As String
        'This will attempt to login the user via SSO first
        'If unsuccessful, it will attempt login via our database
        Dim sAppCode As String = common.GetAppControlCharacter("AMS", "MM", "AppCode")
        Dim sUserIsProgrammer As String = System.Configuration.ConfigurationManager.AppSettings("UserIsProgrammer")
        Dim sDevTestProd As String = common.GetAppControlCharacter("AMS", "MM", "DevTestProd")

        LoginUserDual = ""
        Dim ctx = HttpContext.Current
        ctx.Session("AppCode") = common.GetAppControlCharacter("AMS", "MM", "AppCode")
        Dim dt As DataTable

        Dim sReturnMsg As String = ""
        Dim sRedirectURL As String = ""
        Dim sRedirPath As String = ""
        Dim sCurrentPageName As String = LCase(curPageName)
        Dim queryFromPage As String = Trim(ctx.Request.QueryString("fromPage") & "")

        sRedirPath = Trim(ctx.Request.QueryString("redirPath")) & ""

        'ctx.Response.Write("sCurrentPageName: " & sCurrentPageName)
        'ctx.Response.End()

        'Check login credentials                     
        ctx.Session("LoggedInAtLeastOnce") = False 'If 0 then this is their first login to the site
        'ctx.Response.Write("<br>UseSSO: " & ctx.Session("UseSSO"))
        'ctx.Response.Write("<br>SSOUUID: " & ctx.Session("SSOUUID"))
        'ctx.Response.Write("<br>sDevTestProd: " & sDevTestProd)
        'ctx.Response.End()

        ctx.Session("UseSSO") = "1"
        'If ctx.Session("UseSSO") = "" Then
        ' ctx.Session("UseSSO") = System.Configuration.ConfigurationManager.AppSettings("UseSSO")
        'End If

        'ctx.Response.Write("token: " & ctx.Session("TokenID"))

        '        ctx.Response.Write("<br>SSOUUID: " & ctx.Session("SSOUUID"))
        '       ctx.Response.End()

        If ctx.Session("TokenID") <> "" And ctx.Session("SSOUUID") <> "" Then
            'Came in with a valid token and has already gone through CheckSecurity
            dt = DataCalls.LoginSSOUser(sAppCode, LCase(sUserID), ctx.Session("SSOUUID"))

        Else
            Dim sApiKey As String = common.GetAppControlCharacter("AMS", "MM", "SSO_ApiKey")
            Dim sSSO_URL As String = common.GetAppControlCharacter("AMS", "MM", "SSO_URL")
            Dim sURI As String = sSSO_URL & "/authenticate?apiKey=" & sApiKey & "&loginId=" & sUserID & "&password=" & HttpUtility.UrlEncode(sPassword)
            Dim jsonResponse = SSO.GetJSONDownloadString(sURI)
            Dim jdo As Object
            Dim sStatus As String = ""

            Try
                jdo = JsonConvert.DeserializeObject(Of SSO.SSO_Authenticate)(jsonResponse)
                sStatus = jdo.status
            Catch ex As Exception
                'sso server fail/down
                common.LoginUser(sUserID, sPassword)
                Exit Function
            End Try


            If sUserID = "xreyna@intellimsg.net" Then
                ctx.Response.Write(sURI)
                ctx.Response.Write("<br>DateTime: " & Now.ToString & "<Br>")
                'Response.End()

            End If




            If sStatus = "SUCCESS" Then
                ctx.Session("TokenID") = jdo.token.tokenId
                ctx.Session("SSOUUID") = jdo.subject.uuid

                dt = DataCalls.LoginSSOUser(sAppCode, LCase(sUserID), ctx.Session("SSOUUID"))

                If sUserID = "xreyna@intellimsg.net" Then
                    ctx.Response.Write("AdminLoginSSOUser '" & sAppCode & "', '" & LCase(sUserID) & "', '" & ctx.Session("SSOUUID") & "'")
                    ctx.Response.Write("<br>DateTime: " & Now.ToString & "<br>")
                    ' ctx.Response.End()

                End If
                'ctx.Response.Write("xLoginSSOUser '" & sAppCode & "', '" & LCase(sUserID) & "', '" & ctx.Session("SSOUUID") & "'")
                'ctx.Response.End()
                'maybe this... will see
              
            Else
                'ELSE if that didn't work, run regular login code
                'ctx.Response.Write("<br>xxx: ")
                'ctx.Response.End()
                ctx.Session("UseSSO") = "0"
                dt = DataCalls.LoginAppUser(sAppCode, LCase(sUserID), sPassword, ctx.Session("fromMPA"))
                'ctx.Response.Write("yLoginAppUser '" & sAppCode & "', '" & LCase(sUserID) & "', '" & sPassword & "', '" & ctx.Session("fromMPA") & "'")

            End If

        End If

        If sUserID = "xreyna@intellimsg.net" Then
            ctx.Response.Write("<br>zDateTime: " & Now.ToString & "<br>")
        End If
        If dt.Rows.Count > 0 Then
            sReturnMsg = dt.Rows(0)("ReturnMsg").ToString()
            If sReturnMsg = "." Then

                ctx.Session("AppUserID") = LCase(dt.Rows(0)("AppUserID").ToString())
                ctx.Session("LoggedInAppUserID") = LCase(dt.Rows(0)("AppUserID").ToString())
                ctx.Session("AppGroupID") = LCase(dt.Rows(0)("AppGroupID").ToString())
                ctx.Session("AppGroupIDDisplay") = dt.Rows(0)("AppGroupID").ToString()
                ctx.Session("PrimaryGroup") = LCase(dt.Rows(0)("AppGroupID").ToString())
                ctx.Session("PrimaryGroupDisplay") = dt.Rows(0)("AppGroupID").ToString()
                ctx.Session("PrimaryEmail") = dt.Rows(0)("PrimaryEmail").ToString()
                ctx.Session("AppUserName") = dt.Rows(0)("AppUserName").ToString()
                ctx.Session("TimeZoneOffsetHours") = dt.Rows(0)("TimeZoneOffsetHours").ToString()
                ctx.Session("IT") = CBool(dt.Rows(0)("IT").ToString())
                ctx.Session("LoggedInAtLeastOnce") = dt.Rows(0)("LoggedInAtLeastOnce").ToString()
                ctx.Session("EULA") = dt.Rows(0)("EULAAccepted").ToString()
                ctx.Session("UserType") = dt.Rows(0)("UserType").ToString()
                If sUserID = "xreyna@intellimsg.net" Then
                    ctx.Response.Write("<br>xDateTime: " & Now.ToString & "<br>")
                End If
            End If
        Else
            sReturnMsg = "Invalid Client ID and/or Password."

        End If

        If Not dt Is Nothing Then
            dt.Dispose()
            dt = Nothing
        End If

        If sReturnMsg = "." Then

            ctx.Session("AppUserSecurityGroup") = DataCalls.GetUserSecurityGroup(ctx.Session("LoggedInAppUserID"))
            ctx.Session("LoggedInAppUserSecurityGroup") = ctx.Session("AppUserSecurityGroup")
            ctx.Session("ShowMainSelectedTabForSubMenu") = True
            ctx.Session("AppUserSecurityGroupIT") = ctx.Session("IT")
            If Trim(ctx.Session("TimeZoneOffsetHours")) & "" <> "" Then
                ctx.Session("TimeZoneOffsetHours") = CLng(ctx.Session("TimeZoneOffsetHours"))
            Else
                ctx.Session("TimeZoneOffsetHours") = 0
            End If

            If Trim(ctx.Session("IT")) & "" <> "" Then
                ctx.Session("IT") = CBool(ctx.Session("IT"))
            Else
                ctx.Session("IT") = False
            End If

            If Trim(ctx.Session("LoggedInAtLeastOnce")) & "" <> "" Then
                ctx.Session("LoggedInAtLeastOnce") = CBool(ctx.Session("LoggedInAtLeastOnce"))
            Else
                ctx.Session("LoggedInAtLeastOnce") = False
            End If

            If Trim(ctx.Session("EULA")) & "" <> "" Then
                ctx.Session("EULA") = CBool(ctx.Session("EULA"))
            Else
                ctx.Session("EULA") = False
            End If

            If ctx.Session("UserType") = "Group" Then
                ctx.Session("AppUserSecurityGroup") = "User"
                ctx.Session("LoggedInAppUserSecurityGroup") = "User"
                ctx.Session("EULA") = True
            End If

            'Log login into database!
            Dim LogID As String = Trim(ctx.Session("LogID")) & ""
            'ctx.Response.Write("LoggedInAppUserID: " & ctx.Session("LoggedInAppUserID"))
            'ctx.Response.Write("<br>LogID: " & LogID)
            'ctx.Response.End()
            If LogID = "" Or LogID = "0" Then
                'They probably established a session and are coming back in from Account Manager
                Dim IPAddress As String = Trim(ctx.Request.ServerVariables("Remote_addr"))
                If sUserID = "xreyna@intellimsg.net" Then
                    ctx.Response.Write("<br>before browser type DateTime: " & Now.ToString & "<br>")
                End If
                Dim BrowserType As String = common.GetBrowserType()
                If sUserID = "xreyna@intellimsg.net" Then
                    ctx.Response.Write("<br>after browser type DateTime: " & Now.ToString & "<br>")
                End If
                Dim BrowserUserAgent As String = Trim(ctx.Request.ServerVariables("HTTP_User_Agent"))
                Dim AspNetSessionID As String = System.Web.HttpContext.Current.Session.SessionID
                If sUserID = "xreyna@intellimsg.net" Then
                    ctx.Response.Write("<br>before computer name DateTime: " & Now.ToString & "<br>")
                End If
                Dim ComputerName As String = common.GetComputerName
                If sUserID = "xreyna@intellimsg.net" Then
                    ctx.Response.Write("<br>after computer name DateTime: " & Now.ToString & "<br>")
                End If
                Dim dtl As DataTable
                Dim retVal As String = ""

                dtl = DataCalls.LogUserSessionLogin2(ctx.Session("LoggedInAppUserID"), ctx.Session("AppUserID"), ctx.Session("AppGroupID"), "PCM", IPAddress, BrowserType, BrowserUserAgent, ctx.Session("UseSSO"), ComputerName, AspNetSessionID)
                If sUserID = "xreyna@intellimsg.net" Then
                    ctx.Response.Write("<br>yDateTime: " & Now.ToString & "<br>")
                End If
                If dtl.Rows.Count > 0 Then
                    retVal = dtl.Rows(0)("ReturnMsg").ToString()
                    LogID = dtl.Rows(0)("LogID").ToString()
                Else
                    LogID = "0"
                End If

                If Not dtl Is Nothing Then
                    dtl.Dispose()
                    dtl = Nothing
                End If

                ctx.Session("LogID") = LogID
                ctx.Session("ComputerName") = ComputerName
            End If

            'ctx.Response.Write("LoggedInAppUserID: " & ctx.Session("LoggedInAppUserID"))
            'ctx.Response.Write("<br>LogID: " & LogID)
            'ctx.Response.Write("<br>UseSSO: " & ctx.Session("UseSSO"))
            'ctx.Response.End()
            ctx.Session("ShowMainSelectedTabForSubMenu") = ctx.Session("AppCode")


            If sCurrentPageName = "default.aspx" Or sCurrentPageName = "loginx.aspx" Then
                If ctx.Session("LoggedInAppUserSecurityGroup") = "Basic" Then
                    ctx.Session("HomePage") = "SendMessageBasic.aspx"
                Else
                    Dim sGetLastSavedCurrentPage As String = UserPreferences.GetCharacterParameter(ctx.Session("LoggedInAppUserID"), "SY", "CurrentPage")
                    If sGetLastSavedCurrentPage <> "" Then
                        ctx.Session("HomePage") = sGetLastSavedCurrentPage
                    Else
                        ctx.Session("HomePage") = "MyMessages.aspx"

                    End If
                   
                End If

            ElseIf ctx.Session("HomePage") <> "" Then
                'leave it
            Else
                ctx.Session("HomePage") = sCurrentPageName

            End If

            'ctx.Response.Write("home page: " & ctx.Session("HomePage"))
            'ctx.Response.End()

            If sRedirPath = "" Then
                Dim sGetLastSavedCurrentPage As String = UserPreferences.GetCharacterParameter(ctx.Session("LoggedInAppUserID"), "SY", "CurrentPage")
                sRedirPath = sGetLastSavedCurrentPage
            End If

            If sUserIsProgrammer = "1" Then
                'Do not redirect because I am programming on my local machine!
            Else
                Dim sBetaPath As String = System.Configuration.ConfigurationManager.AppSettings("BetaPath")
                If ctx.Session("AppUserSecurityGroup") = "User" Or ctx.Session("AppUserSecurityGroup") = "Basic" Then
                    sRedirectURL = common.GetAppControlCharacter("AMS", "MM", "MyAccountRedirectURL")
                Else
                    sRedirectURL = common.GetAppControlCharacter("AMS", "MM", "MessageManagerRedirectURL")
                End If
                sRedirectURL = sRedirectURL & sBetaPath
            End If

            sRedirectURL = Trim(sRedirectURL) & ""

            If sRedirectURL <> "" Then
                If Right(sRedirectURL, 1) <> "/" Then
                    sRedirectURL = sRedirectURL & "/"
                End If
            End If

            If ctx.Session("EULA") = False Then

                sRedirectURL = sRedirectURL & "EULA.aspx"

                If InStr(sRedirectURL, "?") = 0 Then
                    sRedirectURL = sRedirectURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedirectURL = sRedirectURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If
                If queryFromPage <> "" Then
                    sRedirectURL = sRedirectURL & "&fromPage=" & queryFromPage
                End If

                'ctx.Response.Write("1z" & sRedirectURL)
                'ctx.Response.End()

                ctx.Response.Redirect(sRedirectURL, False)
                ctx.Response.End()

            ElseIf ctx.Session("LoggedInAtLeastOnce") = True Then
                'ctx.Response.Write("HERE sRedirectURL: " & sRedirectURL)
                'ctx.Response.Write("<br> sRedirPath: " & sRedirPath)
                ' ctx.Response.End()
                If sRedirPath <> "" Then

                    sRedirectURL = sRedirectURL & sRedirPath
                    '   ctx.Response.Write("<br>1r: " & sRedirectURL)

                    If InStr(sRedirectURL, "?") = 0 Then
                        sRedirectURL = sRedirectURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirectURL = sRedirectURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If
                    'ctx.Response.Write("<br>2r: " & sRedirectURL)

                    If queryFromPage <> "" Then
                        sRedirectURL = sRedirectURL & "&fromPage=" & queryFromPage
                    End If
                    'ctx.Response.Write("<br>3r: " & sRedirectURL)

                    If 1 = 2 Or sUserID = "xreyna@intellimsg.net" Then
                        ctx.Response.Write(sRedirectURL)
                        ctx.Response.Write("<br>1DateTime: " & Now.ToString & "<Br>")
                        ctx.Response.End()

                    End If
                    'ctx.Response.Write("2z" & sRedirectURL)
                    'ctx.Response.End()

                    ctx.Response.Redirect(sRedirectURL, False)
                    ctx.Response.End()

                Else

                    sRedirectURL = sRedirectURL & Trim(ctx.Session("HomePage") & "")

                    If InStr(sRedirectURL, "?") = 0 Then
                        sRedirectURL = sRedirectURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirectURL = sRedirectURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If
                    If queryFromPage <> "" Then
                        sRedirectURL = sRedirectURL & "&fromPage=" & queryFromPage
                    End If

                    If sUserID = "xreyna@intellimsg.net" Then
                        ctx.Response.Write(sRedirectURL)
                        ctx.Response.Write("<br>2DateTime: " & Now.ToString & "<Br>")
                        ctx.Response.End()


                    End If
                    'ctx.Response.Write("3z" & sRedirectURL)
                    'ctx.Response.End()

                    ctx.Response.Redirect(sRedirectURL, False)
                    ctx.Response.End()

                End If
            Else


                sRedirectURL = sRedirectURL & Trim(ctx.Session("HomePage") & "")

                If InStr(sRedirectURL, "?") = 0 Then
                    sRedirectURL = sRedirectURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedirectURL = sRedirectURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If
                If queryFromPage <> "" Then
                    sRedirectURL = sRedirectURL & "&fromPage=" & queryFromPage
                End If

                'ctx.Response.Write("4z" & sRedirectURL)
                'ctx.Response.End()


                ctx.Response.Redirect(sRedirectURL, False)
                ctx.Response.End()

            End If
        End If

        LoginUserDual = sReturnMsg

    End Function


    Public Shared Sub CheckSecurityDual(ByVal sAppScreenName As String, Optional ByVal bScreenIsPopup As Boolean = False, Optional ByVal bPopupIsModal As Boolean = False)
        Dim ctx = HttpContext.Current
        Dim sTokenQstring As String = Trim(ctx.Request("TokenID")) & ""
        Dim sLogID As String = Trim(ctx.Session("LogID") & "")
        Dim sLoginHasBeenProcessedAlready As Boolean = True
        Dim retval As String = ""

        If sLogID = "" Then
            sLogID = "0"
        End If



        Dim sCurrentPageName As String = curPageName()
        If sCurrentPageName = "" Then
            sCurrentPageName = "."
        Else
            'save last page they were on to use the next time they log on
            'do not save if they are coming back from Account Manager because it will just
            'send them to my messages
            If sTokenQstring = "" And Trim(ctx.Session("LoggedInAppUserID")) & "" <> "" Then
                ''Not putting in production yet - comment out line below if putting in prod
                Dim retSavePageName As String = DataCalls.SaveUserPreferences(ctx.Session("LoggedInAppUserID"), "SY", "CurrentPage", "", True, sCurrentPageName, 0)

            End If

        End If
        
        Dim sQstring As String = Trim(ctx.Request.ServerVariables("QUERY_STRING")) & ""
        Dim TheCurrentDateAndTime As DateTime = Now()
        If sLogID <> "0" Then
            '   ctx.Response.Write("<br>made it here<br>")
            retval = DataCalls.LogUserSessionDetail(sLogID, sCurrentPageName, sQstring)
        End If

        'ctx.Response.Write("LogID: " & sLogID & "; sCurrentPageName: " & sCurrentPageName & "; sQstring: " & sQstring)
        'ctx.Response.End()
        If sTokenQstring <> "" Then 'coming from sso with a token, must re-login And sLogID <> "0" Then

            sLoginHasBeenProcessedAlready = False
            If ctx.Session("TokenID") <> sTokenQstring Then
                'Different User
                ctx.Session("DateTimeToValidateToken") = TheCurrentDateAndTime
                ctx.Session("DateTimeToRenewToken") = TheCurrentDateAndTime
                ctx.Session("LoggedInAppUserID") = ""
                ctx.Session("UseSSO") = ""
                ctx.Session("INTELLIMSG_USER") = False
            End If


            If sTokenQstring = "" Then
                'ok to log, but skip if they are coming directly from SSO because they will never be passing a LOGID in the string
                If Len(sQstring) <> 0 Then
                    Dim sLogIDQString As String = Trim(ctx.Request.QueryString("LogID")) & ""
                    If sLogIDQString = "" Then
                        sLogIDQString = "0"
                    End If
                    If sLogID <> sLogIDQString Then
                        Dim retval2 As String = DataCalls.LogUserSessionLogIDMismatch(sLogIDQString, sLogID, sCurrentPageName, sQstring)
                    End If
                End If
            End If


        End If

        Dim sAppCode As String = common.GetAppControlCharacter("AMS", "MM", "AppCode")
        Dim sUserIsProgrammer As String = System.Configuration.ConfigurationManager.AppSettings("UserIsProgrammer")
        Dim sDevTestProd As String = common.GetAppControlCharacter("AMS", "MM", "DevTestProd")

        ctx.Session("AppCode") = sAppCode

        Dim sAppUserID As String = ctx.Session("LoggedInAppUserID")
        Dim ScreenAllowed As Long = 0
        Dim isSystemDown As Boolean = CheckSystemDown()
        Dim sCurrentPath As String = ""
        Dim sUseSSO As String = Trim(ctx.Session("UseSSO")) & "" 'System.Configuration.ConfigurationManager.AppSettings("UseSSO")
        Dim sSSOLoginURL As String = LCase(common.GetAppControlCharacter("AMS", "MM", "SSOLoginURL"))
        Dim redir As String = "Default.aspx"
        Dim sReturnMsg As String = ""

        If sUserIsProgrammer = "1" Then
            sSSOLoginURL = "Default.aspx"
        End If

        If sUseSSO = "" Then
            sUseSSO = "1"
            ctx.Session("UseSSO") = "1"
            sLoginHasBeenProcessedAlready = False

        End If



        If sUseSSO = "1" Then

            Dim bValidToken As Boolean = common.TokenValidate()

            If bValidToken Then
                'Valid token, continue
                TokenRenew() 'new
                'ctx.Response.Write("get user info: " & ctx.Session("GetUserInfo"))
                'ctx.Response.End()
                If ctx.Session("GetUserInfo") = False Or ctx.Session("INTELLIMSG_USER") = False Then
                    'ctx.Response.Write("here2")
                    ctx.Session("HomePage") = ""
                    Dim t As String = ctx.Session("TokenID")
                    ' ctx.Session.RemoveAll()
                    ctx.Session("TokenID") = t
                    'ctx.Response.End()
                    If SetUserInfo() = True Then
                        'Can use message manager


                        ' ctx.Response.Write("sLoginHasBeenProcessedAlready: " & sLoginHasBeenProcessedAlready)
                        'ctx.Response.Write("<br>ssoUUID: " & ctx.Session("SSOUUID"))

                        'ctx.Response.End()

                        If sLoginHasBeenProcessedAlready Then
                            Exit Sub
                            'why are we doing this again?
                        Else

                            sReturnMsg = common.LoginUserDual(ctx.Session("UserID"), "")

                            'ctx.Response.Write("sReturnMsg: " & sReturnMsg)
                            'ctx.Response.End()

                            If sReturnMsg <> "" And sReturnMsg <> "." Then

                                sSSOLoginURL = sSSOLoginURL & "?ErrMsg=" & sReturnMsg

                                If InStr(sSSOLoginURL, "?") = 0 Then
                                    sSSOLoginURL = sSSOLoginURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                                Else
                                    sSSOLoginURL = sSSOLoginURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                                End If

                                ' ctx.Response.Write(sSSOLoginURL)
                                ' ctx.Response.End()

                                ctx.Response.Redirect(sSSOLoginURL, False)
                                ctx.Response.End()

                            End If

                        End If

                    Else

                        sSSOLoginURL = sSSOLoginURL & "?ErrMsg=Could Not get user info."

                        If InStr(sSSOLoginURL, "?") = 0 Then
                            sSSOLoginURL = sSSOLoginURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        Else
                            sSSOLoginURL = sSSOLoginURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                        End If
                        '  ctx.Response.Write(sSSOLoginURL)
                        '  ctx.Response.End()
                        ctx.Response.Redirect(sSSOLoginURL, False)
                        ctx.Response.End()

                    End If
                Else
                    ' ctx.Response.Write("here")
                    'ctx.Response.End()
                    TokenRenew()
                End If

            Else
                'Invalid token
                sCurrentPath = IO.Path.GetFileName(ctx.Request.PhysicalPath)

                Dim sRedURL As String = "Logout.aspx?redirPath=" & sCurrentPath & "&p=" & bScreenIsPopup & "&m=" & bPopupIsModal

                If InStr(sRedURL, "?") = 0 Then
                    sRedURL = sRedURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedURL = sRedURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If
                ' ctx.Response.Write(sRedURL)
                ' ctx.Response.End()
                ctx.Response.Redirect(sRedURL, False)
                ctx.Response.End()

            End If


        End If



        If isSystemDown Then
            If ctx.Session("AppUserSecurityGroup") = "Admin" And ctx.Session("IT") = True Then
                'Allowed to use system while shutdown
            Else

                Dim sShutURL As String = "SystemShutdown.aspx"

                If InStr(sShutURL, "?") = 0 Then
                    sShutURL = sShutURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sShutURL = sShutURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If

                ' ctx.Response.Write(sShutURL)
                ' ctx.Response.End()
                ctx.Response.Redirect(sShutURL, False)
                ctx.Response.End()

            End If
        End If


        'ctx.Response.Write("sAppUserID: " & sAppUserID & "<br>")


        If sAppUserID = "" Then


            'ctx.Response.Write("useSSO: " & sUseSSO)

            'ctx.Response.End()
            If sUseSSO = "1" Then
                If InStr(sSSOLoginURL, "?referrer=") = 0 Then
                    sSSOLoginURL = LCase(sSSOLoginURL) & "?referrer=" & ctx.Request.Url.Scheme + "://" + ctx.Request.Url.Authority + ctx.Request.ApplicationPath & IO.Path.GetFileName(ctx.Request.PhysicalPath)

                    'ctx.Response.Write("sSSOLoginURL: " & sSSOLoginURL)
                    'ctx.Response.End()

                    'sSSOLoginURL = sSSOLoginURL & "&ErrMsg=Invalid Token-a."
                End If

                If InStr(sSSOLoginURL, "?") = 0 Then
                    sSSOLoginURL = sSSOLoginURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sSSOLoginURL = sSSOLoginURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If
                'ctx.Response.Write(sSSOLoginURL)
                'ctx.Response.End()
                ctx.Response.Redirect(sSSOLoginURL, False)
                ctx.Response.End()

            Else


                sCurrentPath = IO.Path.GetFileName(ctx.Request.PhysicalPath)

                Dim sLUrl As String = "Logout.aspx?redirPath=" & sCurrentPath & "&p=" & bScreenIsPopup & "&m=" & bPopupIsModal

                If InStr(sLUrl, "?") = 0 Then
                    sLUrl = sLUrl & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sLUrl = sLUrl & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If
                ' ctx.Response.Write(sLUrl)
                ' ctx.Response.End()
                ctx.Response.Redirect(sLUrl, False)
                ctx.Response.End()


            End If
            Exit Sub
        End If

        Dim UserGroup As String = Trim(LCase(DataCalls.GetUserGroup(ctx.Session("AppUserID"), True)))
        Dim AppScreenName As String = Trim(ctx.Session("HelpScreen"))

        If AppScreenName = "" Then
        Else
            Dim retvallog As String = DataCalls.LogViewAnotherUser(0, ctx.Session("LoggedInAppUserID"), ctx.Session("AppGroupID"), ctx.Session("AppUserID"), UserGroup, AppScreenName)

        End If

        Dim sRedirURL As String = ""

        ScreenAllowed = DataCalls.CheckSecurity(sAppUserID, sAppScreenName)
        If ScreenAllowed = 1 Then
            'Ok to view
            'Log the change

        ElseIf ScreenAllowed = 99 Then 'Inactive user or groups, log out
            sRedirURL = "Logout.aspx?p=" & bScreenIsPopup & "&m=" & bPopupIsModal

            If InStr(sRedirURL, "?") = 0 Then
                sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
            Else
                sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
            End If

            ' ctx.Response.Write(sRedirURL)
            ' ctx.Response.End()

            ctx.Response.Redirect(sRedirURL, False)
            ctx.Response.End()


        ElseIf ScreenAllowed = 97 Then 'Need EULA Acceptance
            If sAppScreenName = "EULA" Then 'do nothing, already on EULA screen
            Else
                sRedirURL = "EULA.aspx"

                If InStr(sRedirURL, "?") = 0 Then
                    sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If

                '  ctx.Response.Write(sRedirURL)
                '  ctx.Response.End()



                ctx.Response.Redirect(sRedirURL, False)
                ctx.Response.End()

            End If
        ElseIf ScreenAllowed = 98 Then 'User must change password
            If sAppScreenName = "ChangePasswordFirst" Then 'do nothing, already on change password screen
                Exit Sub
            Else
                ctx.Session("ForcePasswordChange") = True

                sRedirURL = "AppUserChangePasswordFirst.aspx"

                If InStr(sRedirURL, "?") = 0 Then
                    sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                Else
                    sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                End If

                '  ctx.Response.Write(sRedirURL)
                '  ctx.Response.End()



                ctx.Response.Redirect(sRedirURL, False)
                ctx.Response.End()
            End If
        ElseIf ScreenAllowed = 0 Then
            If sAppScreenName = "MessageHistory" Or sAppScreenName = "MyMessages" Then
                'logout, otherwise it gets in a redirect loop
                If sUseSSO = "1" And sUserIsProgrammer <> "1" Then

                    sRedirURL = sSSOLoginURL

                    If InStr(sRedirURL, "?") = 0 Then
                        sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If

                    'ctx.Response.Write(sRedirURL)
                    'ctx.Response.End()



                    ctx.Response.Redirect(sRedirURL, False)
                    ctx.Response.End()

                Else

                    sRedirURL = redir & "?ErrMsg=Access Denied."

                    If InStr(sRedirURL, "?") = 0 Then
                        sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If
                    ' ctx.Response.Write(sRedirURL)
                    ' ctx.Response.End()


                    ctx.Response.Redirect(sRedirURL, False)
                    ctx.Response.End()

                End If
            Else
                If sAppScreenName = "Loginx" Then

                    sRedirURL = Trim(ctx.Session("HomePage") & "")

                    If InStr(sRedirURL, "?") = 0 Then
                        sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If
                    'ctx.Response.Write(sRedirURL)
                    'ctx.Response.End()


                    ctx.Response.Redirect(sRedirURL, False)
                    ctx.Response.End()

                Else

                    sRedirURL = Trim(ctx.Session("HomePage") & "") & "?Error=Access Denied."

                    If InStr(sRedirURL, "?") = 0 Then
                        sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    Else
                        sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
                    End If
                    'ctx.Response.Write(sRedirURL)
                    'ctx.Response.End()


                    ctx.Response.Redirect(sRedirURL, False)
                    ctx.Response.End()

                End If
            End If

        Else

            sRedirURL = Trim(ctx.Session("HomePage") & "") & "?Error=Unable to Check Security Access."

            If InStr(sRedirURL, "?") = 0 Then
                sRedirURL = sRedirURL & "?q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
            Else
                sRedirURL = sRedirURL & "&q1=" & Date.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt") & "&LogID=" & ctx.Session("LogID") & "&cn=" & ctx.Session("ComputerName")
            End If

            'ctx.Response.Write(sRedirURL)
            'ctx.Response.End()
            ctx.Response.Redirect(sRedirURL, False)
            ctx.Response.End()

        End If

    End Sub
End Class
