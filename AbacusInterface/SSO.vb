Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Net
Imports System.IO

Public Class SSOModifyUser

    Property firstName As String

    Property lastName As String

    Property email As String

    Property loginId As String

    Property password As String

    Property newPassword As String

    Property active As String

End Class

Public Class SSODeleteRole

    Property role As String

End Class

Public Class SSOCreateRole

    Property role As String

End Class

Public Class SSOCreateUser

    Property firstName As String

    Property lastName As String

    Property email As String

    Property loginId As String

    Property password As String

    Property phoneNumber As String

End Class

'Public Class SSOAddPrincipal

'    Property apiKey As String

'    Property token As String

'End Class
Public Class SSOAddPrincipal

    Property principal As String

    Property uuid As String

    Property appUuid As String

End Class

Public Class SSO

    Public Shared Function GetJSONDownloadString(sUri As String) As String
        Dim jsonResponse As String = ""

        Try
            Dim wClient = New WebClient()
            Try
                jsonResponse = wClient.DownloadString(New Uri(sUri))
            Catch wex As System.Net.WebException
                If wex.Response IsNot Nothing Then
                    Using errorResponse = DirectCast(wex.Response, HttpWebResponse)
                        Using reader = New StreamReader(errorResponse.GetResponseStream())
                            'TODO: use JSON.net to parse this string and look at the error message
                            jsonResponse = reader.ReadToEnd()
                        End Using
                    End Using
                End If
            End Try
        Catch wex As System.Net.WebException
            'resume
        End Try
        Return jsonResponse
    End Function

    Public Shared Function SendJSONRequest(uri As Uri, jsonDataBytes As Byte(), contentType As String, method As String, Optional header As String = "") As String
        Dim req As WebRequest = WebRequest.Create(uri)
        req.ContentType = contentType
        req.Method = method
        'req.ServicePoint.Expect100Continue = false;

        req.ContentLength = jsonDataBytes.Length
        If header.Trim <> "" Then
            req.Headers.Add(header)
        End If

        req.Timeout = 15000

        Dim stream = req.GetRequestStream()
        stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
        stream.Close()

        Dim res As String = ""

        Try
            Using response As System.IO.Stream = req.GetResponse().GetResponseStream()
                Using reader = New StreamReader(response)
                    res = reader.ReadToEnd
                End Using
            End Using

        Catch wex As WebException

            If wex.Response IsNot Nothing Then
                Using errorResponse = DirectCast(wex.Response, HttpWebResponse)
                    Using errorReader = New StreamReader(errorResponse.GetResponseStream())
                        'Use JSON.net to parse this string and look at the error message
                        res = errorReader.ReadToEnd()
                    End Using
                End Using
            End If

        End Try

        If res = "" Then
            res = "Waited for 15 seconds but process timed out. {Most likely Proxy Error}"
        End If
        Return res
    End Function

#Region "ForgotLogin"
    Public Class SSO_ForgotLogin

        Public Property requestTimestamp() As String
            Get
                Return m_requestTimestamp
            End Get
            Set(value As String)
                m_requestTimestamp = value
            End Set
        End Property
        Private m_requestTimestamp As String

        Public Property status() As String
            Get
                Return m_status
            End Get
            Set(value As String)
                m_status = value
            End Set
        End Property
        Private m_status As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property [error]() As ErrorInfo
            Get
                Return m_error
            End Get
            Set(value As ErrorInfo)
                m_error = value
            End Set
        End Property
        Private m_error As ErrorInfo
    End Class
#End Region

#Region "ForgotPassword"
    Public Class SSO_ForgotPassword

        Public Property requestTimestamp() As String
            Get
                Return m_requestTimestamp
            End Get
            Set(value As String)
                m_requestTimestamp = value
            End Set
        End Property
        Private m_requestTimestamp As String

        Public Property status() As String
            Get
                Return m_status
            End Get
            Set(value As String)
                m_status = value
            End Set
        End Property
        Private m_status As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property [error]() As ErrorInfo
            Get
                Return m_error
            End Get
            Set(value As ErrorInfo)
                m_error = value
            End Set
        End Property
        Private m_error As ErrorInfo
    End Class
#End Region

#Region "Subject" 'create/get user

    Public Class SSO_Subject

        Public Property subject() As SubjectInfo
            Get
                Return m_subject
            End Get
            Set(value As SubjectInfo)
                m_subject = value
            End Set
        End Property
        Private m_subject As SubjectInfo

        Public Property requestTimestamp() As String
            Get
                Return m_requestTimestamp
            End Get
            Set(value As String)
                m_requestTimestamp = value
            End Set
        End Property
        Private m_requestTimestamp As String

        Public Property status() As String
            Get
                Return m_status
            End Get
            Set(value As String)
                m_status = value
            End Set
        End Property
        Private m_status As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property [error]() As ErrorInfo
            Get
                Return m_error
            End Get
            Set(value As ErrorInfo)
                m_error = value
            End Set
        End Property
        Private m_error As ErrorInfo
    End Class
#End Region

#Region "Authorize"

    Public Class SSO_Authorize

        Public Property authorization() As List(Of UriAuthorization)
            Get
                Return m_authorization
            End Get
            Set(value As List(Of UriAuthorization))
                m_authorization = value
            End Set
        End Property
        Private m_authorization As List(Of UriAuthorization)

        Public Property requestTimestamp() As String
            Get
                Return m_requestTimestamp
            End Get
            Set(value As String)
                m_requestTimestamp = value
            End Set
        End Property
        Private m_requestTimestamp As String

        Public Property status() As String
            Get
                Return m_status
            End Get
            Set(value As String)
                m_status = value
            End Set
        End Property
        Private m_status As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property [error]() As ErrorInfo
            Get
                Return m_error
            End Get
            Set(value As ErrorInfo)
                m_error = value
            End Set
        End Property
        Private m_error As ErrorInfo
    End Class

    Public Class UriAuthorization 'stuff in authorization square bracket
        Public Property uri() As String
            Get
                Return m_uri
            End Get
            Set(value As String)
                m_uri = value
            End Set
        End Property
        Private m_uri As String

        Public Property authorized() As String
            Get
                Return m_authorized
            End Get
            Set(value As String)
                m_authorized = value
            End Set
        End Property
        Private m_authorized As String
    End Class

#End Region


#Region "TokenRenew"
    Public Class SSO_TokenRenew

        Public Property token() As TokenInfo
            Get
                Return m_token
            End Get
            Set(value As TokenInfo)
                m_token = value
            End Set
        End Property
        Private m_token As TokenInfo

        Public Property requestTimestamp() As String
            Get
                Return m_requestTimestamp
            End Get
            Set(value As String)
                m_requestTimestamp = value
            End Set
        End Property
        Private m_requestTimestamp As String

        Public Property status() As String
            Get
                Return m_status
            End Get
            Set(value As String)
                m_status = value
            End Set
        End Property
        Private m_status As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property [error]() As ErrorInfo
            Get
                Return m_error
            End Get
            Set(value As ErrorInfo)
                m_error = value
            End Set
        End Property
        Private m_error As ErrorInfo
    End Class
#End Region

#Region "Token"
    Public Class SSO_Token
        Public Property token() As TokenInfo
            Get
                Return m_token
            End Get
            Set(value As TokenInfo)
                m_token = value
            End Set
        End Property
        Private m_token As TokenInfo

        Public Property requestTimestamp() As String
            Get
                Return m_requestTimestamp
            End Get
            Set(value As String)
                m_requestTimestamp = value
            End Set
        End Property
        Private m_requestTimestamp As String

        Public Property status() As String
            Get
                Return m_status
            End Get
            Set(value As String)
                m_status = value
            End Set
        End Property
        Private m_status As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property [error]() As ErrorInfo
            Get
                Return m_error
            End Get
            Set(value As ErrorInfo)
                m_error = value
            End Set
        End Property
        Private m_error As ErrorInfo
    End Class
#End Region

#Region "Logout"
    Public Class SSO_Logout

        Public Property requestTimestamp() As String
            Get
                Return m_requestTimestamp
            End Get
            Set(value As String)
                m_requestTimestamp = value
            End Set
        End Property
        Private m_requestTimestamp As String

        Public Property status() As String
            Get
                Return m_status
            End Get
            Set(value As String)
                m_status = value
            End Set
        End Property
        Private m_status As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property [error]() As ErrorInfo
            Get
                Return m_error
            End Get
            Set(value As ErrorInfo)
                m_error = value
            End Set
        End Property
        Private m_error As ErrorInfo
    End Class
#End Region

#Region "Authenticate"

    Public Class SSO_Authenticate
        Public Property token() As TokenInfo
            Get
                Return m_token
            End Get
            Set(value As TokenInfo)
                m_token = value
            End Set
        End Property
        Private m_token As TokenInfo

        Public Property subject() As SubjectInfo
            Get
                Return m_subject
            End Get
            Set(value As SubjectInfo)
                m_subject = value
            End Set
        End Property
        Private m_subject As SubjectInfo

        Public Property requestTimestamp() As String
            Get
                Return m_requestTimestamp
            End Get
            Set(value As String)
                m_requestTimestamp = value
            End Set
        End Property
        Private m_requestTimestamp As String

        Public Property status() As String
            Get
                Return m_status
            End Get
            Set(value As String)
                m_status = value
            End Set
        End Property
        Private m_status As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property [error]() As ErrorInfo
            Get
                Return m_error
            End Get
            Set(value As ErrorInfo)
                m_error = value
            End Set
        End Property
        Private m_error As ErrorInfo
    End Class

    Public Class TokenInfo 'token:
        Public Property tokenId() As String
            Get
                Return m_tokenId
            End Get
            Set(value As String)
                m_tokenId = value
            End Set
        End Property
        Private m_tokenId As String

        Public Property expiresDate() As String
            Get
                Return m_expiresDate
            End Get
            Set(value As String)
                m_expiresDate = value
            End Set
        End Property
        Private m_expiresDate As String

        Public Property expired() As Boolean
            Get
                Return m_expired
            End Get
            Set(value As Boolean)
                m_expired = value
            End Set
        End Property
        Private m_expired As Boolean

        Public Property expiresTimestamp() As String
            Get
                Return m_expiresTimestamp
            End Get
            Set(value As String)
                m_expiresTimestamp = value
            End Set
        End Property
        Private m_expiresTimestamp As String


    End Class

    Public Class SubjectInfo 'subject:
        Public Property principal() As String
            Get
                Return m_principal
            End Get
            Set(value As String)
                m_principal = value
            End Set
        End Property
        Private m_principal As String

        Public Property email() As String
            Get
                Return m_email
            End Get
            Set(value As String)
                m_email = value
            End Set
        End Property
        Private m_email As String

        Public Property firstName() As String
            Get
                Return m_firstName
            End Get
            Set(value As String)
                m_firstName = value
            End Set
        End Property
        Private m_firstName As String

        Public Property lastName() As String
            Get
                Return m_lastName
            End Get
            Set(value As String)
                m_lastName = value
            End Set
        End Property
        Private m_lastName As String

        Public Property active() As String
            Get
                Return m_active
            End Get
            Set(value As String)
                m_active = value
            End Set
        End Property
        Private m_active As String

        Public Property uuid() As String
            Get
                Return m_uuid
            End Get
            Set(value As String)
                m_uuid = value
            End Set
        End Property
        Private m_uuid As String

        Public Property api() As String
            Get
                Return m_api
            End Get
            Set(value As String)
                m_api = value
            End Set
        End Property
        Private m_api As String

        Public Property permissions() As List(Of UserPermissions)
            Get
                Return m_permissions
            End Get
            Set(value As List(Of UserPermissions))
                m_permissions = value
            End Set
        End Property
        Private m_permissions As List(Of UserPermissions)

        Public Property roles() As List(Of UserRoles)
            Get
                Return m_roles
            End Get
            Set(value As List(Of UserRoles))
                m_roles = value
            End Set
        End Property
        Private m_roles As List(Of UserRoles)

        Public Property alternateIds() As List(Of UseralternateIds)
            Get
                Return m_alternateIds
            End Get
            Set(value As List(Of UseralternateIds))
                m_alternateIds = value
            End Set
        End Property
        Private m_alternateIds As List(Of UseralternateIds)
    End Class

    Public Class UserPermissions 'stuff in permissions square bracket
        Public Property uri() As String
            Get
                Return m_uri
            End Get
            Set(value As String)
                m_uri = value
            End Set
        End Property
        Private m_uri As String
    End Class

    Public Class UserRoles 'stuff in roles square bracket
        Public Property name() As String
            Get
                Return m_name
            End Get
            Set(value As String)
                m_name = value
            End Set
        End Property
        Private m_name As String
    End Class

    Public Class UseralternateIds
        Public Property alternateId() As String
            Get
                Return m_alternateId
            End Get
            Set(value As String)
                m_alternateId = value
            End Set
        End Property
        Private m_alternateId As String
    End Class



    Public Class ErrorInfo 'error:
        Public Property message() As String
            Get
                Return m_message
            End Get
            Set(value As String)
                m_message = value
            End Set
        End Property
        Private m_message As String

        Public Property code() As String
            Get
                Return m_code
            End Get
            Set(value As String)
                m_code = value
            End Set
        End Property
        Private m_code As String
    End Class

#End Region

End Class
