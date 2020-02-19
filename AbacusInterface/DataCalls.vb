Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class DataCalls

    Public Shared Function BindLoadToUsers(ByVal AppUserId As String, ByVal AppGroupId As String, Optional ByVal bActiveOnly As Boolean = False) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSendAMessageToUsers"
            .Parameters.AddWithValue("@AppUserId", AppUserId)
            .Parameters.AddWithValue("@AppGroupId", AppGroupId)
            .Parameters.AddWithValue("@ActiveOnly", bActiveOnly)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadToUsers = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAppUsers(Optional ByVal sAppUserId As String = "", Optional ByVal bAdminOnly As Boolean = False, Optional ByVal sAppGroupID As String = "", Optional ByVal sSearchTerm As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppUsers"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AdminOnly", bAdminOnly)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAppUsers = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAppCodes() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppCodes"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAppCodes = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadContentGroups(ByVal sAppGroupID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminContentGroups"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadContentGroups = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadDeviceTypes() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeviceTypes"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadDeviceTypes = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadDeviceTypesHelp() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeviceTypesHelp"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadDeviceTypesHelp = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadPushService() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminPushService"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadPushService = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAppGroups(ByVal sAppUserId As String, Optional ByVal bActiveOnly As Boolean = False) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppGroupsSelect"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@ActiveOnly", bActiveOnly)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAppGroups = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAppGroups2(ByVal sAppUserId As String, Optional ByVal bActiveOnly As Boolean = False, Optional ByVal sSearchTerm As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppGroupsSelect2"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@ActiveOnly", bActiveOnly)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAppGroups2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAppUserDevice(ByVal sAppCode As String, ByVal sAppUserId As String, ByVal sDeviceId As String, ByVal sDeviceTypeCode As String, sUserDeviceDescription As String, sAppVersion As String, sAppInstalledDateTime As String, bLoggedIn As Boolean, ByVal bActive As Boolean, sDevicePhoneNumber As String, sDeviceCapCode As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUserDevice"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@DeviceId", sDeviceId)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@UserDeviceDescription", sUserDeviceDescription)
            .Parameters.AddWithValue("@AppVersion", sAppVersion)
            .Parameters.AddWithValue("@AppInstalledDateTime", sAppInstalledDateTime)
            .Parameters.AddWithValue("@LoggedIn", bLoggedIn)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@DevicePhoneNumber", sDevicePhoneNumber)
            .Parameters.AddWithValue("@DeviceCapCode", sDeviceCapCode)
            .Parameters.AddWithValue("@ReturnMessage", "")
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

        End Try

        Return retValue

    End Function

    Public Shared Function CheckAppUser(ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCheckAppUser"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
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

        End Try

        Return retValue
    End Function

    Public Shared Function ClearMessage(ByVal sAppUserID As String, ByVal sAppCode As String, ByVal sMessageNumber As String, ByVal sLoggedInUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminClearMessage"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@MessageNumber", sMessageNumber)
            .Parameters.AddWithValue("@LoggedInUserID", sLoggedInUserID)
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

        End Try

        Return retValue
    End Function

    Public Shared Function SaveAppUser(ByVal sActionType As String, ByVal sAppUserID As String, ByVal sName As String, ByVal sPrimaryEmail As String, ByVal sReceiptNotificationEmail As String, ByVal sReplyNotificationEmail As String, ByVal sPagerNumber As String, ByVal sPrimaryCell As String, ByVal sCellCarrierCode As String, ByVal sTimeZoneCode As String, ByVal bActive As Boolean, ByVal bNoAutoCreateReply As Boolean, ByVal sFirstName As String, ByVal sLastName As String, ByVal sUserType As String, ByVal sBillRateCode As String, ByVal sCoCode As String, ByVal sAcctNo As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, Optional ByVal sPrimaryGroup As String = "", Optional ByVal sForwardEmail1 As String = "", Optional ByVal sForwardEmail2 As String = "", Optional ByVal sAppCode As String = "", Optional ByVal sLoggedInAppUser As String = "", Optional ByVal sPagerTypeCode As String = "", Optional ByVal bClearMessageByDevice As Boolean = False) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUser"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@ReceiptNotificationEmail", sReceiptNotificationEmail)
            .Parameters.AddWithValue("@ReplyNotificationEmail", sReplyNotificationEmail)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@CellCarrierCode", sCellCarrierCode)
            .Parameters.AddWithValue("@TimeZoneCode", sTimeZoneCode)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@NoAutoCreateReply", bNoAutoCreateReply)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@UserType", sUserType)
            .Parameters.AddWithValue("@BillRateCode", sBillRateCode)
            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@PrimaryGroup", sPrimaryGroup)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@ForwardEmail1", sForwardEmail1)
            .Parameters.AddWithValue("@ForwardEmail2", sForwardEmail2)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUser)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerTypeCode)
            .Parameters.AddWithValue("@ClearMessageByDevice", bClearMessageByDevice)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAppUserEmergencyAlerts(ByVal sAppUserID As String, ByVal sEmergencyAlertLabel1 As String, ByVal sEmergencyAlertLabel2 As String, ByVal sEmergencyAlertLabel3 As String, ByVal sEmergencyAlertNumber1 As String, ByVal sEmergencyAlertNumber2 As String, ByVal sEmergencyAlertNumber3 As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUserEmergencyAlerts"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@EmergencyAlertLabel1", sEmergencyAlertLabel1)
            .Parameters.AddWithValue("@EmergencyAlertLabel2", sEmergencyAlertLabel2)
            .Parameters.AddWithValue("@EmergencyAlertLabel3", sEmergencyAlertLabel3)
            .Parameters.AddWithValue("@EmergencyAlertNumber1", sEmergencyAlertNumber1)
            .Parameters.AddWithValue("@EmergencyAlertNumber2", sEmergencyAlertNumber2)
            .Parameters.AddWithValue("@EmergencyAlertNumber3", sEmergencyAlertNumber3)
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

        End Try

        Return retValue

    End Function

    Public Shared Function AdminChangePasswordAppUser(ByVal sAppUserID As String, ByVal sNewPassword As String, ByVal sNewPassword2 As String, ByVal sCurrentPassword As String, ByVal sLoggedInUser As String) As String
        Dim retValue As String
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

        If Not cmdSQL Is Nothing Then
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End If

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAppUserGroup(ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal bIsChecked As Boolean, ByVal bIsPrimary As Boolean, ByVal bIsAdmin As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUserGroup"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@isChecked", bIsChecked)
            .Parameters.AddWithValue("@isPrimary", bIsPrimary)
            .Parameters.AddWithValue("@isAdmin", bIsAdmin)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAppUserGroup2(ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal bIsChecked As Boolean, ByVal bIsPrimary As Boolean, ByVal bIsAdmin As Boolean, ByVal bIsOperator As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUserGroup2"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@isChecked", bIsChecked)
            .Parameters.AddWithValue("@isPrimary", bIsPrimary)
            .Parameters.AddWithValue("@isAdmin", bIsAdmin)
            .Parameters.AddWithValue("@isOperator", bIsOperator)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetAppUser(ByVal sAppUserID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUser"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAppUser = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function LoginAppUser(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sPassword As String, ByVal sFromMPA As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLoginAppUser"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@Password", sPassword)
            .Parameters.AddWithValue("@FromMPA", sFromMPA)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        LoginAppUser = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function LoginSSOUser(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sSSOUUID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLoginSSOUser"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@SSOUUID", sSSOUUID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        LoginSSOUser = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function LogoutAppUser(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String, ByVal sDeviceTypeCode As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "Logout"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        LogoutAppUser = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetMasterProcessControl(ByVal sAppUserID As String, ByVal sProcessDescription As String, ByVal bShowTopOne As Boolean) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterProcessControl"
            .Parameters.AddWithValue("@AppUserID", sAppUserId)
            .Parameters.AddWithValue("@ProcessDescription", sProcessDescription)
            .Parameters.AddWithValue("@ShowTopOne", bShowTopOne)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterProcessControl = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveMasterProcessControl(ByVal sActionType As String, ByVal sProcessDescription As String, ByVal bProcessRunning As Boolean, ByVal sProcessBegin As String, ByVal sProcessEnd As String, ByVal lProcessRecordsProcessed As Long, ByVal lProcessWaitFor As Long, ByVal lProcessTerminateEmptyLoops As Long, ByVal bProcessShouldTerminate As Boolean, ByVal lDailyRecordsProcessed As Long, ByVal lWeeklyRecordsProcessed As Long, ByVal lHeartbeatRecordFrequency As Long, ByVal sHeartbeatDateTime As String, ByVal lHeartbeatSecondsAutoRestart As Long, ByVal lHeartbeatWarningAfterSeconds As Long, ByVal lHeartbeatEmailsToSend As Long, ByVal lHeartbeatEmailsSent As Long, ByVal sHeartbeatEmailAddress As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterProcessControl"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@ProcessDescription", sProcessDescription)
            .Parameters.AddWithValue("@ProcessRunning", bProcessRunning)
            .Parameters.AddWithValue("@ProcessBegin", sProcessBegin)
            .Parameters.AddWithValue("@ProcessEnd", sProcessEnd)
            .Parameters.AddWithValue("@ProcessRecordsProcessed", lProcessRecordsProcessed)
            .Parameters.AddWithValue("@ProcessWaitFor", lProcessWaitFor)
            .Parameters.AddWithValue("@ProcessTerminateEmptyLoops", lProcessTerminateEmptyLoops)
            .Parameters.AddWithValue("@ProcessShouldTerminate", bProcessShouldTerminate)
            .Parameters.AddWithValue("@DailyRecordsProcessed", lDailyRecordsProcessed)
            .Parameters.AddWithValue("@WeeklyRecordsProcessed", lWeeklyRecordsProcessed)
            .Parameters.AddWithValue("@HeartbeatRecordFrequency", lHeartbeatRecordFrequency)
            .Parameters.AddWithValue("@HeartbeatDateTime", sHeartbeatDateTime)
            .Parameters.AddWithValue("@HeartbeatSecondsAutoRestart", lHeartbeatSecondsAutoRestart)
            .Parameters.AddWithValue("@HeartbeatWarningAfterSeconds", lHeartbeatWarningAfterSeconds)
            .Parameters.AddWithValue("@HeartbeatEmailsToSend", lHeartbeatEmailsToSend)
            .Parameters.AddWithValue("@HeartbeatEmailsSent", lHeartbeatEmailsSent)
            .Parameters.AddWithValue("@HeartbeatEmailAddress", sHeartbeatEmailAddress)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetMasterSystemControl(ByVal sAppUserID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterSystemControl"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterSystemControl = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveMasterSystemControl(ByVal sAppUserID As String, ByVal sSystemCode As String, ByVal sSystemDescription As String, ByVal sCompanyName As String, ByVal sCompanyAddress1 As String, ByVal sCompanyAddress2 As String, ByVal sCompanyCity As String, ByVal sCompanyState As String, ByVal sCompanyZip As String, ByVal sPrimaryContactName As String, ByVal sPrimaryPhone As String, ByVal sPrimaryCell As String, ByVal sPrimaryEmail As String, ByVal sErrorEmail As String, ByVal lLastMessageNumber As Long, ByVal bSystemShutdown As Boolean, ByVal sSystemShutdownDate As String, ByVal sSystemShutdownBy As String, ByVal sSystemShutdownReason As String, ByVal lPushInterval As Long, ByVal lPollInterval As Long, ByVal sNextWebServiceURL As String, ByVal lInfoURLInterval As Long, ByVal sOutboundEmailMethod As String, ByVal lLastDeviceId As Long, ByVal sSMTPKeepAliveFromAddress As String, ByVal sPushWebServiceURl As String, ByVal sWebPagesURL As String, ByVal lNotDeliveredAfterSeconds As Long, ByVal lNotReadAfterSeconds As Long, ByVal sCannedQuestion1 As String, ByVal sCannedQuestion2 As String, ByVal sCannedQuestion3 As String, ByVal sCannedQuestion4 As String, ByVal sCannedQuestion5 As String, ByVal sCannedQuestion6 As String, ByVal lLastSubscriberId As Long) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterSystemControl"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@SystemCode", sSystemCode)
            .Parameters.AddWithValue("@SystemDescription", sSystemDescription)
            .Parameters.AddWithValue("@CompanyName", sCompanyName)
            .Parameters.AddWithValue("@CompanyAddress1", sCompanyAddress1)
            .Parameters.AddWithValue("@CompanyAddress2", sCompanyAddress2)
            .Parameters.AddWithValue("@CompanyCity", sCompanyCity)
            .Parameters.AddWithValue("@CompanyState", sCompanyState)
            .Parameters.AddWithValue("@CompanyZip", sCompanyZip)
            .Parameters.AddWithValue("@PrimaryContactName", sPrimaryContactName)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@ErrorEmail", sErrorEmail)
            .Parameters.AddWithValue("@LastMessageNumber", lLastMessageNumber)
            .Parameters.AddWithValue("@SystemShutdown", bSystemShutdown)
            .Parameters.AddWithValue("@SystemShutdownDate", sSystemShutdownDate)
            .Parameters.AddWithValue("@SystemShutdownBy", sSystemShutdownBy)
            .Parameters.AddWithValue("@SystemShutdownReason", sSystemShutdownReason)
            .Parameters.AddWithValue("@PushInterval", lPushInterval)
            .Parameters.AddWithValue("@PollInterval", lPollInterval)
            .Parameters.AddWithValue("@NextWebServiceURL", sNextWebServiceURL)
            .Parameters.AddWithValue("@InfoURLInterval", lInfoURLInterval)
            .Parameters.AddWithValue("@OutboundEmailMethod", sOutboundEmailMethod)
            .Parameters.AddWithValue("@LastDeviceId", lLastDeviceId)
            .Parameters.AddWithValue("@SMTPKeepAliveFromAddress", sSMTPKeepAliveFromAddress)
            .Parameters.AddWithValue("@PushWebServiceURL", sPushWebServiceURl)
            .Parameters.AddWithValue("@WebPagesURL", sWebPagesURL)
            .Parameters.AddWithValue("@NotDeliveredAfterSeconds", lNotDeliveredAfterSeconds)
            .Parameters.AddWithValue("@NotReadAfterSeconds", lNotReadAfterSeconds)
            .Parameters.AddWithValue("@CannedQuestion1", sCannedQuestion1)
            .Parameters.AddWithValue("@CannedQuestion2", sCannedQuestion2)
            .Parameters.AddWithValue("@CannedQuestion3", sCannedQuestion3)
            .Parameters.AddWithValue("@CannedQuestion4", sCannedQuestion4)
            .Parameters.AddWithValue("@CannedQuestion5", sCannedQuestion5)
            .Parameters.AddWithValue("@CannedQuestion6", sCannedQuestion6)
            .Parameters.AddWithValue("@LastSubscriberId", lLastSubscriberId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetMasterAppGroup(ByVal sAppGroupID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterAppGroup"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterAppGroup = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveMasterAppGroup(ByVal sActionType As String, ByVal sAppGroupID As String, ByVal sAppGroupDescription As String, ByVal sName As String, ByVal sPrimaryContactName As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sPrimaryCell As String, ByVal sPrimaryEmail As String, ByVal sEmergencyEmail As String, ByVal sEmergencyAlertLabel1 As String, ByVal sEmergencyAlertLabel2 As String, ByVal sEmergencyAlertLabel3 As String, ByVal sEmergencyAlertNumber1 As String, ByVal sEmergencyAlertNumber2 As String, ByVal sEmergencyAlertNumber3 As String, ByVal lNotDeliveredAfterSeconds As Long, ByVal lNotReadAfterSeconds As Long, ByVal sCannedQuestion1 As String, ByVal sCannedQuestion2 As String, ByVal sCannedQuestion3 As String, ByVal sCannedQuestion4 As String, ByVal sCannedQuestion5 As String, ByVal sCannedQuestion6 As String, ByVal sCannedReply1 As String, ByVal sCannedReply2 As String, ByVal sCannedReply3 As String, ByVal sCannedReply4 As String, ByVal sCannedReply5 As String, ByVal sCannedReply6 As String, ByVal sReceiptNotificationEmail As String, ByVal sReplyNotificationEmail As String, ByVal lInfoURLInterval As Long, ByVal bChangeHeaderImage As Boolean, ByVal bChangeFooterImage As Boolean, ByVal bActive As Boolean, ByVal bAutoCreateReply As Boolean, ByVal sDomainName As String, ByVal sCoCode As String, ByVal sAcctNo As String, ByVal sTimeZoneCode As String, ByVal bInternalGroup As Boolean, ByVal bGroupMessaging As Boolean, ByVal bOnlyAllowMessagesFromWithinGroup As Boolean, ByVal bAllowMessageFromPrimaryEmail As Boolean, ByVal bAllowReplyOnGroupMessage As Boolean, ByVal lMessageRetentionDays As Long, ByVal sGroupMessageFromIndividualOrGroup As String, ByVal sAppUserIdFormatCode As String) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterAppGroup"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppGroupDescription", sAppGroupDescription)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@PrimaryContactName", sPrimaryContactName)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@EmergencyEmailAddress", sEmergencyEmail)
            .Parameters.AddWithValue("@EmergencyAlertLabel1", sEmergencyAlertLabel1)
            .Parameters.AddWithValue("@EmergencyAlertLabel2", sEmergencyAlertLabel2)
            .Parameters.AddWithValue("@EmergencyAlertLabel3", sEmergencyAlertLabel3)
            .Parameters.AddWithValue("@EmergencyAlertNumber1", sEmergencyAlertNumber1)
            .Parameters.AddWithValue("@EmergencyAlertNumber2", sEmergencyAlertNumber2)
            .Parameters.AddWithValue("@EmergencyAlertNumber3", sEmergencyAlertNumber3)
            .Parameters.AddWithValue("@NotDeliveredAfterSeconds", lNotDeliveredAfterSeconds)
            .Parameters.AddWithValue("@NotReadAfterSeconds", lNotReadAfterSeconds)
            .Parameters.AddWithValue("@CannedQuestion1", sCannedQuestion1)
            .Parameters.AddWithValue("@CannedQuestion2", sCannedQuestion2)
            .Parameters.AddWithValue("@CannedQuestion3", sCannedQuestion3)
            .Parameters.AddWithValue("@CannedQuestion4", sCannedQuestion4)
            .Parameters.AddWithValue("@CannedQuestion5", sCannedQuestion5)
            .Parameters.AddWithValue("@CannedQuestion6", sCannedQuestion6)
            .Parameters.AddWithValue("@CannedReply1Email", sCannedReply1)
            .Parameters.AddWithValue("@CannedReply2Email", sCannedReply2)
            .Parameters.AddWithValue("@CannedReply3Email", sCannedReply3)
            .Parameters.AddWithValue("@CannedReply4Email", sCannedReply4)
            .Parameters.AddWithValue("@CannedReply5Email", sCannedReply5)
            .Parameters.AddWithValue("@CannedReply6Email", sCannedReply6)
            .Parameters.AddWithValue("@ReceiptNotificationEmail", sReceiptNotificationEmail)
            .Parameters.AddWithValue("@ReplyNotificationEmail", sReplyNotificationEmail)
            .Parameters.AddWithValue("@InfoURLInterval", lInfoURLInterval)
            .Parameters.AddWithValue("@ChangeHeaderImage", bChangeHeaderImage)
            .Parameters.AddWithValue("@ChangeFooterImage", bChangeFooterImage)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@AutoCreateReply", bAutoCreateReply)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@TimeZoneCode", sTimeZoneCode)
            .Parameters.AddWithValue("@InternalGroup", bInternalGroup)
            .Parameters.AddWithValue("@GroupMessaging", bGroupMessaging)
            .Parameters.AddWithValue("@OnlyAllowMessagesFromWithinGroup", bOnlyAllowMessagesFromWithinGroup)
            .Parameters.AddWithValue("@AllowMessageFromPrimaryEmail", bAllowMessageFromPrimaryEmail)
            .Parameters.AddWithValue("@AllowReplyOnGroupMessage", bAllowReplyOnGroupMessage)
            .Parameters.AddWithValue("@MessageRetentionDays", lMessageRetentionDays)
            .Parameters.AddWithValue("@GroupMessageFromIndividualOrGroup", sGroupMessageFromIndividualOrGroup)
            .Parameters.AddWithValue("@AppUserIdFormatCode", sAppUserIdFormatCode)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveMasterAppGroup2(ByVal sActionType As String, ByVal sAppGroupID As String, ByVal sAppGroupDescription As String, ByVal sName As String, ByVal sPrimaryContactName As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sPrimaryCell As String, ByVal sPrimaryEmail As String, ByVal sEmergencyEmail As String, ByVal sEmergencyAlertLabel1 As String, ByVal sEmergencyAlertLabel2 As String, ByVal sEmergencyAlertLabel3 As String, ByVal sEmergencyAlertNumber1 As String, ByVal sEmergencyAlertNumber2 As String, ByVal sEmergencyAlertNumber3 As String, ByVal lNotDeliveredAfterSeconds As Long, ByVal lNotReadAfterSeconds As Long, ByVal sCannedQuestion1 As String, ByVal sCannedQuestion2 As String, ByVal sCannedQuestion3 As String, ByVal sCannedQuestion4 As String, ByVal sCannedQuestion5 As String, ByVal sCannedQuestion6 As String, ByVal sCannedReply1 As String, ByVal sCannedReply2 As String, ByVal sCannedReply3 As String, ByVal sCannedReply4 As String, ByVal sCannedReply5 As String, ByVal sCannedReply6 As String, ByVal sReceiptNotificationEmail As String, ByVal sReplyNotificationEmail As String, ByVal lInfoURLInterval As Long, ByVal bChangeHeaderImage As Boolean, ByVal bChangeFooterImage As Boolean, ByVal bActive As Boolean, ByVal bAutoCreateReply As Boolean, ByVal sDomainName As String, ByVal sCoCode As String, ByVal sAcctNo As String, ByVal sTimeZoneCode As String, ByVal bInternalGroup As Boolean, ByVal bGroupMessaging As Boolean, ByVal bOnlyAllowMessagesFromWithinGroup As Boolean, ByVal bAllowMessageFromPrimaryEmail As Boolean, ByVal bAllowReplyOnGroupMessage As Boolean, ByVal lMessageRetentionDays As Long, ByVal sGroupMessageFromIndividualOrGroup As String, ByVal sAppUserIdFormatCode As String, ByVal bSendMessageDetail As Boolean) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterAppGroup2"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppGroupDescription", sAppGroupDescription)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@PrimaryContactName", sPrimaryContactName)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@EmergencyEmailAddress", sEmergencyEmail)
            .Parameters.AddWithValue("@EmergencyAlertLabel1", sEmergencyAlertLabel1)
            .Parameters.AddWithValue("@EmergencyAlertLabel2", sEmergencyAlertLabel2)
            .Parameters.AddWithValue("@EmergencyAlertLabel3", sEmergencyAlertLabel3)
            .Parameters.AddWithValue("@EmergencyAlertNumber1", sEmergencyAlertNumber1)
            .Parameters.AddWithValue("@EmergencyAlertNumber2", sEmergencyAlertNumber2)
            .Parameters.AddWithValue("@EmergencyAlertNumber3", sEmergencyAlertNumber3)
            .Parameters.AddWithValue("@NotDeliveredAfterSeconds", lNotDeliveredAfterSeconds)
            .Parameters.AddWithValue("@NotReadAfterSeconds", lNotReadAfterSeconds)
            .Parameters.AddWithValue("@CannedQuestion1", sCannedQuestion1)
            .Parameters.AddWithValue("@CannedQuestion2", sCannedQuestion2)
            .Parameters.AddWithValue("@CannedQuestion3", sCannedQuestion3)
            .Parameters.AddWithValue("@CannedQuestion4", sCannedQuestion4)
            .Parameters.AddWithValue("@CannedQuestion5", sCannedQuestion5)
            .Parameters.AddWithValue("@CannedQuestion6", sCannedQuestion6)
            .Parameters.AddWithValue("@CannedReply1Email", sCannedReply1)
            .Parameters.AddWithValue("@CannedReply2Email", sCannedReply2)
            .Parameters.AddWithValue("@CannedReply3Email", sCannedReply3)
            .Parameters.AddWithValue("@CannedReply4Email", sCannedReply4)
            .Parameters.AddWithValue("@CannedReply5Email", sCannedReply5)
            .Parameters.AddWithValue("@CannedReply6Email", sCannedReply6)
            .Parameters.AddWithValue("@ReceiptNotificationEmail", sReceiptNotificationEmail)
            .Parameters.AddWithValue("@ReplyNotificationEmail", sReplyNotificationEmail)
            .Parameters.AddWithValue("@InfoURLInterval", lInfoURLInterval)
            .Parameters.AddWithValue("@ChangeHeaderImage", bChangeHeaderImage)
            .Parameters.AddWithValue("@ChangeFooterImage", bChangeFooterImage)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@AutoCreateReply", bAutoCreateReply)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@TimeZoneCode", sTimeZoneCode)
            .Parameters.AddWithValue("@InternalGroup", bInternalGroup)
            .Parameters.AddWithValue("@GroupMessaging", bGroupMessaging)
            .Parameters.AddWithValue("@OnlyAllowMessagesFromWithinGroup", bOnlyAllowMessagesFromWithinGroup)
            .Parameters.AddWithValue("@AllowMessageFromPrimaryEmail", bAllowMessageFromPrimaryEmail)
            .Parameters.AddWithValue("@AllowReplyOnGroupMessage", bAllowReplyOnGroupMessage)
            .Parameters.AddWithValue("@MessageRetentionDays", lMessageRetentionDays)
            .Parameters.AddWithValue("@GroupMessageFromIndividualOrGroup", sGroupMessageFromIndividualOrGroup)
            .Parameters.AddWithValue("@AppUserIdFormatCode", sAppUserIdFormatCode)
            .Parameters.AddWithValue("@SendMessageDetail", bSendMessageDetail)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveMasterAppGroup3(ByVal sActionType As String, ByVal sAppGroupID As String, ByVal sAppGroupDescription As String, ByVal sName As String, ByVal sPrimaryContactName As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sPrimaryCell As String, ByVal sPrimaryEmail As String, ByVal sEmergencyEmail As String, ByVal sEmergencyAlertLabel1 As String, ByVal sEmergencyAlertLabel2 As String, ByVal sEmergencyAlertLabel3 As String, ByVal sEmergencyAlertNumber1 As String, ByVal sEmergencyAlertNumber2 As String, ByVal sEmergencyAlertNumber3 As String, ByVal lNotDeliveredAfterSeconds As Long, ByVal lNotReadAfterSeconds As Long, ByVal sCannedQuestion1 As String, ByVal sCannedQuestion2 As String, ByVal sCannedQuestion3 As String, ByVal sCannedQuestion4 As String, ByVal sCannedQuestion5 As String, ByVal sCannedQuestion6 As String, ByVal sCannedReply1 As String, ByVal sCannedReply2 As String, ByVal sCannedReply3 As String, ByVal sCannedReply4 As String, ByVal sCannedReply5 As String, ByVal sCannedReply6 As String, ByVal sReceiptNotificationEmail As String, ByVal sReplyNotificationEmail As String, ByVal lInfoURLInterval As Long, ByVal bChangeHeaderImage As Boolean, ByVal bChangeFooterImage As Boolean, ByVal bActive As Boolean, ByVal bAutoCreateReply As Boolean, ByVal sDomainName As String, ByVal sCoCode As String, ByVal sAcctNo As String, ByVal sTimeZoneCode As String, ByVal bInternalGroup As Boolean, ByVal bGroupMessaging As Boolean, ByVal bOnlyAllowMessagesFromWithinGroup As Boolean, ByVal bAllowMessageFromPrimaryEmail As Boolean, ByVal bAllowReplyOnGroupMessage As Boolean, ByVal lMessageRetentionDays As Long, ByVal sGroupMessageFromIndividualOrGroup As String, ByVal sAppUserIdFormatCode As String, ByVal bSendMessageDetail As Boolean, lMessageExpireDays As String, sDefaultAddressBook As String) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterAppGroup3"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppGroupDescription", sAppGroupDescription)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@PrimaryContactName", sPrimaryContactName)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@EmergencyEmailAddress", sEmergencyEmail)
            .Parameters.AddWithValue("@EmergencyAlertLabel1", sEmergencyAlertLabel1)
            .Parameters.AddWithValue("@EmergencyAlertLabel2", sEmergencyAlertLabel2)
            .Parameters.AddWithValue("@EmergencyAlertLabel3", sEmergencyAlertLabel3)
            .Parameters.AddWithValue("@EmergencyAlertNumber1", sEmergencyAlertNumber1)
            .Parameters.AddWithValue("@EmergencyAlertNumber2", sEmergencyAlertNumber2)
            .Parameters.AddWithValue("@EmergencyAlertNumber3", sEmergencyAlertNumber3)
            .Parameters.AddWithValue("@NotDeliveredAfterSeconds", lNotDeliveredAfterSeconds)
            .Parameters.AddWithValue("@NotReadAfterSeconds", lNotReadAfterSeconds)
            .Parameters.AddWithValue("@CannedQuestion1", sCannedQuestion1)
            .Parameters.AddWithValue("@CannedQuestion2", sCannedQuestion2)
            .Parameters.AddWithValue("@CannedQuestion3", sCannedQuestion3)
            .Parameters.AddWithValue("@CannedQuestion4", sCannedQuestion4)
            .Parameters.AddWithValue("@CannedQuestion5", sCannedQuestion5)
            .Parameters.AddWithValue("@CannedQuestion6", sCannedQuestion6)
            .Parameters.AddWithValue("@CannedReply1Email", sCannedReply1)
            .Parameters.AddWithValue("@CannedReply2Email", sCannedReply2)
            .Parameters.AddWithValue("@CannedReply3Email", sCannedReply3)
            .Parameters.AddWithValue("@CannedReply4Email", sCannedReply4)
            .Parameters.AddWithValue("@CannedReply5Email", sCannedReply5)
            .Parameters.AddWithValue("@CannedReply6Email", sCannedReply6)
            .Parameters.AddWithValue("@ReceiptNotificationEmail", sReceiptNotificationEmail)
            .Parameters.AddWithValue("@ReplyNotificationEmail", sReplyNotificationEmail)
            .Parameters.AddWithValue("@InfoURLInterval", lInfoURLInterval)
            .Parameters.AddWithValue("@ChangeHeaderImage", bChangeHeaderImage)
            .Parameters.AddWithValue("@ChangeFooterImage", bChangeFooterImage)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@AutoCreateReply", bAutoCreateReply)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@TimeZoneCode", sTimeZoneCode)
            .Parameters.AddWithValue("@InternalGroup", bInternalGroup)
            .Parameters.AddWithValue("@GroupMessaging", bGroupMessaging)
            .Parameters.AddWithValue("@OnlyAllowMessagesFromWithinGroup", bOnlyAllowMessagesFromWithinGroup)
            .Parameters.AddWithValue("@AllowMessageFromPrimaryEmail", bAllowMessageFromPrimaryEmail)
            .Parameters.AddWithValue("@AllowReplyOnGroupMessage", bAllowReplyOnGroupMessage)
            .Parameters.AddWithValue("@MessageRetentionDays", lMessageRetentionDays)
            .Parameters.AddWithValue("@GroupMessageFromIndividualOrGroup", sGroupMessageFromIndividualOrGroup)
            .Parameters.AddWithValue("@AppUserIdFormatCode", sAppUserIdFormatCode)
            .Parameters.AddWithValue("@SendMessageDetail", bSendMessageDetail)
            .Parameters.AddWithValue("@MessageExpireDays", lMessageExpireDays)
            .Parameters.AddWithValue("@DefaultAddressBook", sDefaultAddressBook)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetMasterProcessControlHistoryDetail(ByVal sAppUserID As String, ByVal sProcessDescription As String, ByVal sProcessDate As String, ByVal bShowTopOne As Boolean) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterProcessControlHistoryDetail"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@ProcessDescription", sProcessDescription)
            .Parameters.AddWithValue("@ProcessDate", sProcessDate)
            .Parameters.AddWithValue("@ShowTopOne", bShowTopOne)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterProcessControlHistoryDetail = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveMasterContentGroupURL(ByVal sActionType As String, ByVal sAppGroupID As String, ByVal sContentGroupID As String, ByVal lContentGroupURLSequence As Long, ByVal sContentGroupURL As String, ByVal bURLOnly As Boolean, ByVal sOriginalFileName As String, ByVal sContentDescription As String, ByVal sSubmittedBy As String, ByVal sApprovedBy As String, ByVal sDateApproved As String, ByVal sExpirationDate As String, ByVal sRejectedBy As String, ByVal sDateRejected As String, ByVal sReasonRejected As String, ByVal bActive As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterContentGroupURL"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@ContentGroupId", sContentGroupID)
            .Parameters.AddWithValue("@ContentGroupURLSequence", lContentGroupURLSequence)
            .Parameters.AddWithValue("@ContentGroupURL", sContentGroupURL)
            .Parameters.AddWithValue("@URLOnly", bURLOnly)
            .Parameters.AddWithValue("@OriginalFileName", sOriginalFileName)
            .Parameters.AddWithValue("@ContentDescription", sContentDescription)
            .Parameters.AddWithValue("@SubmittedBy", sSubmittedBy)
            .Parameters.AddWithValue("@ApprovedBy", sApprovedBy)
            .Parameters.AddWithValue("@DateApproved", sDateApproved)
            .Parameters.AddWithValue("@ExpirationDate", sExpirationDate)
            .Parameters.AddWithValue("@RejectedBy", sRejectedBy)
            .Parameters.AddWithValue("@DateRejected", sDateRejected)
            .Parameters.AddWithValue("@ReasonRejected", sReasonRejected)
            .Parameters.AddWithValue("@Active", bActive)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetMasterContentGroupURL(ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sContentGroupID As String, ByVal lContentGroupURLSequence As Long, ByVal bShowTopOne As Boolean) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterContentGroupURL"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@ContentGroupID", sContentGroupID)
            .Parameters.AddWithValue("@ContentGroupURLSequence", lContentGroupURLSequence)
            .Parameters.AddWithValue("@ShowTopOne", bShowTopOne)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterContentGroupURL = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveMasterContentGroupAdministrator(ByVal sActionType As String, ByVal sAppGroupID As String, ByVal sContentGroupID As String, ByVal sAppUserID As String, ByVal bActive As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterContentGroupAdministrator"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@ContentGroupId", sContentGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@Active", bActive)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetMasterContentGroupAdministrator(ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sContentGroupID As String, ByVal bShowTopOne As Boolean) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterContentGroupAdministrator"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@ContentGroupID", sContentGroupID)
            .Parameters.AddWithValue("@ShowTopOne", bShowTopOne)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterContentGroupAdministrator = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveMasterContentGroupSource(ByVal sActionType As String, ByVal sAppGroupID As String, ByVal sContentGroupID As String, ByVal sSourceEmailID As String, ByVal sSourceName As String, ByVal bActive As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterContentGroupSource"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@ContentGroupId", sContentGroupID)
            .Parameters.AddWithValue("@SourceEmailID", sSourceEmailID)
            .Parameters.AddWithValue("@SourceName", sSourceName)
            .Parameters.AddWithValue("@Active", bActive)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetMasterContentGroupSource(ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sContentGroupID As String, ByVal sSourceEmailID As String, ByVal bShowTopOne As Boolean) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterContentGroupSource"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@ContentGroupID", sContentGroupID)
            .Parameters.AddWithValue("@SourceEmailID", sSourceEmailID)
            .Parameters.AddWithValue("@ShowTopOne", bShowTopOne)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterContentGroupSource = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function UserIsAdmin(ByVal sAppUserId As String) As String
        Dim retValue As String = "0"
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUserIsAdmin"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetMasterMessageManagerDetail(ByVal AppUserID As String, ByVal AppCode As String, ByVal MessageNumber As String, ByVal DeviceID As String, ByVal LoggedInUser As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterMessageManagerDetail"
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@MessageNumber", MessageNumber)
            .Parameters.AddWithValue("@DeviceID", DeviceID)
            .Parameters.AddWithValue("@LoggedInUser", LoggedInUser)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterMessageManagerDetail = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetMasterMessageHistoryDetail(ByVal AppUserID As String, ByVal AppCode As String, ByVal MessageNumber As String, ByVal DeviceID As String, ByVal LoggedInUser As String, Optional ByVal bGetLast As Boolean = False) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterMessageHistoryDetail"
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@MessageNumber", MessageNumber)
            .Parameters.AddWithValue("@DeviceID", DeviceID)
            .Parameters.AddWithValue("@LoggedInUser", LoggedInUser)
            .Parameters.AddWithValue("@GetLast", bGetLast)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterMessageHistoryDetail = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetUserMessageHistoryMessage(ByVal AppGroupID As String, ByVal AppUserID As String, ByVal RR As String, ByVal BeginDate As String, ByVal EndDate As String, ByVal SearchTerm As String, ByVal LoggedInUser As String, ByVal MessageNumber As String, ByVal PrevNextFirstLast As String, ByVal DeviceID As String, Optional ByVal OrigMessageNumber As String = "0", Optional ByVal TimerMinutes As String = "2") As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserMessageHistoryMessage"
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@RR", RR)
            .Parameters.AddWithValue("@BeginDate", BeginDate)
            .Parameters.AddWithValue("@EndDate", EndDate)
            .Parameters.AddWithValue("@SearchTerm", SearchTerm)
            .Parameters.AddWithValue("@LoggedInUser", LoggedInUser)
            .Parameters.AddWithValue("@MessageNumber", MessageNumber)
            .Parameters.AddWithValue("@PrevNextFirstLast", PrevNextFirstLast)
            .Parameters.AddWithValue("@DeviceID", DeviceID)
            .Parameters.AddWithValue("@OrigMessageNumber", OrigMessageNumber)
            .Parameters.AddWithValue("@TimerMinutes", TimerMinutes)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserMessageHistoryMessage = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetUserMessageHistoryDetail(ByVal AppGroupID As String, ByVal AppUserID As String, ByVal RR As String, ByVal BeginDate As String, ByVal EndDate As String, ByVal SearchTerm As String, ByVal LoggedInUser As String, ByVal MessageNumber As String, ByVal PrevNextFirstLast As String, ByVal DeviceID As String, Optional ByVal OrigMessageNumber As String = "0", Optional ByVal TimerMinutes As String = "2") As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserMessageHistoryDetail"
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@RR", RR)
            .Parameters.AddWithValue("@BeginDate", BeginDate)
            .Parameters.AddWithValue("@EndDate", EndDate)
            .Parameters.AddWithValue("@SearchTerm", SearchTerm)
            .Parameters.AddWithValue("@LoggedInUser", LoggedInUser)
            .Parameters.AddWithValue("@MessageNumber", MessageNumber)
            .Parameters.AddWithValue("@PrevNextFirstLast", PrevNextFirstLast)
            .Parameters.AddWithValue("@DeviceID", DeviceID)
            .Parameters.AddWithValue("@OrigMessageNumber", OrigMessageNumber)
            .Parameters.AddWithValue("@TimerMinutes", TimerMinutes)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserMessageHistoryDetail = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SendCreatedMessage(ByVal sAppCode As String, ByVal sAppUserId As String, ByVal sDeviceId As String, ByVal sMessageTo As String, ByVal sMessageSubject As String, ByVal sMessageBody As String, Optional ByVal lReplyToMessageNumber As Long = 0) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim rdr As SqlDataReader
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "SendCreatedMessage2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@DeviceId", sDeviceId)
            .Parameters.AddWithValue("@MessageTo", sMessageTo)
            .Parameters.AddWithValue("@MessageSubject", sMessageSubject)
            .Parameters.AddWithValue("@MessageBody", sMessageBody)
            .Parameters.AddWithValue("@ReplyToMessageNumber", lReplyToMessageNumber)
        End With

        'HttpContext.Current.Response.Write("EXEC SendCreatedMessage '" & sAppCode & "', '" & sAppUserId & "', '" & sDeviceId & "', '" & sMessageTo & "', '" & sMessageSubject & "', '" & sMessageBody & "'")
        'HttpContext.Current.Response.End()
        cn.Open()
        rdr = cmdSQL.ExecuteReader()
        While rdr.Read
            retValue = rdr("ReturnMsg")
        End While
        
        Try
            rdr.Close()
            rdr = Nothing
        Catch ex As Exception

        End Try
        
        If Not cmdSQL Is Nothing Then
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End If

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

        Return retValue

    End Function

    Public Shared Function GetUserSecurityGroup(ByVal sAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserSecurityGroup"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadAppScreen() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppScreens"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAppScreen = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadSecurityGroup(ByVal sSecurityGroupName As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSecurityGroups"
            .Parameters.AddWithValue("@SecurityGroupName", sSecurityGroupName)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadSecurityGroup = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function CheckSecurity(ByVal sAppUserID As String, ByVal sAppScreenName As String) As Long
        Dim retValue As Long = 0
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCheckSecurity"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppScreenName", sAppScreenName)
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

        End Try

        Return retValue

    End Function

    Public Shared Function StartAllProcesses() As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "StartAllProcesses"
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

        End Try

        Return retValue

    End Function

    Public Shared Function StopAllProcesses(ByVal sAppUserID As String, ByVal sSystemShutdownReason As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "StopAllProcesses"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@SystemShutdownReason", sSystemShutdownReason)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SendAndroidEmail(ByVal sEmailAddress As String, ByVal sAttachment As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSendAndroidEmail"
            .Parameters.AddWithValue("@To", sEmailAddress)
            .Parameters.AddWithValue("@Attachment", sAttachment)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetUserGroup(ByVal sAppUserId As String, Optional ByVal bGetPrimary As Boolean = False) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserGroup"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@GetPrimary", bGetPrimary)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadTimeZones() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminTimeZones"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadTimeZones = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadBillRates() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminBillRate"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadBillRates = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadUserTypes() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUserType"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadUserTypes = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadCellCarriers() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCellCarriers"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadCellCarriers = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAdminUsers() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUserGetUsers"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAdminUsers = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function UserIsGroupAdminForGroup(ByVal sAppUserId As String, ByVal sAppGroupId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUserIsGroupAdminForGroup"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadInviteUsers(ByVal AppUserID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminInviteUserUsers"
            .Parameters.AddWithValue("@AppUserId", AppUserID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadInviteUsers = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadInviteUsers2(ByVal AppUserID As String, Optional ByVal sSearchTerm As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminInviteUserUsers2"
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadInviteUsers2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAppUserName(ByVal sAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUserName"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SendInviteEmail(ByVal sTo As String, ByVal sFrom As String, ByVal sSubject As String, ByVal sBody As String, Optional ByVal sCC As String = "", Optional ByVal sBCC As String = "") As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSendInviteEmail"
            .Parameters.AddWithValue("@To", sTo)
            .Parameters.AddWithValue("@From", sFrom)
            .Parameters.AddWithValue("@Subject", sSubject)
            .Parameters.AddWithValue("@Body", sBody)
            .Parameters.AddWithValue("@CC", sCC)
            .Parameters.AddWithValue("@BCC", sBCC)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SendInviteEmail2(ByVal sAppUserID As String, ByVal sTo As String, ByVal sFrom As String, ByVal sSubject As String, Optional ByVal sCC As String = "", Optional ByVal sBCC As String = "") As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSendInviteEmail2"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@To", sTo)
            .Parameters.AddWithValue("@From", sFrom)
            .Parameters.AddWithValue("@Subject", sSubject)
            .Parameters.AddWithValue("@CC", sCC)
            .Parameters.AddWithValue("@BCC", sBCC)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetAppUserPrimaryEmail(ByVal sAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUserPrimaryEmail"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetLogoDirectory() As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetLogoDirectory"
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetLogoURLDirectory() As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetLogoURLDirectory"
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAppGroupFooterImageName(ByVal sAppCode As String, ByVal sAppGroupID As String, ByVal sDeviceTypeCode As String, ByVal sFooterImageName As String, ByVal sImageLocationURL As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppGroupFooterImageName"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@FooterImageName", sFooterImageName)
            .Parameters.AddWithValue("@ImageLocationURL", sImageLocationURL)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAppGroupHeaderImageName(ByVal sAppCode As String, ByVal sAppGroupID As String, ByVal sDeviceTypeCode As String, ByVal sHeaderImageName As String, ByVal sImageLocationURL As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppGroupHeaderImageName"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@HeaderImageName", sHeaderImageName)
            .Parameters.AddWithValue("@ImageLocationURL", sImageLocationURL)
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

        End Try

        Return retValue

    End Function

    Public Shared Function UserCanChangeFooterLogo(ByVal sAppUserID As String, ByVal sAppGroupID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUserCanChangeFooterLogo"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function UserCanChangeHeaderLogo(ByVal sAppUserID As String, ByVal sAppGroupID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUserCanChangeHeaderLogo"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetAppGroupInfoURLDisplayCurrent(ByVal sAppGroupID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppGroupInfoURLDisplayCurrent"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAppGroupInfoURLDisplayCurrent = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetFooterImageURL(ByVal sAppCode As String, ByVal sAppGroupID As String, ByVal sDeviceTypeCode As String, Optional ByVal bShowDefault As Boolean = True) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppGroupDeviceTypeFooterURL"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@ShowDefault", bShowDefault)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetHeaderImageURL(ByVal sAppCode As String, ByVal sAppGroupID As String, ByVal sDeviceTypeCode As String, Optional ByVal bShowDefault As Boolean = True) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppGroupDeviceTypeHeaderURL"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@ShowDefault", bShowDefault)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveUploadInfoScreen(ByVal sAppGroupID As String, ByVal lInfoURLSequence As Long, ByVal sInfoURL As String, ByVal bURLOnly As Boolean, ByVal sSelectedBy As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveUploadInfoScreen"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@InfoURLSequence", lInfoURLSequence)
            .Parameters.AddWithValue("@InfoURL", sInfoURL)
            .Parameters.AddWithValue("@URLOnly", bURLOnly)
            .Parameters.AddWithValue("@SelectedBy", sSelectedBy)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveImportAppUsers(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal lLineNumber As Long, ByVal sCompanyAppGroupId As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sBusinessEmailAddress As String, ByVal sPagerNumber As String, ByVal sAppUserIdDefaultName As String, ByVal sAppUserIdOptionalPager As String, ByVal sInitialPassword As String, ByVal sGroupId1 As String, ByVal sGroupId2 As String, ByVal sGroupId3 As String, ByVal sGroupId4 As String, ByVal sGroupId5 As String, ByVal sGroupId6 As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveImportAppUsers"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", lLineNumber)
            .Parameters.AddWithValue("@CompanyAppGroupId", sCompanyAppGroupId)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@BusinessEmailAddress", sBusinessEmailAddress)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@AppUserIdDefaultName", sAppUserIdDefaultName)
            .Parameters.AddWithValue("@AppUserIdOptionalPager", sAppUserIdOptionalPager)
            .Parameters.AddWithValue("@InitialPassword", sInitialPassword)
            .Parameters.AddWithValue("@GroupID1", sGroupId1)
            .Parameters.AddWithValue("@GroupID2", sGroupId2)
            .Parameters.AddWithValue("@GroupID3", sGroupId3)
            .Parameters.AddWithValue("@GroupID4", sGroupId4)
            .Parameters.AddWithValue("@GroupID5", sGroupId5)
            .Parameters.AddWithValue("@GroupID6", sGroupId6)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveImportAppUsers2(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal lLineNumber As Long, ByVal sCompanyAppGroupId As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sBusinessEmailAddress As String, ByVal sPagerNumber As String, ByVal sAppUserIdDefaultName As String, ByVal sAppUserIdOptionalPager As String, ByVal sInitialPassword As String, ByVal sGroupId1 As String, ByVal sGroupId2 As String, ByVal sGroupId3 As String, ByVal sGroupId4 As String, ByVal sGroupId5 As String, ByVal sGroupId6 As String, ByVal sSendInvitation As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveImportAppUsers2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", lLineNumber)
            .Parameters.AddWithValue("@CompanyAppGroupId", sCompanyAppGroupId)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@BusinessEmailAddress", sBusinessEmailAddress)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@AppUserIdDefaultName", sAppUserIdDefaultName)
            .Parameters.AddWithValue("@AppUserIdOptionalPager", sAppUserIdOptionalPager)
            .Parameters.AddWithValue("@InitialPassword", sInitialPassword)
            .Parameters.AddWithValue("@GroupID1", sGroupId1)
            .Parameters.AddWithValue("@GroupID2", sGroupId2)
            .Parameters.AddWithValue("@GroupID3", sGroupId3)
            .Parameters.AddWithValue("@GroupID4", sGroupId4)
            .Parameters.AddWithValue("@GroupID5", sGroupId5)
            .Parameters.AddWithValue("@GroupID6", sGroupId6)
            .Parameters.AddWithValue("@SendInvitation", sSendInvitation)
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

        End Try

        Return retValue

    End Function

    Public Shared Function AdminUpdateImportAppUsers(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal lLineNumber As Long, ByVal sCompanyAppGroupID As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sBusinessEmailAddress As String, ByVal sPagerNumber As String, ByVal sAppUserIdDefaultName As String, ByVal sAppUserIdOptionalPager As String, ByVal sInitialPassword As String, ByVal sGroupId1 As String, ByVal sGroupId2 As String, ByVal sGroupId3 As String, ByVal sGroupId4 As String, ByVal sGroupId5 As String, ByVal sGroupId6 As String, ByVal bDoNotImport As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUpdateImportAppUsers"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", lLineNumber)
            .Parameters.AddWithValue("@CompanyAppGroupID", sCompanyAppGroupID)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@BusinessEmailAddress", sBusinessEmailAddress)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@AppUserIdDefaultName", sAppUserIdDefaultName)
            .Parameters.AddWithValue("@AppUserIdOptionalPager", sAppUserIdOptionalPager)
            .Parameters.AddWithValue("@InitialPassword", sInitialPassword)
            .Parameters.AddWithValue("@GroupId1", sGroupId1)
            .Parameters.AddWithValue("@GroupId2", sGroupId2)
            .Parameters.AddWithValue("@GroupId3", sGroupId3)
            .Parameters.AddWithValue("@GroupId4", sGroupId4)
            .Parameters.AddWithValue("@GroupId5", sGroupId5)
            .Parameters.AddWithValue("@GroupId6", sGroupId6)
            .Parameters.AddWithValue("@DoNotImport", bDoNotImport)

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

        End Try

        Return retValue

    End Function

    Public Shared Function AdminUpdateImportAppUsers2(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal lLineNumber As Long, ByVal sCompanyAppGroupID As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sBusinessEmailAddress As String, ByVal sPagerNumber As String, ByVal sAppUserIdDefaultName As String, ByVal sAppUserIdOptionalPager As String, ByVal sInitialPassword As String, ByVal sGroupId1 As String, ByVal sGroupId2 As String, ByVal sGroupId3 As String, ByVal sGroupId4 As String, ByVal sGroupId5 As String, ByVal sGroupId6 As String, ByVal bDoNotImport As Boolean, ByVal sSendInvitation As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUpdateImportAppUsers2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", lLineNumber)
            .Parameters.AddWithValue("@CompanyAppGroupID", sCompanyAppGroupID)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@BusinessEmailAddress", sBusinessEmailAddress)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@AppUserIdDefaultName", sAppUserIdDefaultName)
            .Parameters.AddWithValue("@AppUserIdOptionalPager", sAppUserIdOptionalPager)
            .Parameters.AddWithValue("@InitialPassword", sInitialPassword)
            .Parameters.AddWithValue("@GroupId1", sGroupId1)
            .Parameters.AddWithValue("@GroupId2", sGroupId2)
            .Parameters.AddWithValue("@GroupId3", sGroupId3)
            .Parameters.AddWithValue("@GroupId4", sGroupId4)
            .Parameters.AddWithValue("@GroupId5", sGroupId5)
            .Parameters.AddWithValue("@GroupId6", sGroupId6)
            .Parameters.AddWithValue("@DoNotImport", bDoNotImport)
            .Parameters.AddWithValue("@SendInvitation", sSendInvitation)

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

        End Try

        Return retValue

    End Function

    Public Shared Function CheckImportAppUsers(ByVal sAppCode As String, ByVal sImportAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAppUsers_Check"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@ReturnMsg", "")
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

        End Try

        Return retValue

    End Function

    Public Shared Function DeleteImportAppUsers(ByVal sAppCode As String, ByVal sImportAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeleteImportAppUsers"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function UpdateImportAppUsers(ByVal sAppCode As String, ByVal sImportAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAppUsers_Update"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function ForgotPassword(ByVal sAppUserID As String, Optional ByVal sAlternateEmail As String = "") As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminForgotPassword"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AlternateEmail", sAlternateEmail)
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

        End Try

        Return retValue

    End Function

    Public Shared Function ForgotLogin(ByVal sPrimaryEmail As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminForgotLogin"
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetDeviceTypes(Optional ByVal sDeviceTypeCode As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterDeviceType"
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetDeviceTypes = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetNextInfoUrl(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String, Optional ByVal sAppGroupID As String = "", Optional ByVal sInfoDate As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "GetNextInfoUrl"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@InfoDate", sInfoDate)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetNextInfoUrl = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetPrevInfoUrl(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String, Optional ByVal sAppGroupID As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "GetPrevInfoUrl"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetPrevInfoUrl = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetFirstInfoUrl(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String, Optional ByVal sAppGroupID As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "GetFirstInfoUrl"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetFirstInfoUrl = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetLastInfoUrl(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String, Optional ByVal sAppGroupID As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "GetLastInfoUrl"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetLastInfoUrl = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetUserTimeZoneDateTime(ByVal sAppUserID As String, ByVal sDateTimeToConvert As String) As String

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserTimeZoneDateTime"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DateTimeToConvert", sDateTimeToConvert)
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

        End Try

        Return retValue

    End Function

    Public Shared Function FlagSendTestPush(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String) As String

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUserDeviceSendTestPush"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function CanClearMessage(ByVal sMessageNumber As String) As Boolean

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCanClearMessage"
            .Parameters.AddWithValue("@MessageNumber", sMessageNumber)
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

        End Try

        Return retValue

    End Function

    Public Shared Function IsSystemShutdown(ByVal sSystemCode As String) As String
        Dim retValue As String = "1"
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminIsSystemShutdown"
            .Parameters.AddWithValue("@SystemCode", sSystemCode)
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

        End Try
        'for testing retValue = "1"
        Return retValue

    End Function

    Public Shared Function DeactivateUnusedDevices(ByVal sAppCode As String, ByVal sAppUserID As String) As String

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeactivateAppUserDeviceUnused"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetEULAText(ByVal sAppCode As String, ByVal sDeviceTypeCode As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetEULAText"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveEULAAccept(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveEULAAccept"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadAddressBookTypes(ByVal sAppCode As String, ByVal sAppUserID As String, Optional ByVal sAddressBookTypeCode As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBookType"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookTypes = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAddressBookTypes2(ByVal sAppCode As String, ByVal sAppUserID As String, Optional ByVal sAddressBookTypeCode As String = "", Optional ByVal bIsImport As Boolean = True) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBookType2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@isImport", bIsImport)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookTypes2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAddressBook(ByVal sAppCode As String, ByVal sAddressBookTypeCode As String, ByVal sAppUserId As String, ByVal sAppGroupId As String, ByVal sDomainName As String, ByVal sAddressBookID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBook"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAddressBook = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAddressBookContact(ByVal sActionType As String, sAppCode As String, sAddressBookTypeCode As String, sAppUserID As String, sAppGroupId As String, sDomainName As String, sAddressBookID As String, sTitle As String, sFirstName As String, sMiddleName As String, sLastName As String, sSuffix As String, sEmailAddress As String, sEmailDisplayName As String, ByVal sCompany As String, ByVal sDepartment As String, ByVal sJobTitle As String, ByVal sBusinessStreet As String, ByVal sBusinessStreet2 As String, ByVal sBusinessStreet3 As String, ByVal sBusinessCity As String, ByVal sBusinessState As String, ByVal sBusinessPostalCode As String, ByVal sBusinessCountryRegion As String, ByVal sHomeStreet As String, ByVal sHomeStreet2 As String, ByVal sHomeStreet3 As String, ByVal sHomeCity As String, ByVal sHomeState As String, ByVal sHomePostalCode As String, ByVal sHomeCountryRegion As String, ByVal sOtherStreet As String, ByVal sOtherStreet2 As String, ByVal sOtherStreet3 As String, ByVal sOtherCity As String, ByVal sOtherState As String, ByVal sOtherPostalCode As String, ByVal sOtherCountryRegion As String, ByVal sAssistantsPhone As String, ByVal sBusinessFax As String, ByVal sBusinessPhone As String, ByVal sBusinessPhone2 As String, ByVal sCallback As String, ByVal sCarPhone As String, ByVal sCompanyMainPhone As String, ByVal sHomeFax As String, ByVal sHomePhone As String, ByVal sHomePhone2 As String, ByVal sISDN As String, ByVal sMobilePhone As String, ByVal sOtherFax As String, ByVal sOtherPhone As String, ByVal sPager As String, ByVal sPrimaryPhone As String, ByVal sRadioPhone As String, ByVal sTTYTDDPhone As String, ByVal sTelex As String, ByVal sAccount As String, ByVal sAnniversary As String, ByVal sAssistantsName As String, ByVal sBillingInformation As String, ByVal sBirthday As String, ByVal sBusinessAddressPOBox As String, ByVal sCategories As String, ByVal sChildren As String, ByVal sDirectoryServer As String, ByVal sEmailType As String, ByVal sEmail2Address As String, ByVal sEmail2Type As String, ByVal sEmail2DisplayName As String, ByVal sEmail3Address As String, ByVal sEmail3Type As String, ByVal sEmail3DisplayName As String, ByVal sGender As String, ByVal sGovernmentIDNumber As String, ByVal sHobby As String, ByVal sHomeAddressPOBox As String, ByVal sInitials As String, ByVal sInternetFreeBusy As String, ByVal sKeywords As String, ByVal sLanguage1 As String, ByVal sLocation As String, ByVal sManagersName As String, ByVal sMileage As String, ByVal sNotes As String, ByVal sOfficeLocation As String, ByVal sOrganizationalIDNumber As String, ByVal sOtherAddressPOBox As String, ByVal sPriority As String, ByVal sPrivate As String, ByVal sProfession As String, ByVal sReferredBy As String, ByVal sSensitivity As String, ByVal sSpouse As String, ByVal sUser1 As String, ByVal sUser2 As String, ByVal sUser3 As String, ByVal sUser4 As String, ByVal sWebPage As String, ByVal sSupervisor As String, ByVal sSupervisorPhone As String, ByVal sSupervisorEmail As String, ByVal sSupervisorAssistant As String, ByVal sSupervisorAssistantPhone As String, ByVal sSupervisorAssistantEmail As String, ByVal sDepartmentEscalationsContact As String, ByVal sDepartmentEscalationsContactNumber As String, ByVal sDepartmentEscalationsEmail As String, ByVal sCurrentEscalationsContact As String, ByVal sCurrentEscalationsContactPhoneNumber As String, ByVal sCurrentEscalationsEmail As String, ByVal sCurrentEscalationDateFrom As String, ByVal sCurrentEscalationDateTo As String, ByVal sSecondaryEscalationsContact As String, ByVal sSecondaryEscalationsContactPhoneNumber As String, ByVal sSecondaryEscalationsEmail As String, ByVal sSecondaryEscalationsDateFrom As String, ByVal sSecondaryEscalationsDateTo As String, ByVal sTemporaryForwardingEmailAddress1 As String, ByVal sTemporaryForwardingAddress1FromDate As String, ByVal sTemporaryForwardingAddress1ToDate As String, ByVal sTemporaryForwardingEmailAddress2 As String, ByVal sTemporaryForwardingAddress2FromDate As String, ByVal sTemporaryForwardingAddress2ToDate As String, ByVal sBestWaytoContactDuringBusinessHours As String, ByVal sBestWaytoContactAfterBusinessHours As String, ByVal sContactInformationNotes As String, ByVal sEscalationInformationNotes As String, ByVal sInCaseofEmergencyContactInformationName As String, ByVal sInCaseofEmergencyContactRelationship As String, ByVal sInCaseofEmergencyContactInformationPhoneNumber As String, ByVal sInCaseOfEmergencyContactInformationEmailAddress As String, ByVal sContactPriority1 As String, ByVal sContactPriority2 As String, ByVal sContactPriority3 As String, ByVal sFileAs As String, ByVal sPagerType As String, ByVal sDispatcherInfo As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookContact"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
            .Parameters.AddWithValue("@Title", sTitle)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@MiddleName", sMiddleName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@Suffix", sSuffix)
            .Parameters.AddWithValue("@EmailAddress", sEmailAddress)
            .Parameters.AddWithValue("@EmailDisplayName", sEmailDisplayName)
            .Parameters.AddWithValue("@Company", sCompany)
            .Parameters.AddWithValue("@Department", sDepartment)
            .Parameters.AddWithValue("@JobTitle", sJobTitle)
            .Parameters.AddWithValue("@BusinessStreet", sBusinessStreet)
            .Parameters.AddWithValue("@BusinessStreet2", sBusinessStreet2)
            .Parameters.AddWithValue("@BusinessStreet3", sBusinessStreet3)
            .Parameters.AddWithValue("@BusinessCity", sBusinessCity)
            .Parameters.AddWithValue("@BusinessState", sBusinessState)
            .Parameters.AddWithValue("@BusinessPostalCode", sBusinessPostalCode)
            .Parameters.AddWithValue("@BusinessCountryRegion", sBusinessCountryRegion)
            .Parameters.AddWithValue("@HomeStreet", sHomeStreet)
            .Parameters.AddWithValue("@HomeStreet2", sHomeStreet2)
            .Parameters.AddWithValue("@HomeStreet3", sHomeStreet3)
            .Parameters.AddWithValue("@HomeCity", sHomeCity)
            .Parameters.AddWithValue("@HomeState", sHomeState)
            .Parameters.AddWithValue("@HomePostalCode", sHomePostalCode)
            .Parameters.AddWithValue("@HomeCountryRegion", sHomeCountryRegion)
            .Parameters.AddWithValue("@OtherStreet", sOtherStreet)
            .Parameters.AddWithValue("@OtherStreet2", sOtherStreet2)
            .Parameters.AddWithValue("@OtherStreet3", sOtherStreet3)
            .Parameters.AddWithValue("@OtherCity", sOtherCity)
            .Parameters.AddWithValue("@OtherState", sOtherState)
            .Parameters.AddWithValue("@OtherPostalCode", sOtherPostalCode)
            .Parameters.AddWithValue("@OtherCountryRegion", sOtherCountryRegion)
            .Parameters.AddWithValue("@AssistantsPhone", sAssistantsPhone)
            .Parameters.AddWithValue("@BusinessFax", sBusinessFax)
            .Parameters.AddWithValue("@BusinessPhone", sBusinessPhone)
            .Parameters.AddWithValue("@BusinessPhone2", sBusinessPhone2)
            .Parameters.AddWithValue("@Callback", sCallback)
            .Parameters.AddWithValue("@CarPhone", sCarPhone)
            .Parameters.AddWithValue("@CompanyMainPhone", sCompanyMainPhone)
            .Parameters.AddWithValue("@HomeFax", sHomeFax)
            .Parameters.AddWithValue("@HomePhone", sHomePhone)
            .Parameters.AddWithValue("@HomePhone2", sHomePhone2)
            .Parameters.AddWithValue("@ISDN", sISDN)
            .Parameters.AddWithValue("@MobilePhone", sMobilePhone)
            .Parameters.AddWithValue("@OtherFax", sOtherFax)
            .Parameters.AddWithValue("@OtherPhone", sOtherPhone)
            .Parameters.AddWithValue("@Pager", sPager)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@RadioPhone", sRadioPhone)
            .Parameters.AddWithValue("@TTYTDDPhone", sTTYTDDPhone)
            .Parameters.AddWithValue("@Telex", sTelex)
            .Parameters.AddWithValue("@Account", sAccount)
            .Parameters.AddWithValue("@Anniversary", sAnniversary)
            .Parameters.AddWithValue("@AssistantsName", sAssistantsName)
            .Parameters.AddWithValue("@BillingInformation", sBillingInformation)
            .Parameters.AddWithValue("@Birthday", sBirthday)
            .Parameters.AddWithValue("@BusinessAddressPOBox", sBusinessAddressPOBox)
            .Parameters.AddWithValue("@Categories", sCategories)
            .Parameters.AddWithValue("@Children", sChildren)
            .Parameters.AddWithValue("@DirectoryServer", sDirectoryServer)
            .Parameters.AddWithValue("@EmailType", sEmailType)
            .Parameters.AddWithValue("@Email2Address", sEmail2Address)
            .Parameters.AddWithValue("@Email2Type", sEmail2Type)
            .Parameters.AddWithValue("@Email2DisplayName", sEmail2DisplayName)
            .Parameters.AddWithValue("@Email3Address", sEmail3Address)
            .Parameters.AddWithValue("@Email3Type", sEmail3Type)
            .Parameters.AddWithValue("@Email3DisplayName", sEmail3DisplayName)
            .Parameters.AddWithValue("@Gender", sGender)
            .Parameters.AddWithValue("@GovernmentIDNumber", sGovernmentIDNumber)
            .Parameters.AddWithValue("@Hobby", sHobby)
            .Parameters.AddWithValue("@HomeAddressPOBox", sHomeAddressPOBox)
            .Parameters.AddWithValue("@Initials", sInitials)
            .Parameters.AddWithValue("@InternetFreeBusy", sInternetFreeBusy)
            .Parameters.AddWithValue("@Keywords", sKeywords)
            .Parameters.AddWithValue("@Language1", sLanguage1)
            .Parameters.AddWithValue("@Location", sLocation)
            .Parameters.AddWithValue("@ManagersName", sManagersName)
            .Parameters.AddWithValue("@Mileage", sMileage)
            .Parameters.AddWithValue("@Notes", sNotes)
            .Parameters.AddWithValue("@OfficeLocation", sOfficeLocation)
            .Parameters.AddWithValue("@OrganizationalIDNumber", sOrganizationalIDNumber)
            .Parameters.AddWithValue("@OtherAddressPOBox", sOtherAddressPOBox)
            .Parameters.AddWithValue("@Priority", sPriority)
            .Parameters.AddWithValue("@Private", sPrivate)
            .Parameters.AddWithValue("@Profession", sProfession)
            .Parameters.AddWithValue("@ReferredBy", sReferredBy)
            .Parameters.AddWithValue("@Sensitivity", sSensitivity)
            .Parameters.AddWithValue("@Spouse", sSpouse)
            .Parameters.AddWithValue("@User1", sUser1)
            .Parameters.AddWithValue("@User2", sUser2)
            .Parameters.AddWithValue("@User3", sUser3)
            .Parameters.AddWithValue("@User4", sUser4)
            .Parameters.AddWithValue("@WebPage", sWebPage)
            .Parameters.AddWithValue("@Supervisor", sSupervisor)
            .Parameters.AddWithValue("@SupervisorPhone", sSupervisorPhone)
            .Parameters.AddWithValue("@SupervisorEmail", sSupervisorEmail)
            .Parameters.AddWithValue("@SupervisorAssistant", sSupervisorAssistant)
            .Parameters.AddWithValue("@SupervisorAssistantPhone", sSupervisorAssistantPhone)
            .Parameters.AddWithValue("@SupervisorAssistantEmail", sSupervisorAssistantEmail)
            .Parameters.AddWithValue("@DepartmentEscalationsContact", sDepartmentEscalationsContact)
            .Parameters.AddWithValue("@DepartmentEscalationsContactNumber", sDepartmentEscalationsContactNumber)
            .Parameters.AddWithValue("@DepartmentEscalationsEmail", sDepartmentEscalationsEmail)
            .Parameters.AddWithValue("@CurrentEscalationsContact", sCurrentEscalationsContact)
            .Parameters.AddWithValue("@CurrentEscalationsContactPhoneNumber", sCurrentEscalationsContactPhoneNumber)
            .Parameters.AddWithValue("@CurrentEscalationsEmail", sCurrentEscalationsEmail)
            .Parameters.AddWithValue("@CurrentEscalationDateFrom", sCurrentEscalationDateFrom)
            .Parameters.AddWithValue("@CurrentEscalationDateTo", sCurrentEscalationDateTo)
            .Parameters.AddWithValue("@SecondaryEscalationsContact", sSecondaryEscalationsContact)
            .Parameters.AddWithValue("@SecondaryEscalationsContactPhoneNumber", sSecondaryEscalationsContactPhoneNumber)
            .Parameters.AddWithValue("@SecondaryEscalationsEmail", sSecondaryEscalationsEmail)
            .Parameters.AddWithValue("@SecondaryEscalationsDateFrom", sSecondaryEscalationsDateFrom)
            .Parameters.AddWithValue("@SecondaryEscalationsDateTo", sSecondaryEscalationsDateTo)
            .Parameters.AddWithValue("@TemporaryForwardingEmailAddress1", sTemporaryForwardingEmailAddress1)
            .Parameters.AddWithValue("@TemporaryForwardingAddress1FromDate", sTemporaryForwardingAddress1FromDate)
            .Parameters.AddWithValue("@TemporaryForwardingAddress1ToDate", sTemporaryForwardingAddress1ToDate)
            .Parameters.AddWithValue("@TemporaryForwardingEmailAddress2", sTemporaryForwardingEmailAddress2)
            .Parameters.AddWithValue("@TemporaryForwardingAddress2FromDate", sTemporaryForwardingAddress2FromDate)
            .Parameters.AddWithValue("@TemporaryForwardingAddress2ToDate", sTemporaryForwardingAddress2ToDate)
            .Parameters.AddWithValue("@BestWaytoContactDuringBusinessHours", sBestWaytoContactDuringBusinessHours)
            .Parameters.AddWithValue("@BestWaytoContactAfterBusinessHours", sBestWaytoContactAfterBusinessHours)
            .Parameters.AddWithValue("@ContactInformationNotes", sContactInformationNotes)
            .Parameters.AddWithValue("@EscalationInformationNotes", sEscalationInformationNotes)
            .Parameters.AddWithValue("@InCaseofEmergencyContactInformationName", sInCaseofEmergencyContactInformationName)
            .Parameters.AddWithValue("@InCaseofEmergencyContactRelationship", sInCaseofEmergencyContactRelationship)
            .Parameters.AddWithValue("@InCaseofEmergencyContactInformationPhoneNumber", sInCaseofEmergencyContactInformationPhoneNumber)
            .Parameters.AddWithValue("@InCaseOfEmergencyContactInformationEmailAddress", sInCaseOfEmergencyContactInformationEmailAddress)
            .Parameters.AddWithValue("@ContactPriority1", sContactPriority1)
            .Parameters.AddWithValue("@ContactPriority2", sContactPriority2)
            .Parameters.AddWithValue("@ContactPriority3", sContactPriority3)
            .Parameters.AddWithValue("@FileAs", sFileAs)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerType)
            .Parameters.AddWithValue("@DispatcherInformationNotes", sDispatcherInfo)
        End With

        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        SaveAddressBookContact = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadDomainNames(ByVal sAppUserId As String, Optional ByVal bActiveOnly As Boolean = False) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDomainNamesSelect"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@ActiveOnly", bActiveOnly)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadDomainNames = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAddressBookEmails(ByVal sAppCode As String, ByVal sAddressBookTypeCode As String, ByVal sAppUserID As String, ByVal sSearchTerm As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AddressBookEmails"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", "")
            .Parameters.AddWithValue("@DomainName", "")
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@BeginRecordNumber", 0)
            .Parameters.AddWithValue("@EndRecordNumber", 5)
            .Parameters.AddWithValue("@sortColumns", "EmailAddress")
            .Parameters.AddWithValue("@ShowPagedResults", False)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookEmails = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAppCodeHelp(ByVal sAppCode As String, ByVal sDeviceTypeCode As String, ByVal sHelpTypeCode As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppCodeHelp"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@HelpTypeCode", sHelpTypeCode)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAppCodeHelp = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAppCodeHelpDetails(ByVal sAppCode As String, ByVal sDeviceTypeCode As String, ByVal sHelpTypeCode As String, ByVal sParentID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppCodeHelpDetails"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@HelpTypeCode", sHelpTypeCode)
            .Parameters.AddWithValue("@ParentID", sParentID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAppCodeHelpDetails = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadHelpTypes() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminHelpTypes"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadHelpTypes = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadHelpParents(sAppCode As String, sDeviceTypeCode As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminHelpParents"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadHelpParents = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function ImportFromMPA(sAppCode As String, sFirstName As String, sLastName As String, sCompanyAppGroupID As String, sBusinessEmailAddress As String, sPagerNumber As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportFromMPA"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@CompanyAppGroupID", sCompanyAppGroupID)
            .Parameters.AddWithValue("@BusinessEmailAddress", sBusinessEmailAddress)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        ImportFromMPA = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAddressBookContactQuick(ByVal sActionType As String, sAppCode As String, sAddressBookTypeCode As String, sAppUserID As String, sAppGroupId As String, sAddressBookID As String, sTitle As String, sFirstName As String, sMiddleName As String, sLastName As String, sSuffix As String, sEmailAddress As String, sEmailDisplayName As String, ByVal sDepartment As String, ByVal sJobTitle As String, sBusinessPhone As String, sHomePhone As String, sMobilePhone As String, sPager As String, sManagersName As String, sNotes As String, sBestWaytoContactDuringBusinessHours As String, sBestWaytoContactAfterBusinessHours As String, sContactInformationNotes As String, sEscalationInformationNotes As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookContactQuick"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
            .Parameters.AddWithValue("@Title", sTitle)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@MiddleName", sMiddleName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@Suffix", sSuffix)
            .Parameters.AddWithValue("@EmailAddress", sEmailAddress)
            .Parameters.AddWithValue("@EmailDisplayName", sEmailDisplayName)
            .Parameters.AddWithValue("@Department", sDepartment)
            .Parameters.AddWithValue("@JobTitle", sJobTitle)
            .Parameters.AddWithValue("@BusinessPhone", sBusinessPhone)
            .Parameters.AddWithValue("@HomePhone", sHomePhone)
            .Parameters.AddWithValue("@MobilePhone", sMobilePhone)
            .Parameters.AddWithValue("@Pager", sPager)
            .Parameters.AddWithValue("@ManagersName", sManagersName)
            .Parameters.AddWithValue("@Notes", sNotes)
            .Parameters.AddWithValue("@BestWaytoContactDuringBusinessHours", sBestWaytoContactDuringBusinessHours)
            .Parameters.AddWithValue("@BestWaytoContactAfterBusinessHours", sBestWaytoContactAfterBusinessHours)
            .Parameters.AddWithValue("@ContactInformationNotes", sContactInformationNotes)
            .Parameters.AddWithValue("@EscalationInformationNotes", sEscalationInformationNotes)

        End With

        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        SaveAddressBookContactQuick = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetUserDeviceDetail(ByVal AppCode As String, ByVal AppUserID As String, ByVal LoggedInUser As String, ByVal DeviceID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUserDeviceDetails"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@LoggedInUser", LoggedInUser)
            .Parameters.AddWithValue("@DeviceID", DeviceID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserDeviceDetail = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAppCodeHelpDemo(ByVal sAppCode As String, ByVal sDeviceTypeCode As String, ByVal sHelpTypeCode As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "GetAppCodeHelp"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@HelpTypeCode", sHelpTypeCode)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAppCodeHelpDemo = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAppCodeHelpDetailsDemo(ByVal sAppCode As String, ByVal sDeviceTypeCode As String, ByVal sHelpTypeCode As String, ByVal sParentID As String, ByVal sSearchTerm As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "GetAppCodeHelpDetails"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@HelpTypeCode", sHelpTypeCode)
            .Parameters.AddWithValue("@ParentID", sParentID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAppCodeHelpDetailsDemo = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAppCodeHelp(ByVal sActionType As String, ByVal sAppCode As String, ByVal sDeviceTypeCode As String, ByVal sHelpID As String, ByVal sParentID As String, ByVal sHelpTypeCode As String, ByVal sHelpText As String, ByVal sSeqNo As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppCodeHelp"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@DeviceTypeCode", sDeviceTypeCode)
            .Parameters.AddWithValue("@HelpID", sHelpID)
            .Parameters.AddWithValue("@ParentID", sParentID)
            .Parameters.AddWithValue("@HelpTypeCode", sHelpTypeCode)
            .Parameters.AddWithValue("@HelpText", sHelpText)
            .Parameters.AddWithValue("@SequenceNo", sSeqNo)
            .Parameters.AddWithValue("@ReturnMessage", "")
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetLastMessageNumber(ByVal AppCode As String, ByVal AppUserID As String) As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetLastMessageNumber"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetMessageHistoryChangeStatus(ByVal AppCode As String, ByVal AppUserID As String, ByVal TimerMinutes As Long) As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMessageHistoryChangeStatus"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetDeviceType(ByVal sAppCode As String, ByVal sAppUserId As String, ByVal sDeviceId As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "GetDeviceType"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@DeviceId", sDeviceId)

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

        End Try

        Return retValue

    End Function

    Public Shared Function GetGroupDomainName(ByVal sAppGroupID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppGroupDomainName"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function ResetBilling(ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ResetBilling"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function AddressBookCheckExistingEntry(AppCode As String, AddressBookTypeCode As String, AddressBookID As String, EmailAddress As String, AppUserID As String, AppGroupID As String, DomainName As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AddressBookCheckExistingEntry"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", AddressBookTypeCode)
            .Parameters.AddWithValue("@AddressBookID", AddressBookID)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@DomainName", DomainName)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetGroupTimeZone(AppGroupID As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetGroupTimeZone"
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadAppUsersAC(Optional ByVal sAppUserId As String = "", Optional ByVal bAdminOnly As Boolean = False, Optional ByVal sSearchTerm As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppUsersAC"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AdminOnly", bAdminOnly)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAppUsersAC = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAppUsersAC2(Optional ByVal sAppUserId As String = "", Optional ByVal bAdminOnly As Boolean = False, Optional ByVal sSearchTerm As String = "", Optional ByVal sAppGroupID As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppUsersAC2"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AdminOnly", bAdminOnly)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAppUsersAC2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAddressBookSubjects(ByVal sAppUserID As String, ByVal sSearchTerm As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AddressBookSubjects"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookSubjects = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAddressBookMessageBody(ByVal sAppUserID As String, ByVal sSearchTerm As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AddressBookBody"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookMessageBody = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function CanDeleteUser(AppUserID As String) As Boolean
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCanDeleteAppUser"
            .Parameters.AddWithValue("@AppUserID", AppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function DeleteUser(AppUserID As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeleteAppUser"
            .Parameters.AddWithValue("@AppUserID", AppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveUserPreferences(AppUserID As String, ModuleCode As String, PreferenceCode As String, PreferenceNotes As String, TrueFalseParameter As Boolean, CharacterParameter As String, NumericParameter As Long) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveUserPreferences"
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@ModuleCode", ModuleCode)
            .Parameters.AddWithValue("@PreferenceCode", PreferenceCode)
            .Parameters.AddWithValue("@PreferenceNotes", PreferenceNotes)
            .Parameters.AddWithValue("@TrueFalseParameter", TrueFalseParameter)
            .Parameters.AddWithValue("@CharacterParameter", CharacterParameter)
            .Parameters.AddWithValue("@NumericParameter", NumericParameter)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetUserPreferences(ByVal AppUserID As String, ByVal ModuleCode As String, ByVal PreferenceCode As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserPreferences"
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@ModuleCode", ModuleCode)
            .Parameters.AddWithValue("@PreferenceCode", PreferenceCode)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserPreferences = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetUserPreferencesTrueFalse(ByVal AppUserID As String, ByVal ModuleCode As String, ByVal PreferenceCode As String) As Boolean
        Dim retValue As Boolean = False
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserPreferencesTrueFalse"
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@ModuleCode", ModuleCode)
            .Parameters.AddWithValue("@PreferenceCode", PreferenceCode)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetUserPreferencesCharacter(ByVal AppUserID As String, ByVal ModuleCode As String, ByVal PreferenceCode As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserPreferencesCharacter"
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@ModuleCode", ModuleCode)
            .Parameters.AddWithValue("@PreferenceCode", PreferenceCode)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetUserPreferencesNumeric(ByVal AppUserID As String, ByVal ModuleCode As String, ByVal PreferenceCode As String) As Double
        Dim retValue As Double = 0
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserPreferencesNumeric"
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@ModuleCode", ModuleCode)
            .Parameters.AddWithValue("@PreferenceCode", PreferenceCode)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetGroupQuickMessage(ByVal AppCode As String, ByVal AppGroupID As String, ByVal QuickMessageID As Long) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetGroupQuickMessage"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupId", AppGroupID)
            .Parameters.AddWithValue("@QuickMessageID", QuickMessageID)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetGroupQuickMessage = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadQuickMessageButtonSelect(ByVal AppGroupId As String, ByVal AppUserId As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminQuickMessageButtonSelect"
            .Parameters.AddWithValue("@AppGroupId", AppGroupId)
            .Parameters.AddWithValue("@AppUserId", AppUserId)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadQuickMessageButtonSelect = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function GetUserQuickMessage(ByVal AppCode As String, ByVal AppGroupID As String, ByVal AppUserID As String, ByVal QuickMessageID As Long) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetQuickMessage"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupId", AppGroupID)
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@QuickMessageID", QuickMessageID)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserQuickMessage = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAppQuickMessage(ByVal sActionType As String, ByVal sAppCode As String, ByVal sAppGroupID As String, ByVal sAppUserId As String, ByVal sQuickMessageID As String, ByVal sQuickMessageDescription As String, sMessageTo As String, sMessageSubject As String, sMessageBody As String, sButtonColor As String, sButtonText As String, ByVal bActive As Boolean) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppQuickMessage"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@QuickMessageID", sQuickMessageID)
            .Parameters.AddWithValue("@QuickMessageDescription", sQuickMessageDescription)
            .Parameters.AddWithValue("@MessageTo", sMessageTo)
            .Parameters.AddWithValue("@MessageSubject", sMessageSubject)
            .Parameters.AddWithValue("@MessageBody", sMessageBody)
            .Parameters.AddWithValue("@ButtonColor", sButtonColor)
            .Parameters.AddWithValue("@ButtonText", sButtonText)
            .Parameters.AddWithValue("@Active", bActive)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        SaveAppQuickMessage = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetUserCurrentStatus(ByVal sAppUserID As String, Optional ByVal bIncludeForwardTo As Boolean = False) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "GetUserCurrentStatus"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@IncludeForwardTo", bIncludeForwardTo)
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

        End Try

        Return retValue

    End Function


    Public Shared Function GetGroupCanned(ByVal AppCode As String, ByVal AppGroupID As String, ByVal CannedID As Long) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetGroupCanned"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupId", AppGroupID)
            .Parameters.AddWithValue("@CannedID", CannedID)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetGroupCanned = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadCannedTextSelect(ByVal AppGroupId As String, ByVal AppUserId As String, ByVal CannedType As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCannedTextSelect"
            .Parameters.AddWithValue("@AppGroupId", AppGroupId)
            .Parameters.AddWithValue("@AppUserId", AppUserId)
            .Parameters.AddWithValue("@CannedType", CannedType)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadCannedTextSelect = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function GetUserCanned(ByVal AppCode As String, ByVal AppGroupID As String, ByVal AppUserID As String, ByVal CannedID As Long) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetCanned"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupId", AppGroupID)
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@CannedID", CannedID)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserCanned = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAppCanned(ByVal sActionType As String, ByVal sAppCode As String, ByVal sAppGroupID As String, ByVal sAppUserId As String, ByVal sCannedID As String, ByVal sCannedType As String, sCannedDescription As String, sCannedText As String, ByVal bActive As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppCanned"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@CannedID", sCannedID)
            .Parameters.AddWithValue("@CannedType", sCannedType)
            .Parameters.AddWithValue("@CannedDescription", sCannedDescription)
            .Parameters.AddWithValue("@CannedText", sCannedText)
            .Parameters.AddWithValue("@Active", bActive)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveNewGroup(ByVal sAppCode As String, ByVal sAppUserId As String, ByVal sAppGroupID As String, ByVal sNewGroupName As String, ByVal sAppGroupType As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveNewGroup"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@NewGroupName", sNewGroupName)
            .Parameters.AddWithValue("@AppGroupType", sAppGroupType)

        End With

        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)
        SaveNewGroup = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadGroupTypeSelect() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppGroupTypes"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadGroupTypeSelect = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetNewGroupDomain(ByVal sAppGroupID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetNewGroupDomain"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadGroups(ByVal sAppUserId As String, ByVal bActiveOnly As Boolean, ByVal sAppGroupType As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGroupsSelect"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@ActiveOnly", bActiveOnly)
            .Parameters.AddWithValue("@AppGroupType", sAppGroupType)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadGroups = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveGroup(ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sAppGroupType As String, ByVal sAppGroupDescription As String, ByVal sName As String, ByVal bActive As Boolean, ByVal bOnlyAllowMessagesFromWithinGroup As Boolean, ByVal bAllowMessageFromPrimaryEmail As Boolean, bAllowReplyOnGroupMessage As Boolean, sGroupMessageFromIndividualOrGroup As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveGroup"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@AppGroupType", sAppGroupType)
            .Parameters.AddWithValue("@AppGroupDescription", sAppGroupDescription)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@OnlyAllowMessagesFromWithinGroup", bOnlyAllowMessagesFromWithinGroup)
            .Parameters.AddWithValue("@AllowMessageFromPrimaryEmail", bAllowMessageFromPrimaryEmail)
            .Parameters.AddWithValue("@AllowReplyOnGroupMessage", bAllowReplyOnGroupMessage)
            .Parameters.AddWithValue("@GroupMessageFromIndividualOrGroup", sGroupMessageFromIndividualOrGroup)

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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveGroup2(ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sAppGroupType As String, ByVal sAppGroupDescription As String, ByVal sName As String, ByVal bActive As Boolean, ByVal bOnlyAllowMessagesFromWithinGroup As Boolean, ByVal bAllowMessageFromPrimaryEmail As Boolean, bAllowReplyOnGroupMessage As Boolean, sGroupMessageFromIndividualOrGroup As String, bSendMessageDetail As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveGroup2"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@AppGroupType", sAppGroupType)
            .Parameters.AddWithValue("@AppGroupDescription", sAppGroupDescription)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@OnlyAllowMessagesFromWithinGroup", bOnlyAllowMessagesFromWithinGroup)
            .Parameters.AddWithValue("@AllowMessageFromPrimaryEmail", bAllowMessageFromPrimaryEmail)
            .Parameters.AddWithValue("@AllowReplyOnGroupMessage", bAllowReplyOnGroupMessage)
            .Parameters.AddWithValue("@GroupMessageFromIndividualOrGroup", sGroupMessageFromIndividualOrGroup)
            .Parameters.AddWithValue("@SendMessageDetail", bSendMessageDetail)

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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadGroupMembers(ByVal sAppUserId As String, ByVal sAppGroupID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGroupMembers"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadGroupMembers = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAvailableUsers(Optional ByVal sAppUserId As String = "", Optional ByVal bAdminOnly As Boolean = False, Optional ByVal sAppGroupID As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAvailableUsers"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AdminOnly", bAdminOnly)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAvailableUsers = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function AddGroupMember(ByVal sMemberType As String, ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sAppCode As String, ByVal sAddressBookTypeCode As String, ByVal sAddressBookID As String, ByVal EmailAddressTypeCode As String, ByVal EmailAddress As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAddGroupMember"
            .Parameters.AddWithValue("@MemberType", sMemberType)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
            .Parameters.AddWithValue("@EmailAddressTypeCode", EmailAddressTypeCode)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
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

        End Try

        Return retValue

    End Function

    Public Shared Function RemoveGroupMember(ByVal sMemberType As String, ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sAppCode As String, ByVal sAddressBookTypeCode As String, ByVal sAddressBookID As String, ByVal EmailAddressTypeCode As String, ByVal EmailAddress As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminRemoveGroupMember"
            .Parameters.AddWithValue("@MemberType", sMemberType)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
            .Parameters.AddWithValue("@EmailAddressTypeCode", EmailAddressTypeCode)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadPagerTypes() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminPagerType"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadPagerTypes = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetGroupEmailAddress(ByVal sAppGroupID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetGroupEmailAddress"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetGroupMessageHistoryMessage(ByVal AppUserID As String, ByVal MessageNumber As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetGroupMessageHistoryMessage"
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@MessageNumber", MessageNumber)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetGroupMessageHistoryMessage = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveUserStatus(ByVal AppUserId As String, ByVal StatusLocationCode As String, ByVal StatusAvailabilityCode As String, ByVal ForwardTo As String, ByVal StatusSet As String, ByVal StatusEnds As String, ByVal CustomStatus As String, ByVal StatusReason As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveUserStatus"
            .Parameters.AddWithValue("@AppUserId", AppUserId)
            .Parameters.AddWithValue("@StatusLocationCode", StatusLocationCode)
            .Parameters.AddWithValue("@StatusAvailabilityCode", StatusAvailabilityCode)
            .Parameters.AddWithValue("@ForwardTo", ForwardTo)
            .Parameters.AddWithValue("@StatusSet", StatusSet)
            .Parameters.AddWithValue("@StatusEnds", StatusEnds)
            .Parameters.AddWithValue("@CustomStatus", CustomStatus)
            .Parameters.AddWithValue("@StatusReason", StatusReason)

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

        End Try

        Return retValue

    End Function

    Public Shared Function GetUserStatusInfo(ByVal AppUserID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserStatus"
            .Parameters.AddWithValue("@AppUserID", AppUserID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserStatusInfo = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try


    End Function

    Public Shared Function GetPageBreadcrumbs(ByVal AppCode As String, ByVal SecurityGroupName As String, ByVal AppScreenName As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetPageBreadcrumbs"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@SecurityGroupName", SecurityGroupName)
            .Parameters.AddWithValue("@AppScreenName", AppScreenName)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetPageBreadcrumbs = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetGroupMessageOriginatorsList(ByVal AppGroupID As String, Optional ByVal Status As String = "") As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGroupMessageOriginators"
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@Status", Status)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetGroupMessageOriginatorsList = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try


    End Function

    Public Shared Function SaveGroupMessageOriginator(ByVal AppGroupID As String, ByVal EmailAddress As String, ByVal CanSendGroupMessage As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveGroupMessageOriginator"
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
            .Parameters.AddWithValue("@CanSendGroupMessage", CanSendGroupMessage)

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

        End Try

        Return retValue

    End Function

    Public Shared Function GetIsUserSubscriber(ByVal AppUserID As String) As Boolean
        Dim retValue As Boolean
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim mycommon As New common

        AppUserID = mycommon.StripDisplayName(AppUserID)

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminIsUserSubscriber"
            .Parameters.AddWithValue("@AppUserID", AppUserID)

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

        End Try

        Return retValue

    End Function


    Public Shared Function GetCannedText(ByVal AppCode As String, ByVal AppGroupID As String, ByVal AppUserID As String, ByVal CannedType As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetCannedText"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@CannedType", CannedType)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetCannedText = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function CheckAddressExists(ByVal sAppUserID As String, ByVal sEmailAddress As String) As Boolean
        Dim retValue As Boolean
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCheckAddressExists"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@EmailAddress", sEmailAddress)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadToEmails(ByVal sAppUserID As String, ByVal sSearchTerm As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetToEmails"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadToEmails = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try


    End Function

    Public Shared Function AddGroupMemberAuto(ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sAppCode As String, ByVal EmailAddress As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim mycommon As New common

        EmailAddress = mycommon.StripDisplayName(EmailAddress)


        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAddGroupMemberAuto"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
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

        End Try

        Return retValue

    End Function

    Public Shared Function AddQuickMessageMemberAuto(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sQuickMessageID As String, ByVal EmailAddress As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAddQuickMessageMemberAuto"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@QuickMessageID", sQuickMessageID)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
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

        End Try

        Return retValue

    End Function


    Public Shared Function GetUserQuickMessageEmails(ByVal AppCode As String, ByVal AppGroupID As String, ByVal AppUserID As String, ByVal QuickMessageID As Long) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetQuickMessageEmails"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupId", AppGroupID)
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@QuickMessageID", QuickMessageID)

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

        End Try

        Return retValue

    End Function

    Public Shared Function ClearAllCanned(ByVal AppUserID As String) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminClearCannedPreferences"
            .Parameters.AddWithValue("@AppUserId", AppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetAppUserPassword(ByVal sAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "PasswordRetrieval"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadClientIDFormats() As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminClientIDFormats"
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadClientIDFormats = dt

        If Not cmdSQL Is Nothing Then
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End If

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try


    End Function

    Public Shared Function GetNewUserID(ByVal sAppGroupID As String, ByVal sFirstName As String, ByVal sLastName As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCreateNewUserID"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
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

        End Try

        Return retValue

    End Function

    Public Shared Function IsValidAllowedUser(ByVal sLoggedInAppUserID As String, ByVal sAppUserID As String) As Boolean
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminIsValidAllowedUser"
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUserID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
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

        End Try

        Return retValue

    End Function


    Public Shared Function GetPagerError(ByVal sPagerNumbers As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetPagerError"
            .Parameters.AddWithValue("@PagerNumbers", sPagerNumbers)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetPagerError = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try


    End Function

    Public Shared Function GetPagerType(ByVal sPhoneNo As String) As Boolean
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetPagerType"
            .Parameters.AddWithValue("@PhoneNo", sPhoneNo)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetFirstUserInGroup(ByVal sAppGroupID As String) As String
        Dim retValue As String = False
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetFirstUserInGroup"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetFirstUserInGroup2(ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String = False
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetFirstUserInGroup2"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetOpenMessageCount(ByVal sAppUserID As String) As String
        Dim retValue As String = False
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminOpenMessageCount"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
        End With

        Try
            cn.Open()
            retValue = cmdSQL.ExecuteScalar()


        Catch ex As Exception
            retValue = "0"
        End Try

        GetOpenMessageCount = retValue


       If Not cmdSQL Is Nothing Then
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End If

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try


    End Function

    Public Shared Function DeactivateDevices(ByVal sAppCode As String, ByVal sAppUserID As String) As String

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeactivateAppUserDevices"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetNextMessage(ByVal AppCode As String, ByVal AppUserID As String, ByVal DeviceID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetNextMessage"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@DeviceId", DeviceID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetNextMessage = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SendCannedReply(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String, ByVal lMessageNumber As Long, ByVal lCannedReplyNumber As Long) As String

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSendCannedReply"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
            .Parameters.AddWithValue("@MessageNumber", lMessageNumber)
            .Parameters.AddWithValue("@CannedReplyNumber", lCannedReplyNumber)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SendReadReceipt(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sDeviceID As String, ByVal lMessageNumber As Long) As String

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSendReadReceipt"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@DeviceID", sDeviceID)
            .Parameters.AddWithValue("@MessageNumber", lMessageNumber)
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

        End Try

        Return retValue

    End Function

    Public Shared Function ConsultUpdateConversation(ByVal sConsultNumber As String, ByVal sAppGroupID As String, ByVal sSourceEmailAddress As String, ByVal sMessageBody As String) As String

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminConsultUpdateConversation"
            .Parameters.AddWithValue("@ConsultNumber", sConsultNumber)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@SourceEmailAddress", sSourceEmailAddress)
            .Parameters.AddWithValue("@MessageBody", sMessageBody)
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

        End Try

        Return retValue

    End Function

    Public Shared Function ForcePasswordChange(ByVal ForceType As String, ByVal LoggedInAppUserId As String, ByVal AppGroupId As String, ByVal AppUserId As String, ByVal SystemForcePasswordDays As Long, ByVal GroupForcePasswordDays As Long, ByVal ForceChangeNow As Boolean) As String

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminForcePasswordChange"
            .Parameters.AddWithValue("@ForceType", ForceType)
            .Parameters.AddWithValue("@LoggedInAppUserId", LoggedInAppUserId)
            .Parameters.AddWithValue("@AppGroupId", AppGroupId)
            .Parameters.AddWithValue("@AppUserId", AppUserId)

            .Parameters.AddWithValue("@SystemForcePasswordDays", SystemForcePasswordDays)
            .Parameters.AddWithValue("@GroupForcePasswordDays", GroupForcePasswordDays)
            .Parameters.AddWithValue("@ForceChangeNow", ForceChangeNow)
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

        End Try

        Return retValue

    End Function

    Public Shared Function AdminChangePasswordAppUser2(ByVal sAppUserID As String, ByVal sNewPassword As String, ByVal sLoggedInUser As String) As String
        Dim retValue As String
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

        If Not cmdSQL Is Nothing Then
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End If

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

        Return retValue

    End Function

    Public Shared Function GetUserPassword(ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserPassword"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAppUserSSOUUID(ByVal sAppUserID As String, ByVal sSSOUUID As String) As String
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

        End Try

        Return retValue
    End Function

    Public Shared Function GetCannedTextToSend(ByVal AppCode As String, ByVal AppGroupID As String, ByVal AppUserID As String, ByVal CannedID As String) As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String = ""

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetCannedTextToSend"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@CannedID", CannedID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAppUserFromSSO(ByVal sAppUserID As String, ByVal sPrimaryEmail As String, sFirstName As String, sLastName As String, ByVal sActive As String, ByVal sSSOUUID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        If LCase(Trim(sActive)) = "true" Or LCase(Trim(sActive)) = "1" Then
            sActive = "1"
        Else
            sActive = "0"
        End If

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUpdateAppUserFromSSO"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@Active", sActive)
            .Parameters.AddWithValue("@SSOUUID", sSSOUUID)
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

        End Try

        Return retValue
    End Function

    Public Shared Function IsGroupAllowed(ByVal sAppUserID As String, ByVal sAppGroupID As String) As Boolean
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppGroupsAllow"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
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

        End Try

        Return retValue
    End Function

    Public Shared Function GetApplicationControl(ByVal sCompanyCode As String, ByVal sModuleCode As String, ByVal sApplicationControlCode As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetSSOUUID(ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUserSSOUUID"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
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

        End Try

        Return retValue
    End Function

    Public Shared Function SaveImportAddressBook(ActionType As String, AppCode As String, ImportAppUserId As String, LineNumber As Long, EmailAddress As String, EmailType As String, EmailDisplayName As String, AddressBookTypeCode As String, DomainName As String, AppGroupID As String, AppUserID As String, Title As String, FirstName As String, MiddleName As String, LastName As String, Suffix As String, Company As String, Department As String, JobTitle As String, BusinessStreet As String, BusinessStreet2 As String, BusinessStreet3 As String, BusinessCity As String, BusinessState As String, BusinessPostalCode As String, BusinessCountryRegion As String, HomeStreet As String, HomeStreet2 As String, HomeStreet3 As String, HomeCity As String, HomeState As String, HomePostalCode As String, HomeCountryRegion As String, OtherStreet As String, OtherStreet2 As String, OtherStreet3 As String, OtherCity As String, OtherState As String, OtherPostalCode As String, OtherCountryRegion As String, AssistantsPhone As String, BusinessFax As String, BusinessPhone As String, BusinessPhone2 As String, Callback As String, CarPhone As String, CompanyMainPhone As String, HomeFax As String, HomePhone As String, HomePhone2 As String, ISDN As String, MobilePhone As String, OtherFax As String, OtherPhone As String, Pager As String, PrimaryPhone As String, RadioPhone As String, TTYTDDPhone As String, Telex As String, Account As String, Anniversary As String, AssistantsName As String, BillingInformation As String, Birthday As String, BusinessAddressPOBox As String, Categories As String, Children As String, DirectoryServer As String, Email2Address As String, Email2Type As String, Email2DisplayName As String, Email3Address As String, Email3Type As String, Email3DisplayName As String, Gender As String, GovernmentIDNumber As String, Hobby As String, HomeAddressPOBox As String, Initials As String, InternetFreeBusy As String, Keywords As String, Language1 As String, Location As String, ManagersName As String, Mileage As String, Notes As String, OfficeLocation As String, OrganizationalIDNumber As String, OtherAddressPOBox As String, Priority As String, sPrivate As String, Profession As String, ReferredBy As String, Sensitivity As String, Spouse As String, User1 As String, User2 As String, User3 As String, User4 As String, WebPage As String, PagerTypeCode As String, DoNotImport As Boolean) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveImportAddressBook"
            .Parameters.AddWithValue("@ActionType", ActionType)
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@ImportAppUserId", ImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", LineNumber)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
            .Parameters.AddWithValue("@EmailType", EmailType)
            .Parameters.AddWithValue("@EmailDisplayName", EmailDisplayName)
            .Parameters.AddWithValue("@AddressBookTypeCode", AddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", DomainName)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@Title", Title)
            .Parameters.AddWithValue("@FirstName", FirstName)
            .Parameters.AddWithValue("@MiddleName", MiddleName)
            .Parameters.AddWithValue("@LastName", LastName)
            .Parameters.AddWithValue("@Suffix", Suffix)
            .Parameters.AddWithValue("@Company", Company)
            .Parameters.AddWithValue("@Department", Department)
            .Parameters.AddWithValue("@JobTitle", JobTitle)
            .Parameters.AddWithValue("@BusinessStreet", BusinessStreet)
            .Parameters.AddWithValue("@BusinessStreet2", BusinessStreet2)
            .Parameters.AddWithValue("@BusinessStreet3", BusinessStreet3)
            .Parameters.AddWithValue("@BusinessCity", BusinessCity)
            .Parameters.AddWithValue("@BusinessState", BusinessState)
            .Parameters.AddWithValue("@BusinessPostalCode", BusinessPostalCode)
            .Parameters.AddWithValue("@BusinessCountryRegion", BusinessCountryRegion)
            .Parameters.AddWithValue("@HomeStreet", HomeStreet)
            .Parameters.AddWithValue("@HomeStreet2", HomeStreet2)
            .Parameters.AddWithValue("@HomeStreet3", HomeStreet3)
            .Parameters.AddWithValue("@HomeCity", HomeCity)
            .Parameters.AddWithValue("@HomeState", HomeState)
            .Parameters.AddWithValue("@HomePostalCode", HomePostalCode)
            .Parameters.AddWithValue("@HomeCountryRegion", HomeCountryRegion)
            .Parameters.AddWithValue("@OtherStreet", OtherStreet)
            .Parameters.AddWithValue("@OtherStreet2", OtherStreet2)
            .Parameters.AddWithValue("@OtherStreet3", OtherStreet3)
            .Parameters.AddWithValue("@OtherCity", OtherCity)
            .Parameters.AddWithValue("@OtherState", OtherState)
            .Parameters.AddWithValue("@OtherPostalCode", OtherPostalCode)
            .Parameters.AddWithValue("@OtherCountryRegion", OtherCountryRegion)
            .Parameters.AddWithValue("@AssistantsPhone", AssistantsPhone)
            .Parameters.AddWithValue("@BusinessFax", BusinessFax)
            .Parameters.AddWithValue("@BusinessPhone", BusinessPhone)
            .Parameters.AddWithValue("@BusinessPhone2", BusinessPhone2)
            .Parameters.AddWithValue("@Callback", Callback)
            .Parameters.AddWithValue("@CarPhone", CarPhone)
            .Parameters.AddWithValue("@CompanyMainPhone", CompanyMainPhone)
            .Parameters.AddWithValue("@HomeFax", HomeFax)
            .Parameters.AddWithValue("@HomePhone", HomePhone)
            .Parameters.AddWithValue("@HomePhone2", HomePhone2)
            .Parameters.AddWithValue("@ISDN", ISDN)
            .Parameters.AddWithValue("@MobilePhone", MobilePhone)
            .Parameters.AddWithValue("@OtherFax", OtherFax)
            .Parameters.AddWithValue("@OtherPhone", OtherPhone)
            .Parameters.AddWithValue("@Pager", Pager)
            .Parameters.AddWithValue("@PrimaryPhone", PrimaryPhone)
            .Parameters.AddWithValue("@RadioPhone", RadioPhone)
            .Parameters.AddWithValue("@TTYTDDPhone", TTYTDDPhone)
            .Parameters.AddWithValue("@Telex", Telex)
            .Parameters.AddWithValue("@Account", Account)
            .Parameters.AddWithValue("@Anniversary", Anniversary)
            .Parameters.AddWithValue("@AssistantsName", AssistantsName)
            .Parameters.AddWithValue("@BillingInformation", BillingInformation)
            .Parameters.AddWithValue("@Birthday", Birthday)
            .Parameters.AddWithValue("@BusinessAddressPOBox", BusinessAddressPOBox)
            .Parameters.AddWithValue("@Categories", Categories)
            .Parameters.AddWithValue("@Children", Children)
            .Parameters.AddWithValue("@DirectoryServer", DirectoryServer)
            .Parameters.AddWithValue("@Email2Address", Email2Address)
            .Parameters.AddWithValue("@Email2Type", Email2Type)
            .Parameters.AddWithValue("@Email2DisplayName", Email2DisplayName)
            .Parameters.AddWithValue("@Email3Address", Email3Address)
            .Parameters.AddWithValue("@Email3Type", Email3Type)
            .Parameters.AddWithValue("@Email3DisplayName", Email3DisplayName)
            .Parameters.AddWithValue("@Gender", Gender)
            .Parameters.AddWithValue("@GovernmentIDNumber", GovernmentIDNumber)
            .Parameters.AddWithValue("@Hobby", Hobby)
            .Parameters.AddWithValue("@HomeAddressPOBox", HomeAddressPOBox)
            .Parameters.AddWithValue("@Initials", Initials)
            .Parameters.AddWithValue("@InternetFreeBusy", InternetFreeBusy)
            .Parameters.AddWithValue("@Keywords", Keywords)
            .Parameters.AddWithValue("@Language1", Language1)
            .Parameters.AddWithValue("@Location", Location)
            .Parameters.AddWithValue("@ManagersName", ManagersName)
            .Parameters.AddWithValue("@Mileage", Mileage)
            .Parameters.AddWithValue("@Notes", Notes)
            .Parameters.AddWithValue("@OfficeLocation", OfficeLocation)
            .Parameters.AddWithValue("@OrganizationalIDNumber", OrganizationalIDNumber)
            .Parameters.AddWithValue("@OtherAddressPOBox", OtherAddressPOBox)
            .Parameters.AddWithValue("@Priority", Priority)
            .Parameters.AddWithValue("@Private", sPrivate)
            .Parameters.AddWithValue("@Profession", Profession)
            .Parameters.AddWithValue("@ReferredBy", ReferredBy)
            .Parameters.AddWithValue("@Sensitivity", Sensitivity)
            .Parameters.AddWithValue("@Spouse", Spouse)
            .Parameters.AddWithValue("@User1", User1)
            .Parameters.AddWithValue("@User2", User2)
            .Parameters.AddWithValue("@User3", User3)
            .Parameters.AddWithValue("@User4", User4)
            .Parameters.AddWithValue("@WebPage", WebPage)
            .Parameters.AddWithValue("@PagerTypeCode", PagerTypeCode)
            .Parameters.AddWithValue("@DoNotImport", DoNotImport)

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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveImportAddressBook2(ActionType As String, AppCode As String, ImportType As String, ImportAppUserId As String, LineNumber As Long, EmailAddress As String, EmailType As String, EmailDisplayName As String, AddressBookTypeCode As String, DomainName As String, AppGroupID As String, AppUserID As String, Title As String, FirstName As String, MiddleName As String, LastName As String, Suffix As String, Company As String, Department As String, JobTitle As String, BusinessStreet As String, BusinessStreet2 As String, BusinessStreet3 As String, BusinessCity As String, BusinessState As String, BusinessPostalCode As String, BusinessCountryRegion As String, HomeStreet As String, HomeStreet2 As String, HomeStreet3 As String, HomeCity As String, HomeState As String, HomePostalCode As String, HomeCountryRegion As String, OtherStreet As String, OtherStreet2 As String, OtherStreet3 As String, OtherCity As String, OtherState As String, OtherPostalCode As String, OtherCountryRegion As String, AssistantsPhone As String, BusinessFax As String, BusinessPhone As String, BusinessPhone2 As String, Callback As String, CarPhone As String, CompanyMainPhone As String, HomeFax As String, HomePhone As String, HomePhone2 As String, ISDN As String, MobilePhone As String, OtherFax As String, OtherPhone As String, Pager As String, PrimaryPhone As String, RadioPhone As String, TTYTDDPhone As String, Telex As String, Account As String, Anniversary As String, AssistantsName As String, BillingInformation As String, Birthday As String, BusinessAddressPOBox As String, Categories As String, Children As String, DirectoryServer As String, Email2Address As String, Email2Type As String, Email2DisplayName As String, Email3Address As String, Email3Type As String, Email3DisplayName As String, Gender As String, GovernmentIDNumber As String, Hobby As String, HomeAddressPOBox As String, Initials As String, InternetFreeBusy As String, Keywords As String, Language1 As String, Location As String, ManagersName As String, Mileage As String, Notes As String, OfficeLocation As String, OrganizationalIDNumber As String, OtherAddressPOBox As String, Priority As String, sPrivate As String, Profession As String, ReferredBy As String, Sensitivity As String, Spouse As String, User1 As String, User2 As String, User3 As String, User4 As String, WebPage As String, PagerTypeCode As String, CriticalMessagingAddress As String, DoNotImport As Boolean) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveImportAddressBook2"
            .Parameters.AddWithValue("@ActionType", ActionType)
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@ImportType", ImportType)
            .Parameters.AddWithValue("@ImportAppUserId", ImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", LineNumber)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
            .Parameters.AddWithValue("@EmailType", EmailType)
            .Parameters.AddWithValue("@EmailDisplayName", EmailDisplayName)
            .Parameters.AddWithValue("@AddressBookTypeCode", AddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", DomainName)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@Title", Title)
            .Parameters.AddWithValue("@FirstName", FirstName)
            .Parameters.AddWithValue("@MiddleName", MiddleName)
            .Parameters.AddWithValue("@LastName", LastName)
            .Parameters.AddWithValue("@Suffix", Suffix)
            .Parameters.AddWithValue("@Company", Company)
            .Parameters.AddWithValue("@Department", Department)
            .Parameters.AddWithValue("@JobTitle", JobTitle)
            .Parameters.AddWithValue("@BusinessStreet", BusinessStreet)
            .Parameters.AddWithValue("@BusinessStreet2", BusinessStreet2)
            .Parameters.AddWithValue("@BusinessStreet3", BusinessStreet3)
            .Parameters.AddWithValue("@BusinessCity", BusinessCity)
            .Parameters.AddWithValue("@BusinessState", BusinessState)
            .Parameters.AddWithValue("@BusinessPostalCode", BusinessPostalCode)
            .Parameters.AddWithValue("@BusinessCountryRegion", BusinessCountryRegion)
            .Parameters.AddWithValue("@HomeStreet", HomeStreet)
            .Parameters.AddWithValue("@HomeStreet2", HomeStreet2)
            .Parameters.AddWithValue("@HomeStreet3", HomeStreet3)
            .Parameters.AddWithValue("@HomeCity", HomeCity)
            .Parameters.AddWithValue("@HomeState", HomeState)
            .Parameters.AddWithValue("@HomePostalCode", HomePostalCode)
            .Parameters.AddWithValue("@HomeCountryRegion", HomeCountryRegion)
            .Parameters.AddWithValue("@OtherStreet", OtherStreet)
            .Parameters.AddWithValue("@OtherStreet2", OtherStreet2)
            .Parameters.AddWithValue("@OtherStreet3", OtherStreet3)
            .Parameters.AddWithValue("@OtherCity", OtherCity)
            .Parameters.AddWithValue("@OtherState", OtherState)
            .Parameters.AddWithValue("@OtherPostalCode", OtherPostalCode)
            .Parameters.AddWithValue("@OtherCountryRegion", OtherCountryRegion)
            .Parameters.AddWithValue("@AssistantsPhone", AssistantsPhone)
            .Parameters.AddWithValue("@BusinessFax", BusinessFax)
            .Parameters.AddWithValue("@BusinessPhone", BusinessPhone)
            .Parameters.AddWithValue("@BusinessPhone2", BusinessPhone2)
            .Parameters.AddWithValue("@Callback", Callback)
            .Parameters.AddWithValue("@CarPhone", CarPhone)
            .Parameters.AddWithValue("@CompanyMainPhone", CompanyMainPhone)
            .Parameters.AddWithValue("@HomeFax", HomeFax)
            .Parameters.AddWithValue("@HomePhone", HomePhone)
            .Parameters.AddWithValue("@HomePhone2", HomePhone2)
            .Parameters.AddWithValue("@ISDN", ISDN)
            .Parameters.AddWithValue("@MobilePhone", MobilePhone)
            .Parameters.AddWithValue("@OtherFax", OtherFax)
            .Parameters.AddWithValue("@OtherPhone", OtherPhone)
            .Parameters.AddWithValue("@Pager", Pager)
            .Parameters.AddWithValue("@PrimaryPhone", PrimaryPhone)
            .Parameters.AddWithValue("@RadioPhone", RadioPhone)
            .Parameters.AddWithValue("@TTYTDDPhone", TTYTDDPhone)
            .Parameters.AddWithValue("@Telex", Telex)
            .Parameters.AddWithValue("@Account", Account)
            .Parameters.AddWithValue("@Anniversary", Anniversary)
            .Parameters.AddWithValue("@AssistantsName", AssistantsName)
            .Parameters.AddWithValue("@BillingInformation", BillingInformation)
            .Parameters.AddWithValue("@Birthday", Birthday)
            .Parameters.AddWithValue("@BusinessAddressPOBox", BusinessAddressPOBox)
            .Parameters.AddWithValue("@Categories", Categories)
            .Parameters.AddWithValue("@Children", Children)
            .Parameters.AddWithValue("@DirectoryServer", DirectoryServer)
            .Parameters.AddWithValue("@Email2Address", Email2Address)
            .Parameters.AddWithValue("@Email2Type", Email2Type)
            .Parameters.AddWithValue("@Email2DisplayName", Email2DisplayName)
            .Parameters.AddWithValue("@Email3Address", Email3Address)
            .Parameters.AddWithValue("@Email3Type", Email3Type)
            .Parameters.AddWithValue("@Email3DisplayName", Email3DisplayName)
            .Parameters.AddWithValue("@Gender", Gender)
            .Parameters.AddWithValue("@GovernmentIDNumber", GovernmentIDNumber)
            .Parameters.AddWithValue("@Hobby", Hobby)
            .Parameters.AddWithValue("@HomeAddressPOBox", HomeAddressPOBox)
            .Parameters.AddWithValue("@Initials", Initials)
            .Parameters.AddWithValue("@InternetFreeBusy", InternetFreeBusy)
            .Parameters.AddWithValue("@Keywords", Keywords)
            .Parameters.AddWithValue("@Language1", Language1)
            .Parameters.AddWithValue("@Location", Location)
            .Parameters.AddWithValue("@ManagersName", ManagersName)
            .Parameters.AddWithValue("@Mileage", Mileage)
            .Parameters.AddWithValue("@Notes", Notes)
            .Parameters.AddWithValue("@OfficeLocation", OfficeLocation)
            .Parameters.AddWithValue("@OrganizationalIDNumber", OrganizationalIDNumber)
            .Parameters.AddWithValue("@OtherAddressPOBox", OtherAddressPOBox)
            .Parameters.AddWithValue("@Priority", Priority)
            .Parameters.AddWithValue("@Private", sPrivate)
            .Parameters.AddWithValue("@Profession", Profession)
            .Parameters.AddWithValue("@ReferredBy", ReferredBy)
            .Parameters.AddWithValue("@Sensitivity", Sensitivity)
            .Parameters.AddWithValue("@Spouse", Spouse)
            .Parameters.AddWithValue("@User1", User1)
            .Parameters.AddWithValue("@User2", User2)
            .Parameters.AddWithValue("@User3", User3)
            .Parameters.AddWithValue("@User4", User4)
            .Parameters.AddWithValue("@WebPage", WebPage)
            .Parameters.AddWithValue("@PagerTypeCode", PagerTypeCode)
            .Parameters.AddWithValue("@CriticalMessagingAddress", CriticalMessagingAddress)
            .Parameters.AddWithValue("@DoNotImport", DoNotImport)

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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveImportAddressBookNotesOnly(ActionType As String, AppCode As String, ImportType As String, ImportAppUserId As String, LineNumber As Long, Notes As String) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveImportAddressBookNotesOnly"
            .Parameters.AddWithValue("@ActionType", ActionType)
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@ImportType", ImportType)
            .Parameters.AddWithValue("@ImportAppUserId", ImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", LineNumber)
            .Parameters.AddWithValue("@Notes", Notes)
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

        End Try

        Return retValue

    End Function
    Public Shared Function CheckImportAddressBookQuick(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal sAddressBookTypeCode As String, ByVal sDomainName As String, ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAddressBook_Check"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@ImportType", "Q")
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@ReturnMsg", "")
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

        End Try

        Return retValue

    End Function

    Public Shared Function DeleteImportAddressBook(ByVal sAppCode As String, ByVal sImportAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeleteImportAddressBook"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function DeleteImportAddressBook2(ByVal sAppCode As String, ByVal sImportType As String, ByVal sImportAppUserId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeleteImportAddressBook2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportType", sImportType)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
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

        End Try

        Return retValue

    End Function


    Public Shared Function GetImportAddressBookEntry(ByVal sAppCode As String, ByVal sAppUserId As String, ByVal sLineNumber As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetImportAddressBookEntry"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sAppUserId)
            .Parameters.AddWithValue("@LineNumber", sLineNumber)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetImportAddressBookEntry = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetImportAddressBookEntry2(ByVal sAppCode As String, ByVal sImportType As String, ByVal sAppUserId As String, ByVal sLineNumber As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetImportAddressBookEntry2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportType", sImportType)
            .Parameters.AddWithValue("@ImportAppUserId", sAppUserId)
            .Parameters.AddWithValue("@LineNumber", sLineNumber)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetImportAddressBookEntry2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function UpdateImportAddressBookImportProblems(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal sLineNumber As String, ByVal sImportProblems As String, ByVal bClearImport As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUpdateImportAddressBookImportProblems"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", sLineNumber)
            .Parameters.AddWithValue("@ImportProblems", sImportProblems)
            .Parameters.AddWithValue("@ClearImport", bClearImport)
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

        End Try

        Return retValue

    End Function

    Public Shared Function UpdateImportAddressBookImportProblems2(ByVal sAppCode As String, ByVal sImportType As String, ByVal sImportAppUserId As String, ByVal sLineNumber As String, ByVal sImportProblems As String, ByVal bClearImport As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminUpdateImportAddressBookImportProblems2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportType", sImportType)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", sLineNumber)
            .Parameters.AddWithValue("@ImportProblems", sImportProblems)
            .Parameters.AddWithValue("@ClearImport", bClearImport)
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

        End Try

        Return retValue

    End Function

    Public Shared Function UpdateImportAddressBookQuick(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal sAddressBookTypeCode As String, ByVal sDomainName As String, ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAddressBook_Update2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)

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

        End Try

        Return retValue

    End Function

    Public Shared Function UpdateImportAddressBookQuick2(ByVal sAppCode As String, ByVal sImportType As String, ByVal sImportAppUserId As String, ByVal sAddressBookTypeCode As String, ByVal sDomainName As String, ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAddressBook_Update2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportType", sImportType)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)

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

        End Try

        Return retValue

    End Function


    Public Shared Function SaveAddressBookContactAuto(ByVal sAppCode As String, ByVal sAddressBookTypeCode As String, ByVal sAppUserID As String, ByVal sAppGroupID As String, ByVal sEmailAddress As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookContactAuto"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@EmailAddress", sEmailAddress)

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

        End Try

        Return retValue

    End Function

    Public Shared Function GetGroupDefaultAddressBook(ByVal sAppGroupID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppGroupDefaultAddressBook"
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
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

        End Try

        Return retValue
    End Function

    Public Shared Function BindLoadToEmailsLock(ByVal sAppUserID As String, ByVal sSearchTerm As String, ByVal bLock As Boolean) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetToEmailsLock"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@Lock", bLock)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadToEmailsLock = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAddressBookAppUser(ByVal sAppUserId As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sPagerNumber As String, ByVal sPagerTypeCode As String, ByVal sPrimaryCell As String, ByVal sCellCarrierCode As String, ByVal sPrimaryEmail As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sLoggedInAppUserID As String) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookAppUser"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerTypeCode)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@CellCarrierCode", sCellCarrierCode)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUserID)
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

        End Try

        Return retValue

    End Function

	
    ' Used by the Javascript calendar.
    Public Shared Function GetUserCalendarEvents(ByVal sAppUserID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminCalendarEventsAllUser"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
        End With

        cn.Open()

        da = New SqlDataAdapter(cmdSQL)
        da.Fill(dt)

        GetUserCalendarEvents = dt
        cn.Close()

    End Function



    ' Used by Javascript calendar. It requires the return code plus calendar ID.
    Public Shared Function SaveCalendarEvent2(ByVal AppCode As String, ByVal AppUserId As String, ByVal InsertUpdateDelete As String,
        ByVal CalendarEventNumber As Long, ByVal CalendarEventSource As String, ByVal CalendarEventSourceIdentifier As String,
        ByVal StatusLocationCode As String, ByVal StatusAvailabilityCode As String, ByVal ForwardTo As String, ByVal StatusReason As String,
        ByVal StatusSet As String, ByVal StatusEnds As String, ByVal CustomStatus As String) As DataTable
        Dim dt As New DataTable
        Dim da As SqlDataAdapter

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminPostCalendarEvent"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppUserId", AppUserId)
            .Parameters.AddWithValue("@InsertUpdateDelete", InsertUpdateDelete)
            .Parameters.AddWithValue("@CalendarEventNumber", CalendarEventNumber)
            .Parameters.AddWithValue("@CalendarEventSource", CalendarEventSource)
            .Parameters.AddWithValue("@CalendarEventSourceIdentifier", CalendarEventSourceIdentifier)
            .Parameters.AddWithValue("@StatusLocationCode", StatusLocationCode)
            .Parameters.AddWithValue("@StatusAvailabilityCode", StatusAvailabilityCode)
            .Parameters.AddWithValue("@ForwardTo", ForwardTo)
            .Parameters.AddWithValue("@StatusReason", StatusReason)
            .Parameters.AddWithValue("@StatusSet", StatusSet)
            .Parameters.AddWithValue("@StatusEnds", StatusEnds)
            .Parameters.AddWithValue("@CustomStatus", CustomStatus)
        End With

        cn.Open()
        da = New SqlDataAdapter(cmdSQL)
        da.Fill(dt)

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

        SaveCalendarEvent2 = dt
    End Function

    Public Shared Function SaveCalendarEvent(ByVal AppCode As String, ByVal AppUserId As String, ByVal InsertUpdateDelete As String, ByVal CalendarEventNumber As Long, ByVal CalendarEventSource As String, ByVal CalendarEventSourceIdentifier As String, ByVal StatusLocationCode As String, ByVal StatusAvailabilityCode As String, ByVal ForwardTo As String, ByVal StatusReason As String, ByVal StatusSet As String, ByVal StatusEnds As String, ByVal CustomStatus As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminPostCalendarEvent"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppUserId", AppUserId)
            .Parameters.AddWithValue("@InsertUpdateDelete", InsertUpdateDelete)
            .Parameters.AddWithValue("@CalendarEventNumber", CalendarEventNumber)
            .Parameters.AddWithValue("@CalendarEventSource", CalendarEventSource)
            .Parameters.AddWithValue("@CalendarEventSourceIdentifier", CalendarEventSourceIdentifier)
            .Parameters.AddWithValue("@StatusLocationCode", StatusLocationCode)
            .Parameters.AddWithValue("@StatusAvailabilityCode", StatusAvailabilityCode)
            .Parameters.AddWithValue("@ForwardTo", ForwardTo)
            .Parameters.AddWithValue("@StatusReason", StatusReason)
            .Parameters.AddWithValue("@StatusSet", StatusSet)
            .Parameters.AddWithValue("@StatusEnds", StatusEnds)
            .Parameters.AddWithValue("@CustomStatus", CustomStatus)


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

        End Try

        Return retValue

    End Function

    Public Shared Function GetCalendarEvent(ByVal sAppUserID As String, ByVal sCalendarEventNumber As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetCalendarEvent"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@CalendarEventNumber", sCalendarEventNumber)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetCalendarEvent = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function CheckImportAddressBook(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal sImportType As String, ByVal sAddressBookTypeCode As String, ByVal sDomainName As String, ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAddressBook_Check"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@ImportType", sImportType)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@ReturnMsg", "")
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

        End Try

        Return retValue

    End Function

    Public Shared Function UpdateImportAddressBookByType(ByVal sAppCode As String, ByVal sImportType As String, ByVal sImportAppUserId As String, ByVal sAddressBookTypeCode As String, ByVal sDomainName As String, ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAddressBook_Update2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportType", sImportType)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)

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

        End Try

        Return retValue

    End Function

    Public Shared Function GetCSVUsersInGroup(ByVal sAppGroupID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetCSVUsersInGroup"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetConsultGroup(ByVal sConsultGroupID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetConsultGroup"
            .Parameters.AddWithValue("@ConsultGroupId", sConsultGroupID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetConsultGroup = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAvailableUsers2(Optional ByVal sAppUserId As String = "", Optional ByVal bAdminOnly As Boolean = False, Optional ByVal sAppGroupID As String = "", Optional ByVal sSearchTerm As String = "", Optional ByVal bViewAll As Boolean = False) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAvailableUsers2"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AdminOnly", bAdminOnly)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@ViewAll", bViewAll)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAvailableUsers2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadGroupMembers2(ByVal sAppUserId As String, ByVal sAppGroupID As String, Optional ByVal sSearchTerm As String = "", Optional ByVal bViewAll As Boolean = False) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGroupMembers2"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@ViewAll", bViewAll)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadGroupMembers2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAppUserIDBySSOUUID(ByVal sSSOUUID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUserIDBySSOUUID"
            .Parameters.AddWithValue("@SSOUUID", sSSOUUID)
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

        End Try

        Return retValue
    End Function

    Public Shared Function LogUserSessionLogin(ByVal LoggedInAppUserID As String, ByVal AppUserID As String, ByVal AppGroupID As String, ByVal DeviceID As String, ByVal IPAddress As String, ByVal BrowserType As String, ByVal BrowserUserAgent As String, ByVal SSO As Boolean) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLogUserSessionLogin"
            .Parameters.AddWithValue("@LoggedInAppUserID", LoggedInAppUserID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@DeviceID", DeviceID)
            .Parameters.AddWithValue("@IPAddress", IPAddress)
            .Parameters.AddWithValue("@BrowserType", BrowserType)
            .Parameters.AddWithValue("@BrowserUserAgent", BrowserUserAgent)
            .Parameters.AddWithValue("@SSO", SSO)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        LogUserSessionLogin = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function LogUserSessionLogout(ByVal LogID As String, ByVal InitiatedLogout As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLogUserSessionLogout"
            .Parameters.AddWithValue("@LogID", LogID)
            .Parameters.AddWithValue("@InitiatedLogout", InitiatedLogout)

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

        End Try

        Return retValue
    End Function

    Public Shared Function GetGridMessagesCount(AppGroupID As String, AppUserId As String, RR As String, BeginDate As String, EndDate As String, SearchTerm As String, LoggedInUser As String, ShowLastFew As Boolean) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        Dim sortColumns As String = ""
        Dim BeginRecordNumber As Long = 0
        Dim EndRecordNumber As Long = 0
        Dim GetCountOnly As Boolean = True
        Dim ShowPagedResults As Boolean = False

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserMessageHistory5"
            .Parameters.AddWithValue("@AppGroupId", AppGroupID)
            .Parameters.AddWithValue("@AppUserId", AppUserId)
            .Parameters.AddWithValue("@RR", RR)
            .Parameters.AddWithValue("@BeginDate", BeginDate)
            .Parameters.AddWithValue("@EndDate", EndDate)
            .Parameters.AddWithValue("@SearchTerm", SearchTerm)
            .Parameters.AddWithValue("@LoggedInUser", LoggedInUser)
            .Parameters.AddWithValue("@ShowLastFew", ShowLastFew)
            .Parameters.AddWithValue("@BeginRecordNumber", BeginRecordNumber)
            .Parameters.AddWithValue("@EndRecordNumber", EndRecordNumber)
            .Parameters.AddWithValue("@sortColumns", sortColumns)
            .Parameters.AddWithValue("@ShowPagedResults", ShowPagedResults)
            .Parameters.AddWithValue("@GetCountOnly", GetCountOnly)

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

        End Try

        Return retValue
    End Function


    Public Shared Function SendCreatedMessageIP(ByVal sAppCode As String, ByVal sAppUserId As String, ByVal sDeviceId As String, ByVal sMessageTo As String, ByVal sMessageSubject As String, ByVal sMessageBody As String, Optional ByVal lReplyToMessageNumber As Long = 0) As String
        Dim ctx = HttpContext.Current
        Dim retValue As String = ""
        Dim sIPAddress As String = Trim(ctx.Request.ServerVariables("Remote_addr"))
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim rdr As SqlDataReader
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "SendCreatedMessageIP"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@DeviceId", sDeviceId)
            .Parameters.AddWithValue("@MessageTo", sMessageTo)
            .Parameters.AddWithValue("@MessageSubject", sMessageSubject)
            .Parameters.AddWithValue("@MessageBody", sMessageBody)
            .Parameters.AddWithValue("@ReplyToMessageNumber", lReplyToMessageNumber)
            .Parameters.AddWithValue("@IPAddress", sIPAddress)
        End With

        'HttpContext.Current.Response.Write("EXEC SendCreatedMessageIP '" & sAppCode & "', '" & sAppUserId & "', '" & sDeviceId & "', '" & sMessageTo & "', '" & sMessageSubject & "', '" & sMessageBody & "', '" & sIPAddress & "'")
        'HttpContext.Current.Response.End()
        cn.Open()
        rdr = cmdSQL.ExecuteReader()
        While rdr.Read
            retValue = rdr("ReturnMsg")
        End While

        Try

            rdr.Close()
            rdr = Nothing
        Catch ex As Exception

        End Try
        
        If Not cmdSQL Is Nothing Then
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End If

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

        Return retValue

    End Function


    Public Shared Function GetUserNumOpenSessions(ByVal AppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetUserNumOpenSessions"
            .Parameters.AddWithValue("@AppUserID", AppUserID)

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

        End Try

        Return retValue
    End Function

    Public Shared Function SaveAddressBookContact2(ByVal sActionType As String, sAppCode As String, sAddressBookTypeCode As String, sAppUserID As String, sAppGroupId As String, sDomainName As String, sAddressBookID As String, sTitle As String, sFirstName As String, sMiddleName As String, sLastName As String, sSuffix As String, sEmailAddress As String, sEmailDisplayName As String, ByVal sCompany As String, ByVal sDepartment As String, ByVal sJobTitle As String, ByVal sBusinessStreet As String, ByVal sBusinessStreet2 As String, ByVal sBusinessStreet3 As String, ByVal sBusinessCity As String, ByVal sBusinessState As String, ByVal sBusinessPostalCode As String, ByVal sBusinessCountryRegion As String, ByVal sHomeStreet As String, ByVal sHomeStreet2 As String, ByVal sHomeStreet3 As String, ByVal sHomeCity As String, ByVal sHomeState As String, ByVal sHomePostalCode As String, ByVal sHomeCountryRegion As String, ByVal sOtherStreet As String, ByVal sOtherStreet2 As String, ByVal sOtherStreet3 As String, ByVal sOtherCity As String, ByVal sOtherState As String, ByVal sOtherPostalCode As String, ByVal sOtherCountryRegion As String, ByVal sAssistantsPhone As String, ByVal sBusinessFax As String, ByVal sBusinessPhone As String, ByVal sBusinessPhone2 As String, ByVal sCallback As String, ByVal sCarPhone As String, ByVal sCompanyMainPhone As String, ByVal sHomeFax As String, ByVal sHomePhone As String, ByVal sHomePhone2 As String, ByVal sISDN As String, ByVal sMobilePhone As String, ByVal sOtherFax As String, ByVal sOtherPhone As String, ByVal sPager As String, ByVal sPrimaryPhone As String, ByVal sRadioPhone As String, ByVal sTTYTDDPhone As String, ByVal sTelex As String, ByVal sAccount As String, ByVal sAnniversary As String, ByVal sAssistantsName As String, ByVal sBillingInformation As String, ByVal sBirthday As String, ByVal sBusinessAddressPOBox As String, ByVal sCategories As String, ByVal sChildren As String, ByVal sDirectoryServer As String, ByVal sEmailType As String, ByVal sEmail2Address As String, ByVal sEmail2Type As String, ByVal sEmail2DisplayName As String, ByVal sEmail3Address As String, ByVal sEmail3Type As String, ByVal sEmail3DisplayName As String, ByVal sGender As String, ByVal sGovernmentIDNumber As String, ByVal sHobby As String, ByVal sHomeAddressPOBox As String, ByVal sInitials As String, ByVal sInternetFreeBusy As String, ByVal sKeywords As String, ByVal sLanguage1 As String, ByVal sLocation As String, ByVal sManagersName As String, ByVal sMileage As String, ByVal sNotes As String, ByVal sOfficeLocation As String, ByVal sOrganizationalIDNumber As String, ByVal sOtherAddressPOBox As String, ByVal sPriority As String, ByVal sPrivate As String, ByVal sProfession As String, ByVal sReferredBy As String, ByVal sSensitivity As String, ByVal sSpouse As String, ByVal sUser1 As String, ByVal sUser2 As String, ByVal sUser3 As String, ByVal sUser4 As String, ByVal sWebPage As String, ByVal sSupervisor As String, ByVal sSupervisorPhone As String, ByVal sSupervisorEmail As String, ByVal sSupervisorAssistant As String, ByVal sSupervisorAssistantPhone As String, ByVal sSupervisorAssistantEmail As String, ByVal sDepartmentEscalationsContact As String, ByVal sDepartmentEscalationsContactNumber As String, ByVal sDepartmentEscalationsEmail As String, ByVal sCurrentEscalationsContact As String, ByVal sCurrentEscalationsContactPhoneNumber As String, ByVal sCurrentEscalationsEmail As String, ByVal sCurrentEscalationDateFrom As String, ByVal sCurrentEscalationDateTo As String, ByVal sSecondaryEscalationsContact As String, ByVal sSecondaryEscalationsContactPhoneNumber As String, ByVal sSecondaryEscalationsEmail As String, ByVal sSecondaryEscalationsDateFrom As String, ByVal sSecondaryEscalationsDateTo As String, ByVal sTemporaryForwardingEmailAddress1 As String, ByVal sTemporaryForwardingAddress1FromDate As String, ByVal sTemporaryForwardingAddress1ToDate As String, ByVal sTemporaryForwardingEmailAddress2 As String, ByVal sTemporaryForwardingAddress2FromDate As String, ByVal sTemporaryForwardingAddress2ToDate As String, ByVal sBestWaytoContactDuringBusinessHours As String, ByVal sBestWaytoContactAfterBusinessHours As String, ByVal sContactInformationNotes As String, ByVal sEscalationInformationNotes As String, ByVal sInCaseofEmergencyContactInformationName As String, ByVal sInCaseofEmergencyContactRelationship As String, ByVal sInCaseofEmergencyContactInformationPhoneNumber As String, ByVal sInCaseOfEmergencyContactInformationEmailAddress As String, ByVal sContactPriority1 As String, ByVal sContactPriority2 As String, ByVal sContactPriority3 As String, ByVal sFileAs As String, ByVal sPagerType As String, ByVal sDispatcherInfo As String, ByVal sCriticalMessagingAddress As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookContact2"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
            .Parameters.AddWithValue("@Title", sTitle)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@MiddleName", sMiddleName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@Suffix", sSuffix)
            .Parameters.AddWithValue("@EmailAddress", sEmailAddress)
            .Parameters.AddWithValue("@EmailDisplayName", sEmailDisplayName)
            .Parameters.AddWithValue("@Company", sCompany)
            .Parameters.AddWithValue("@Department", sDepartment)
            .Parameters.AddWithValue("@JobTitle", sJobTitle)
            .Parameters.AddWithValue("@BusinessStreet", sBusinessStreet)
            .Parameters.AddWithValue("@BusinessStreet2", sBusinessStreet2)
            .Parameters.AddWithValue("@BusinessStreet3", sBusinessStreet3)
            .Parameters.AddWithValue("@BusinessCity", sBusinessCity)
            .Parameters.AddWithValue("@BusinessState", sBusinessState)
            .Parameters.AddWithValue("@BusinessPostalCode", sBusinessPostalCode)
            .Parameters.AddWithValue("@BusinessCountryRegion", sBusinessCountryRegion)
            .Parameters.AddWithValue("@HomeStreet", sHomeStreet)
            .Parameters.AddWithValue("@HomeStreet2", sHomeStreet2)
            .Parameters.AddWithValue("@HomeStreet3", sHomeStreet3)
            .Parameters.AddWithValue("@HomeCity", sHomeCity)
            .Parameters.AddWithValue("@HomeState", sHomeState)
            .Parameters.AddWithValue("@HomePostalCode", sHomePostalCode)
            .Parameters.AddWithValue("@HomeCountryRegion", sHomeCountryRegion)
            .Parameters.AddWithValue("@OtherStreet", sOtherStreet)
            .Parameters.AddWithValue("@OtherStreet2", sOtherStreet2)
            .Parameters.AddWithValue("@OtherStreet3", sOtherStreet3)
            .Parameters.AddWithValue("@OtherCity", sOtherCity)
            .Parameters.AddWithValue("@OtherState", sOtherState)
            .Parameters.AddWithValue("@OtherPostalCode", sOtherPostalCode)
            .Parameters.AddWithValue("@OtherCountryRegion", sOtherCountryRegion)
            .Parameters.AddWithValue("@AssistantsPhone", sAssistantsPhone)
            .Parameters.AddWithValue("@BusinessFax", sBusinessFax)
            .Parameters.AddWithValue("@BusinessPhone", sBusinessPhone)
            .Parameters.AddWithValue("@BusinessPhone2", sBusinessPhone2)
            .Parameters.AddWithValue("@Callback", sCallback)
            .Parameters.AddWithValue("@CarPhone", sCarPhone)
            .Parameters.AddWithValue("@CompanyMainPhone", sCompanyMainPhone)
            .Parameters.AddWithValue("@HomeFax", sHomeFax)
            .Parameters.AddWithValue("@HomePhone", sHomePhone)
            .Parameters.AddWithValue("@HomePhone2", sHomePhone2)
            .Parameters.AddWithValue("@ISDN", sISDN)
            .Parameters.AddWithValue("@MobilePhone", sMobilePhone)
            .Parameters.AddWithValue("@OtherFax", sOtherFax)
            .Parameters.AddWithValue("@OtherPhone", sOtherPhone)
            .Parameters.AddWithValue("@Pager", sPager)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@RadioPhone", sRadioPhone)
            .Parameters.AddWithValue("@TTYTDDPhone", sTTYTDDPhone)
            .Parameters.AddWithValue("@Telex", sTelex)
            .Parameters.AddWithValue("@Account", sAccount)
            .Parameters.AddWithValue("@Anniversary", sAnniversary)
            .Parameters.AddWithValue("@AssistantsName", sAssistantsName)
            .Parameters.AddWithValue("@BillingInformation", sBillingInformation)
            .Parameters.AddWithValue("@Birthday", sBirthday)
            .Parameters.AddWithValue("@BusinessAddressPOBox", sBusinessAddressPOBox)
            .Parameters.AddWithValue("@Categories", sCategories)
            .Parameters.AddWithValue("@Children", sChildren)
            .Parameters.AddWithValue("@DirectoryServer", sDirectoryServer)
            .Parameters.AddWithValue("@EmailType", sEmailType)
            .Parameters.AddWithValue("@Email2Address", sEmail2Address)
            .Parameters.AddWithValue("@Email2Type", sEmail2Type)
            .Parameters.AddWithValue("@Email2DisplayName", sEmail2DisplayName)
            .Parameters.AddWithValue("@Email3Address", sEmail3Address)
            .Parameters.AddWithValue("@Email3Type", sEmail3Type)
            .Parameters.AddWithValue("@Email3DisplayName", sEmail3DisplayName)
            .Parameters.AddWithValue("@Gender", sGender)
            .Parameters.AddWithValue("@GovernmentIDNumber", sGovernmentIDNumber)
            .Parameters.AddWithValue("@Hobby", sHobby)
            .Parameters.AddWithValue("@HomeAddressPOBox", sHomeAddressPOBox)
            .Parameters.AddWithValue("@Initials", sInitials)
            .Parameters.AddWithValue("@InternetFreeBusy", sInternetFreeBusy)
            .Parameters.AddWithValue("@Keywords", sKeywords)
            .Parameters.AddWithValue("@Language1", sLanguage1)
            .Parameters.AddWithValue("@Location", sLocation)
            .Parameters.AddWithValue("@ManagersName", sManagersName)
            .Parameters.AddWithValue("@Mileage", sMileage)
            .Parameters.AddWithValue("@Notes", sNotes)
            .Parameters.AddWithValue("@OfficeLocation", sOfficeLocation)
            .Parameters.AddWithValue("@OrganizationalIDNumber", sOrganizationalIDNumber)
            .Parameters.AddWithValue("@OtherAddressPOBox", sOtherAddressPOBox)
            .Parameters.AddWithValue("@Priority", sPriority)
            .Parameters.AddWithValue("@Private", sPrivate)
            .Parameters.AddWithValue("@Profession", sProfession)
            .Parameters.AddWithValue("@ReferredBy", sReferredBy)
            .Parameters.AddWithValue("@Sensitivity", sSensitivity)
            .Parameters.AddWithValue("@Spouse", sSpouse)
            .Parameters.AddWithValue("@User1", sUser1)
            .Parameters.AddWithValue("@User2", sUser2)
            .Parameters.AddWithValue("@User3", sUser3)
            .Parameters.AddWithValue("@User4", sUser4)
            .Parameters.AddWithValue("@WebPage", sWebPage)
            .Parameters.AddWithValue("@Supervisor", sSupervisor)
            .Parameters.AddWithValue("@SupervisorPhone", sSupervisorPhone)
            .Parameters.AddWithValue("@SupervisorEmail", sSupervisorEmail)
            .Parameters.AddWithValue("@SupervisorAssistant", sSupervisorAssistant)
            .Parameters.AddWithValue("@SupervisorAssistantPhone", sSupervisorAssistantPhone)
            .Parameters.AddWithValue("@SupervisorAssistantEmail", sSupervisorAssistantEmail)
            .Parameters.AddWithValue("@DepartmentEscalationsContact", sDepartmentEscalationsContact)
            .Parameters.AddWithValue("@DepartmentEscalationsContactNumber", sDepartmentEscalationsContactNumber)
            .Parameters.AddWithValue("@DepartmentEscalationsEmail", sDepartmentEscalationsEmail)
            .Parameters.AddWithValue("@CurrentEscalationsContact", sCurrentEscalationsContact)
            .Parameters.AddWithValue("@CurrentEscalationsContactPhoneNumber", sCurrentEscalationsContactPhoneNumber)
            .Parameters.AddWithValue("@CurrentEscalationsEmail", sCurrentEscalationsEmail)
            .Parameters.AddWithValue("@CurrentEscalationDateFrom", sCurrentEscalationDateFrom)
            .Parameters.AddWithValue("@CurrentEscalationDateTo", sCurrentEscalationDateTo)
            .Parameters.AddWithValue("@SecondaryEscalationsContact", sSecondaryEscalationsContact)
            .Parameters.AddWithValue("@SecondaryEscalationsContactPhoneNumber", sSecondaryEscalationsContactPhoneNumber)
            .Parameters.AddWithValue("@SecondaryEscalationsEmail", sSecondaryEscalationsEmail)
            .Parameters.AddWithValue("@SecondaryEscalationsDateFrom", sSecondaryEscalationsDateFrom)
            .Parameters.AddWithValue("@SecondaryEscalationsDateTo", sSecondaryEscalationsDateTo)
            .Parameters.AddWithValue("@TemporaryForwardingEmailAddress1", sTemporaryForwardingEmailAddress1)
            .Parameters.AddWithValue("@TemporaryForwardingAddress1FromDate", sTemporaryForwardingAddress1FromDate)
            .Parameters.AddWithValue("@TemporaryForwardingAddress1ToDate", sTemporaryForwardingAddress1ToDate)
            .Parameters.AddWithValue("@TemporaryForwardingEmailAddress2", sTemporaryForwardingEmailAddress2)
            .Parameters.AddWithValue("@TemporaryForwardingAddress2FromDate", sTemporaryForwardingAddress2FromDate)
            .Parameters.AddWithValue("@TemporaryForwardingAddress2ToDate", sTemporaryForwardingAddress2ToDate)
            .Parameters.AddWithValue("@BestWaytoContactDuringBusinessHours", sBestWaytoContactDuringBusinessHours)
            .Parameters.AddWithValue("@BestWaytoContactAfterBusinessHours", sBestWaytoContactAfterBusinessHours)
            .Parameters.AddWithValue("@ContactInformationNotes", sContactInformationNotes)
            .Parameters.AddWithValue("@EscalationInformationNotes", sEscalationInformationNotes)
            .Parameters.AddWithValue("@InCaseofEmergencyContactInformationName", sInCaseofEmergencyContactInformationName)
            .Parameters.AddWithValue("@InCaseofEmergencyContactRelationship", sInCaseofEmergencyContactRelationship)
            .Parameters.AddWithValue("@InCaseofEmergencyContactInformationPhoneNumber", sInCaseofEmergencyContactInformationPhoneNumber)
            .Parameters.AddWithValue("@InCaseOfEmergencyContactInformationEmailAddress", sInCaseOfEmergencyContactInformationEmailAddress)
            .Parameters.AddWithValue("@ContactPriority1", sContactPriority1)
            .Parameters.AddWithValue("@ContactPriority2", sContactPriority2)
            .Parameters.AddWithValue("@ContactPriority3", sContactPriority3)
            .Parameters.AddWithValue("@FileAs", sFileAs)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerType)
            .Parameters.AddWithValue("@DispatcherInformationNotes", sDispatcherInfo)
            .Parameters.AddWithValue("@CriticalMessagingAddress", sCriticalMessagingAddress)
        End With

        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        SaveAddressBookContact2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function CheckImportAddressBookQuick2(ByVal sAppCode As String, ByVal sImportAppUserId As String, ByVal sAddressBookTypeCode As String, ByVal sDomainName As String, ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAddressBook_Check2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@ImportType", "Q")
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@ReturnMsg", "")
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAppUser2(ByVal sActionType As String, ByVal sAppUserID As String, ByVal sName As String, ByVal sPrimaryEmail As String, ByVal sReceiptNotificationEmail As String, ByVal sReplyNotificationEmail As String, ByVal sPagerNumber As String, ByVal sPrimaryCell As String, ByVal sCellCarrierCode As String, ByVal sTimeZoneCode As String, ByVal bActive As Boolean, ByVal bNoAutoCreateReply As Boolean, ByVal sFirstName As String, ByVal sLastName As String, ByVal sUserType As String, ByVal sBillRateCode As String, ByVal sCoCode As String, ByVal sAcctNo As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sPrimaryGroup As String, ByVal sForwardEmail1 As String, ByVal sForwardEmail2 As String, ByVal sAppCode As String, ByVal sLoggedInAppUser As String, ByVal sPagerTypeCode As String, ByVal bClearMessageByDevice As Boolean, sSecurityGroupName As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUser2"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@ReceiptNotificationEmail", sReceiptNotificationEmail)
            .Parameters.AddWithValue("@ReplyNotificationEmail", sReplyNotificationEmail)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@CellCarrierCode", sCellCarrierCode)
            .Parameters.AddWithValue("@TimeZoneCode", sTimeZoneCode)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@NoAutoCreateReply", bNoAutoCreateReply)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@UserType", sUserType)
            .Parameters.AddWithValue("@BillRateCode", sBillRateCode)
            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@PrimaryGroup", sPrimaryGroup)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@ForwardEmail1", sForwardEmail1)
            .Parameters.AddWithValue("@ForwardEmail2", sForwardEmail2)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUser)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerTypeCode)
            .Parameters.AddWithValue("@ClearMessageByDevice", bClearMessageByDevice)
            .Parameters.AddWithValue("@SecurityGroupName", sSecurityGroupName)
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

        End Try

        Return retValue

    End Function

    Public Shared Function LogSendPassword(ByVal LoggedInAppUserID As String, ByVal AppUserID As String, ByVal EmailAddress As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLogUserSendPassword"
            .Parameters.AddWithValue("@LoggedInAppUserID", LoggedInAppUserID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)

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

        End Try

        Return retValue
    End Function

    Public Shared Function LogViewAnotherUser(ByVal LogID As String, ByVal LoggedInAppUserID As String, ByVal LoggedInAppGroupID As String, ByVal AppUserID As String, ByVal AppGroupID As String, ByVal ScreenName As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim PageName As String = common.curPageName

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLogUserViewAnotherUser"
            .Parameters.AddWithValue("@LogID", LogID)
            .Parameters.AddWithValue("@LoggedInAppUserID", LoggedInAppUserID)
            .Parameters.AddWithValue("@LoggedInAppGroupID", LoggedInAppGroupID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@AppScreenName", ScreenName)
            .Parameters.AddWithValue("@PageName", PageName)

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

        End Try

        Return retValue
    End Function

    Public Shared Function GetUserMessageChangeCount(ByVal AppUserID As String) As Long
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUserChangeCount"
            .Parameters.AddWithValue("@AppUserID", AppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function LogUserSessionLogin2(ByVal LoggedInAppUserID As String, ByVal AppUserID As String, ByVal AppGroupID As String, ByVal DeviceID As String, ByVal IPAddress As String, ByVal BrowserType As String, ByVal BrowserUserAgent As String, ByVal SSO As Boolean, ByVal ComputerName As String, ByVal AspNetSessionID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLogUserSessionLogin2"
            .Parameters.AddWithValue("@LoggedInAppUserID", LoggedInAppUserID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@DeviceID", DeviceID)
            .Parameters.AddWithValue("@IPAddress", IPAddress)
            .Parameters.AddWithValue("@BrowserType", BrowserType)
            .Parameters.AddWithValue("@BrowserUserAgent", BrowserUserAgent)
            .Parameters.AddWithValue("@SSO", SSO)
            .Parameters.AddWithValue("@ComputerName", ComputerName)
            .Parameters.AddWithValue("@AspNetSessionID", AspNetSessionID)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        LogUserSessionLogin2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function LogUserSessionDetail(ByVal LogID As Long, ByVal sPageName As String, ByVal sQueryString As String) As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLogUserSessionDetail"
            .Parameters.AddWithValue("@LogID", LogID)
            .Parameters.AddWithValue("@PageName", sPageName)
            .Parameters.AddWithValue("@QueryString", sQueryString)
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

        End Try

        Return retValue

    End Function

    Public Shared Function LogUserSessionLogIDMismatch(ByVal LogID As Long, ByVal SessionLogID As Long, ByVal sPageName As String, ByVal sQueryString As String) As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim retValue As String

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminLogUserSessionLogIDMisMatch"
            .Parameters.AddWithValue("@LogID", LogID)
            .Parameters.AddWithValue("@SessionLogID", SessionLogID)
            .Parameters.AddWithValue("@PageName", sPageName)
            .Parameters.AddWithValue("@QueryString", sQueryString)
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

        End Try

        Return retValue

    End Function


    Public Shared Function GetLogUserSessionDetail(ByVal LoggedInAppUserID As String, ByVal LogID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetLogUserSessionDetail"
            .Parameters.AddWithValue("@LoggedInUser", LoggedInAppUserID)
            .Parameters.AddWithValue("@LogID", LogID)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetLogUserSessionDetail = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function BindLoadDomainNames2(ByVal sAppUserId As String, Optional ByVal bActiveOnly As Boolean = False, Optional ByVal sSearchTerm As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDomainNamesSelect2"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@ActiveOnly", bActiveOnly)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadDomainNames2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAddressBook2(ByVal sAppCode As String, ByVal sAddressBookTypeCode As String, ByVal sAppUserId As String, ByVal sAppGroupId As String, ByVal sDomainName As String, ByVal sAddressBookID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBook2"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAddressBook2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAddressBookContact3(ByVal sActionType As String, sAppCode As String, sAddressBookTypeCode As String, sAppUserID As String, sAppGroupId As String, sDomainName As String, sAddressBookID As String, sTitle As String, sFirstName As String, sMiddleName As String, sLastName As String, sSuffix As String, sEmailAddress As String, sEmailDisplayName As String, ByVal sCompany As String, ByVal sDepartment As String, ByVal sJobTitle As String, ByVal sBusinessStreet As String, ByVal sBusinessStreet2 As String, ByVal sBusinessStreet3 As String, ByVal sBusinessCity As String, ByVal sBusinessState As String, ByVal sBusinessPostalCode As String, ByVal sBusinessCountryRegion As String, ByVal sHomeStreet As String, ByVal sHomeStreet2 As String, ByVal sHomeStreet3 As String, ByVal sHomeCity As String, ByVal sHomeState As String, ByVal sHomePostalCode As String, ByVal sHomeCountryRegion As String, ByVal sOtherStreet As String, ByVal sOtherStreet2 As String, ByVal sOtherStreet3 As String, ByVal sOtherCity As String, ByVal sOtherState As String, ByVal sOtherPostalCode As String, ByVal sOtherCountryRegion As String, ByVal sAssistantsPhone As String, ByVal sBusinessFax As String, ByVal sBusinessPhone As String, ByVal sBusinessPhone2 As String, ByVal sCallback As String, ByVal sCarPhone As String, ByVal sCompanyMainPhone As String, ByVal sHomeFax As String, ByVal sHomePhone As String, ByVal sHomePhone2 As String, ByVal sISDN As String, ByVal sMobilePhone As String, ByVal sOtherFax As String, ByVal sOtherPhone As String, ByVal sPager As String, ByVal sPrimaryPhone As String, ByVal sRadioPhone As String, ByVal sTTYTDDPhone As String, ByVal sTelex As String, ByVal sAccount As String, ByVal sAnniversary As String, ByVal sAssistantsName As String, ByVal sBillingInformation As String, ByVal sBirthday As String, ByVal sBusinessAddressPOBox As String, ByVal sCategories As String, ByVal sChildren As String, ByVal sDirectoryServer As String, ByVal sEmailType As String, ByVal sEmail2Address As String, ByVal sEmail2Type As String, ByVal sEmail2DisplayName As String, ByVal sEmail3Address As String, ByVal sEmail3Type As String, ByVal sEmail3DisplayName As String, ByVal sGender As String, ByVal sGovernmentIDNumber As String, ByVal sHobby As String, ByVal sHomeAddressPOBox As String, ByVal sInitials As String, ByVal sInternetFreeBusy As String, ByVal sKeywords As String, ByVal sLanguage1 As String, ByVal sLocation As String, ByVal sManagersName As String, ByVal sMileage As String, ByVal sNotes As String, ByVal sOfficeLocation As String, ByVal sOrganizationalIDNumber As String, ByVal sOtherAddressPOBox As String, ByVal sPriority As String, ByVal sPrivate As String, ByVal sProfession As String, ByVal sReferredBy As String, ByVal sSensitivity As String, ByVal sSpouse As String, ByVal sUser1 As String, ByVal sUser2 As String, ByVal sUser3 As String, ByVal sUser4 As String, ByVal sWebPage As String, ByVal sSupervisor As String, ByVal sSupervisorPhone As String, ByVal sSupervisorEmail As String, ByVal sSupervisorAssistant As String, ByVal sSupervisorAssistantPhone As String, ByVal sSupervisorAssistantEmail As String, ByVal sDepartmentEscalationsContact As String, ByVal sDepartmentEscalationsContactNumber As String, ByVal sDepartmentEscalationsEmail As String, ByVal sCurrentEscalationsContact As String, ByVal sCurrentEscalationsContactPhoneNumber As String, ByVal sCurrentEscalationsEmail As String, ByVal sCurrentEscalationDateFrom As String, ByVal sCurrentEscalationDateTo As String, ByVal sSecondaryEscalationsContact As String, ByVal sSecondaryEscalationsContactPhoneNumber As String, ByVal sSecondaryEscalationsEmail As String, ByVal sSecondaryEscalationsDateFrom As String, ByVal sSecondaryEscalationsDateTo As String, ByVal sTemporaryForwardingEmailAddress1 As String, ByVal sTemporaryForwardingAddress1FromDate As String, ByVal sTemporaryForwardingAddress1ToDate As String, ByVal sTemporaryForwardingEmailAddress2 As String, ByVal sTemporaryForwardingAddress2FromDate As String, ByVal sTemporaryForwardingAddress2ToDate As String, ByVal sBestWaytoContactDuringBusinessHours As String, ByVal sBestWaytoContactAfterBusinessHours As String, ByVal sContactInformationNotes As String, ByVal sEscalationInformationNotes As String, ByVal sInCaseofEmergencyContactInformationName As String, ByVal sInCaseofEmergencyContactRelationship As String, ByVal sInCaseofEmergencyContactInformationPhoneNumber As String, ByVal sInCaseOfEmergencyContactInformationEmailAddress As String, ByVal sContactPriority1 As String, ByVal sContactPriority2 As String, ByVal sContactPriority3 As String, ByVal sFileAs As String, ByVal sPagerType As String, ByVal sDispatcherInfo As String, ByVal sCriticalMessagingAddress As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookContact3"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
            .Parameters.AddWithValue("@Title", sTitle)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@MiddleName", sMiddleName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@Suffix", sSuffix)
            .Parameters.AddWithValue("@EmailAddress", sEmailAddress)
            .Parameters.AddWithValue("@EmailDisplayName", sEmailDisplayName)
            .Parameters.AddWithValue("@Company", sCompany)
            .Parameters.AddWithValue("@Department", sDepartment)
            .Parameters.AddWithValue("@JobTitle", sJobTitle)
            .Parameters.AddWithValue("@BusinessStreet", sBusinessStreet)
            .Parameters.AddWithValue("@BusinessStreet2", sBusinessStreet2)
            .Parameters.AddWithValue("@BusinessStreet3", sBusinessStreet3)
            .Parameters.AddWithValue("@BusinessCity", sBusinessCity)
            .Parameters.AddWithValue("@BusinessState", sBusinessState)
            .Parameters.AddWithValue("@BusinessPostalCode", sBusinessPostalCode)
            .Parameters.AddWithValue("@BusinessCountryRegion", sBusinessCountryRegion)
            .Parameters.AddWithValue("@HomeStreet", sHomeStreet)
            .Parameters.AddWithValue("@HomeStreet2", sHomeStreet2)
            .Parameters.AddWithValue("@HomeStreet3", sHomeStreet3)
            .Parameters.AddWithValue("@HomeCity", sHomeCity)
            .Parameters.AddWithValue("@HomeState", sHomeState)
            .Parameters.AddWithValue("@HomePostalCode", sHomePostalCode)
            .Parameters.AddWithValue("@HomeCountryRegion", sHomeCountryRegion)
            .Parameters.AddWithValue("@OtherStreet", sOtherStreet)
            .Parameters.AddWithValue("@OtherStreet2", sOtherStreet2)
            .Parameters.AddWithValue("@OtherStreet3", sOtherStreet3)
            .Parameters.AddWithValue("@OtherCity", sOtherCity)
            .Parameters.AddWithValue("@OtherState", sOtherState)
            .Parameters.AddWithValue("@OtherPostalCode", sOtherPostalCode)
            .Parameters.AddWithValue("@OtherCountryRegion", sOtherCountryRegion)
            .Parameters.AddWithValue("@AssistantsPhone", sAssistantsPhone)
            .Parameters.AddWithValue("@BusinessFax", sBusinessFax)
            .Parameters.AddWithValue("@BusinessPhone", sBusinessPhone)
            .Parameters.AddWithValue("@BusinessPhone2", sBusinessPhone2)
            .Parameters.AddWithValue("@Callback", sCallback)
            .Parameters.AddWithValue("@CarPhone", sCarPhone)
            .Parameters.AddWithValue("@CompanyMainPhone", sCompanyMainPhone)
            .Parameters.AddWithValue("@HomeFax", sHomeFax)
            .Parameters.AddWithValue("@HomePhone", sHomePhone)
            .Parameters.AddWithValue("@HomePhone2", sHomePhone2)
            .Parameters.AddWithValue("@ISDN", sISDN)
            .Parameters.AddWithValue("@MobilePhone", sMobilePhone)
            .Parameters.AddWithValue("@OtherFax", sOtherFax)
            .Parameters.AddWithValue("@OtherPhone", sOtherPhone)
            .Parameters.AddWithValue("@Pager", sPager)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@RadioPhone", sRadioPhone)
            .Parameters.AddWithValue("@TTYTDDPhone", sTTYTDDPhone)
            .Parameters.AddWithValue("@Telex", sTelex)
            .Parameters.AddWithValue("@Account", sAccount)
            .Parameters.AddWithValue("@Anniversary", sAnniversary)
            .Parameters.AddWithValue("@AssistantsName", sAssistantsName)
            .Parameters.AddWithValue("@BillingInformation", sBillingInformation)
            .Parameters.AddWithValue("@Birthday", sBirthday)
            .Parameters.AddWithValue("@BusinessAddressPOBox", sBusinessAddressPOBox)
            .Parameters.AddWithValue("@Categories", sCategories)
            .Parameters.AddWithValue("@Children", sChildren)
            .Parameters.AddWithValue("@DirectoryServer", sDirectoryServer)
            .Parameters.AddWithValue("@EmailType", sEmailType)
            .Parameters.AddWithValue("@Email2Address", sEmail2Address)
            .Parameters.AddWithValue("@Email2Type", sEmail2Type)
            .Parameters.AddWithValue("@Email2DisplayName", sEmail2DisplayName)
            .Parameters.AddWithValue("@Email3Address", sEmail3Address)
            .Parameters.AddWithValue("@Email3Type", sEmail3Type)
            .Parameters.AddWithValue("@Email3DisplayName", sEmail3DisplayName)
            .Parameters.AddWithValue("@Gender", sGender)
            .Parameters.AddWithValue("@GovernmentIDNumber", sGovernmentIDNumber)
            .Parameters.AddWithValue("@Hobby", sHobby)
            .Parameters.AddWithValue("@HomeAddressPOBox", sHomeAddressPOBox)
            .Parameters.AddWithValue("@Initials", sInitials)
            .Parameters.AddWithValue("@InternetFreeBusy", sInternetFreeBusy)
            .Parameters.AddWithValue("@Keywords", sKeywords)
            .Parameters.AddWithValue("@Language1", sLanguage1)
            .Parameters.AddWithValue("@Location", sLocation)
            .Parameters.AddWithValue("@ManagersName", sManagersName)
            .Parameters.AddWithValue("@Mileage", sMileage)
            .Parameters.AddWithValue("@Notes", sNotes)
            .Parameters.AddWithValue("@OfficeLocation", sOfficeLocation)
            .Parameters.AddWithValue("@OrganizationalIDNumber", sOrganizationalIDNumber)
            .Parameters.AddWithValue("@OtherAddressPOBox", sOtherAddressPOBox)
            .Parameters.AddWithValue("@Priority", sPriority)
            .Parameters.AddWithValue("@Private", sPrivate)
            .Parameters.AddWithValue("@Profession", sProfession)
            .Parameters.AddWithValue("@ReferredBy", sReferredBy)
            .Parameters.AddWithValue("@Sensitivity", sSensitivity)
            .Parameters.AddWithValue("@Spouse", sSpouse)
            .Parameters.AddWithValue("@User1", sUser1)
            .Parameters.AddWithValue("@User2", sUser2)
            .Parameters.AddWithValue("@User3", sUser3)
            .Parameters.AddWithValue("@User4", sUser4)
            .Parameters.AddWithValue("@WebPage", sWebPage)
            .Parameters.AddWithValue("@Supervisor", sSupervisor)
            .Parameters.AddWithValue("@SupervisorPhone", sSupervisorPhone)
            .Parameters.AddWithValue("@SupervisorEmail", sSupervisorEmail)
            .Parameters.AddWithValue("@SupervisorAssistant", sSupervisorAssistant)
            .Parameters.AddWithValue("@SupervisorAssistantPhone", sSupervisorAssistantPhone)
            .Parameters.AddWithValue("@SupervisorAssistantEmail", sSupervisorAssistantEmail)
            .Parameters.AddWithValue("@DepartmentEscalationsContact", sDepartmentEscalationsContact)
            .Parameters.AddWithValue("@DepartmentEscalationsContactNumber", sDepartmentEscalationsContactNumber)
            .Parameters.AddWithValue("@DepartmentEscalationsEmail", sDepartmentEscalationsEmail)
            .Parameters.AddWithValue("@CurrentEscalationsContact", sCurrentEscalationsContact)
            .Parameters.AddWithValue("@CurrentEscalationsContactPhoneNumber", sCurrentEscalationsContactPhoneNumber)
            .Parameters.AddWithValue("@CurrentEscalationsEmail", sCurrentEscalationsEmail)
            .Parameters.AddWithValue("@CurrentEscalationDateFrom", sCurrentEscalationDateFrom)
            .Parameters.AddWithValue("@CurrentEscalationDateTo", sCurrentEscalationDateTo)
            .Parameters.AddWithValue("@SecondaryEscalationsContact", sSecondaryEscalationsContact)
            .Parameters.AddWithValue("@SecondaryEscalationsContactPhoneNumber", sSecondaryEscalationsContactPhoneNumber)
            .Parameters.AddWithValue("@SecondaryEscalationsEmail", sSecondaryEscalationsEmail)
            .Parameters.AddWithValue("@SecondaryEscalationsDateFrom", sSecondaryEscalationsDateFrom)
            .Parameters.AddWithValue("@SecondaryEscalationsDateTo", sSecondaryEscalationsDateTo)
            .Parameters.AddWithValue("@TemporaryForwardingEmailAddress1", sTemporaryForwardingEmailAddress1)
            .Parameters.AddWithValue("@TemporaryForwardingAddress1FromDate", sTemporaryForwardingAddress1FromDate)
            .Parameters.AddWithValue("@TemporaryForwardingAddress1ToDate", sTemporaryForwardingAddress1ToDate)
            .Parameters.AddWithValue("@TemporaryForwardingEmailAddress2", sTemporaryForwardingEmailAddress2)
            .Parameters.AddWithValue("@TemporaryForwardingAddress2FromDate", sTemporaryForwardingAddress2FromDate)
            .Parameters.AddWithValue("@TemporaryForwardingAddress2ToDate", sTemporaryForwardingAddress2ToDate)
            .Parameters.AddWithValue("@BestWaytoContactDuringBusinessHours", sBestWaytoContactDuringBusinessHours)
            .Parameters.AddWithValue("@BestWaytoContactAfterBusinessHours", sBestWaytoContactAfterBusinessHours)
            .Parameters.AddWithValue("@ContactInformationNotes", sContactInformationNotes)
            .Parameters.AddWithValue("@EscalationInformationNotes", sEscalationInformationNotes)
            .Parameters.AddWithValue("@InCaseofEmergencyContactInformationName", sInCaseofEmergencyContactInformationName)
            .Parameters.AddWithValue("@InCaseofEmergencyContactRelationship", sInCaseofEmergencyContactRelationship)
            .Parameters.AddWithValue("@InCaseofEmergencyContactInformationPhoneNumber", sInCaseofEmergencyContactInformationPhoneNumber)
            .Parameters.AddWithValue("@InCaseOfEmergencyContactInformationEmailAddress", sInCaseOfEmergencyContactInformationEmailAddress)
            .Parameters.AddWithValue("@ContactPriority1", sContactPriority1)
            .Parameters.AddWithValue("@ContactPriority2", sContactPriority2)
            .Parameters.AddWithValue("@ContactPriority3", sContactPriority3)
            .Parameters.AddWithValue("@FileAs", sFileAs)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerType)
            .Parameters.AddWithValue("@DispatcherInformationNotes", sDispatcherInfo)
            .Parameters.AddWithValue("@CriticalMessagingAddress", sCriticalMessagingAddress)
        End With

        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        SaveAddressBookContact3 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetMasterAppGroup2(ByVal sAppGroupID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetMasterAppGroup2"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetMasterAppGroup2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveMasterAppGroup4(ByVal sActionType As String, ByVal sAppGroupID As String, ByVal sAppGroupDescription As String, ByVal sName As String, ByVal sPrimaryContactName As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sPrimaryCell As String, ByVal sPrimaryEmail As String, ByVal sEmergencyEmail As String, ByVal sEmergencyAlertLabel1 As String, ByVal sEmergencyAlertLabel2 As String, ByVal sEmergencyAlertLabel3 As String, ByVal sEmergencyAlertNumber1 As String, ByVal sEmergencyAlertNumber2 As String, ByVal sEmergencyAlertNumber3 As String, ByVal lNotDeliveredAfterSeconds As Long, ByVal lNotReadAfterSeconds As Long, ByVal sCannedQuestion1 As String, ByVal sCannedQuestion2 As String, ByVal sCannedQuestion3 As String, ByVal sCannedQuestion4 As String, ByVal sCannedQuestion5 As String, ByVal sCannedQuestion6 As String, ByVal sCannedReply1 As String, ByVal sCannedReply2 As String, ByVal sCannedReply3 As String, ByVal sCannedReply4 As String, ByVal sCannedReply5 As String, ByVal sCannedReply6 As String, ByVal sReceiptNotificationEmail As String, ByVal sReplyNotificationEmail As String, ByVal lInfoURLInterval As Long, ByVal bChangeHeaderImage As Boolean, ByVal bChangeFooterImage As Boolean, ByVal bActive As Boolean, ByVal bAutoCreateReply As Boolean, ByVal sDomainName As String, ByVal sCoCode As String, ByVal sAcctNo As String, ByVal sTimeZoneCode As String, ByVal bInternalGroup As Boolean, ByVal bGroupMessaging As Boolean, ByVal bOnlyAllowMessagesFromWithinGroup As Boolean, ByVal bAllowMessageFromPrimaryEmail As Boolean, ByVal bAllowReplyOnGroupMessage As Boolean, ByVal lMessageRetentionDays As Long, ByVal sGroupMessageFromIndividualOrGroup As String, ByVal sAppUserIdFormatCode As String, ByVal bSendMessageDetail As Boolean, lMessageExpireDays As String, sDefaultAddressBook As String, sMask As String, lNumPinDigits As Long, bPinsEnabled As Boolean) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveMasterAppGroup4"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppGroupDescription", sAppGroupDescription)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@PrimaryContactName", sPrimaryContactName)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@EmergencyEmailAddress", sEmergencyEmail)
            .Parameters.AddWithValue("@EmergencyAlertLabel1", sEmergencyAlertLabel1)
            .Parameters.AddWithValue("@EmergencyAlertLabel2", sEmergencyAlertLabel2)
            .Parameters.AddWithValue("@EmergencyAlertLabel3", sEmergencyAlertLabel3)
            .Parameters.AddWithValue("@EmergencyAlertNumber1", sEmergencyAlertNumber1)
            .Parameters.AddWithValue("@EmergencyAlertNumber2", sEmergencyAlertNumber2)
            .Parameters.AddWithValue("@EmergencyAlertNumber3", sEmergencyAlertNumber3)
            .Parameters.AddWithValue("@NotDeliveredAfterSeconds", lNotDeliveredAfterSeconds)
            .Parameters.AddWithValue("@NotReadAfterSeconds", lNotReadAfterSeconds)
            .Parameters.AddWithValue("@CannedQuestion1", sCannedQuestion1)
            .Parameters.AddWithValue("@CannedQuestion2", sCannedQuestion2)
            .Parameters.AddWithValue("@CannedQuestion3", sCannedQuestion3)
            .Parameters.AddWithValue("@CannedQuestion4", sCannedQuestion4)
            .Parameters.AddWithValue("@CannedQuestion5", sCannedQuestion5)
            .Parameters.AddWithValue("@CannedQuestion6", sCannedQuestion6)
            .Parameters.AddWithValue("@CannedReply1Email", sCannedReply1)
            .Parameters.AddWithValue("@CannedReply2Email", sCannedReply2)
            .Parameters.AddWithValue("@CannedReply3Email", sCannedReply3)
            .Parameters.AddWithValue("@CannedReply4Email", sCannedReply4)
            .Parameters.AddWithValue("@CannedReply5Email", sCannedReply5)
            .Parameters.AddWithValue("@CannedReply6Email", sCannedReply6)
            .Parameters.AddWithValue("@ReceiptNotificationEmail", sReceiptNotificationEmail)
            .Parameters.AddWithValue("@ReplyNotificationEmail", sReplyNotificationEmail)
            .Parameters.AddWithValue("@InfoURLInterval", lInfoURLInterval)
            .Parameters.AddWithValue("@ChangeHeaderImage", bChangeHeaderImage)
            .Parameters.AddWithValue("@ChangeFooterImage", bChangeFooterImage)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@AutoCreateReply", bAutoCreateReply)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@TimeZoneCode", sTimeZoneCode)
            .Parameters.AddWithValue("@InternalGroup", bInternalGroup)
            .Parameters.AddWithValue("@GroupMessaging", bGroupMessaging)
            .Parameters.AddWithValue("@OnlyAllowMessagesFromWithinGroup", bOnlyAllowMessagesFromWithinGroup)
            .Parameters.AddWithValue("@AllowMessageFromPrimaryEmail", bAllowMessageFromPrimaryEmail)
            .Parameters.AddWithValue("@AllowReplyOnGroupMessage", bAllowReplyOnGroupMessage)
            .Parameters.AddWithValue("@MessageRetentionDays", lMessageRetentionDays)
            .Parameters.AddWithValue("@GroupMessageFromIndividualOrGroup", sGroupMessageFromIndividualOrGroup)
            .Parameters.AddWithValue("@AppUserIdFormatCode", sAppUserIdFormatCode)
            .Parameters.AddWithValue("@SendMessageDetail", bSendMessageDetail)
            .Parameters.AddWithValue("@MessageExpireDays", lMessageExpireDays)
            .Parameters.AddWithValue("@DefaultAddressBook", sDefaultAddressBook)
            .Parameters.AddWithValue("@Mask", sMask)
            .Parameters.AddWithValue("@NumberPinDigits", lNumPinDigits)
            .Parameters.AddWithValue("@PinsEnabled", bPinsEnabled)
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

        End Try

        Return retValue

    End Function


    Public Shared Function GetAppUser2(ByVal sAppUserID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUser2"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAppUser2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function SaveAppUser3(ByVal sActionType As String, ByVal sAppUserID As String, ByVal sName As String, ByVal sPrimaryEmail As String, ByVal sReceiptNotificationEmail As String, ByVal sReplyNotificationEmail As String, ByVal sPagerNumber As String, ByVal sPrimaryCell As String, ByVal sCellCarrierCode As String, ByVal sTimeZoneCode As String, ByVal bActive As Boolean, ByVal bNoAutoCreateReply As Boolean, ByVal sFirstName As String, ByVal sLastName As String, ByVal sUserType As String, ByVal sBillRateCode As String, ByVal sCoCode As String, ByVal sAcctNo As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sPrimaryGroup As String, ByVal sForwardEmail1 As String, ByVal sForwardEmail2 As String, ByVal sAppCode As String, ByVal sLoggedInAppUser As String, ByVal sPagerTypeCode As String, ByVal bClearMessageByDevice As Boolean, sSecurityGroupName As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUser3"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@ReceiptNotificationEmail", sReceiptNotificationEmail)
            .Parameters.AddWithValue("@ReplyNotificationEmail", sReplyNotificationEmail)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@CellCarrierCode", sCellCarrierCode)
            .Parameters.AddWithValue("@TimeZoneCode", sTimeZoneCode)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@NoAutoCreateReply", bNoAutoCreateReply)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@UserType", sUserType)
            .Parameters.AddWithValue("@BillRateCode", sBillRateCode)
            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@PrimaryGroup", sPrimaryGroup)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@ForwardEmail1", sForwardEmail1)
            .Parameters.AddWithValue("@ForwardEmail2", sForwardEmail2)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUser)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerTypeCode)
            .Parameters.AddWithValue("@ClearMessageByDevice", bClearMessageByDevice)
            .Parameters.AddWithValue("@SecurityGroupName", sSecurityGroupName)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAddressBookAppUser2(ByVal sAppUserId As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sPagerNumber As String, ByVal sPagerTypeCode As String, ByVal sPrimaryCell As String, ByVal sCellCarrierCode As String, ByVal sPrimaryEmail As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sLoggedInAppUserID As String) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookAppUser2"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerTypeCode)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@CellCarrierCode", sCellCarrierCode)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUserID)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetAppGroupMask(ByVal sAppGroupId As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppGroupMask"
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadToEmailsLock2(ByVal sAppUserID As String, ByVal sSearchTerm As String, ByVal bLock As Boolean) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetToEmailsLock2"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@Lock", bLock)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadToEmailsLock2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetUserQuickMessage2(ByVal AppCode As String, ByVal AppGroupID As String, ByVal AppUserID As String, ByVal QuickMessageID As Long, ByVal LoggedInAppUserID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetQuickMessage2"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupId", AppGroupID)
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@QuickMessageID", QuickMessageID)
            .Parameters.AddWithValue("@LoggedInAppUserId", LoggedInAppUserID)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserQuickMessage2 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetAppGroupNumberPinDigits(ByVal sAppGroupId As String) As Long
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppGroupNumberPinDigits"
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadGroupMembersMasked(ByVal sAppUserId As String, ByVal sAppGroupID As String, ByVal sLoggedInAppUserId As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGroupMembersMasked"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@LoggedInAppUserId", sLoggedInAppUserId)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadGroupMembersMasked = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadGroupMembers2Masked(ByVal sLoggedInAppUserID As String, ByVal sAppUserId As String, ByVal sAppGroupID As String, Optional ByVal sSearchTerm As String = "", Optional ByVal bViewAll As Boolean = False) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGroupMembers2Masked"
            .Parameters.AddWithValue("@LoggedInAppUserId", sLoggedInAppUserID)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@ViewAll", bViewAll)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadGroupMembers2Masked = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function BindLoadAvailableUsers2Masked(Optional ByVal sLoggedInAppUserId As String = "", Optional ByVal sAppUserId As String = "", Optional ByVal bAdminOnly As Boolean = False, Optional ByVal sAppGroupID As String = "", Optional ByVal sSearchTerm As String = "", Optional ByVal bViewAll As Boolean = False) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAvailableUsers2Masked"
            .Parameters.AddWithValue("@LoggedInAppUserId", sLoggedInAppUserId)
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AdminOnly", bAdminOnly)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@ViewAll", bViewAll)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAvailableUsers2Masked = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetPagerPin(ByVal sLoggedInAppUserID As String, ByVal sPagerNumber As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetPagerPin"
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUserID)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
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

        End Try

        Return retValue

    End Function

    Public Shared Function DeleteAddressBookAll(ByVal sAppCode As String, ByVal sLoggedInAppUserID As String, ByVal sAddressBookTypeCode As String, ByVal sDomainName As String, ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminDeleteAddressBookAll"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@LoggedInAppUserId", sLoggedInAppUserID)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
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

        End Try

        Return retValue

    End Function



    Public Shared Function SaveAppUser4(ByVal sActionType As String, ByVal sAppUserID As String, ByVal sName As String, ByVal sPrimaryEmail As String, ByVal sReceiptNotificationEmail As String, ByVal sReplyNotificationEmail As String, ByVal sPagerNumber As String, ByVal sPrimaryCell As String, ByVal sCellCarrierCode As String, ByVal sTimeZoneCode As String, ByVal bActive As Boolean, ByVal bNoAutoCreateReply As Boolean, ByVal sFirstName As String, ByVal sLastName As String, ByVal sUserType As String, ByVal sBillRateCode As String, ByVal sCoCode As String, ByVal sAcctNo As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sPrimaryGroup As String, ByVal sForwardEmail1 As String, ByVal sForwardEmail2 As String, ByVal sAppCode As String, ByVal sLoggedInAppUser As String, ByVal sPagerTypeCode As String, ByVal bClearMessageByDevice As Boolean, sSecurityGroupName As String, sSecurityPin As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppUser4"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@Name", sName)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@ReceiptNotificationEmail", sReceiptNotificationEmail)
            .Parameters.AddWithValue("@ReplyNotificationEmail", sReplyNotificationEmail)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@CellCarrierCode", sCellCarrierCode)
            .Parameters.AddWithValue("@TimeZoneCode", sTimeZoneCode)
            .Parameters.AddWithValue("@Active", bActive)
            .Parameters.AddWithValue("@NoAutoCreateReply", bNoAutoCreateReply)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@UserType", sUserType)
            .Parameters.AddWithValue("@BillRateCode", sBillRateCode)
            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@PrimaryGroup", sPrimaryGroup)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@ForwardEmail1", sForwardEmail1)
            .Parameters.AddWithValue("@ForwardEmail2", sForwardEmail2)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUser)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerTypeCode)
            .Parameters.AddWithValue("@ClearMessageByDevice", bClearMessageByDevice)
            .Parameters.AddWithValue("@SecurityGroupName", sSecurityGroupName)
            .Parameters.AddWithValue("@SecurityPin", sSecurityPin)
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

        End Try

        Return retValue

    End Function

    Public Shared Function GetAppUser3(ByVal sAppUserID As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAppUser3"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAppUser3 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAddressBookTypes3(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal bIsConsultUser As Boolean) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBookType3"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@isConsultUser", bIsConsultUser)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookTypes3 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function SaveCharacterParameter(ByVal sAppUserID As String, ByVal CompanyCode As String, ByVal ModuleCode As String, ByVal ApplicationControlCode As String, ByVal CharacterParameter As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveCharacterParameter"
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@CompanyCode", CompanyCode)
            .Parameters.AddWithValue("@ModuleCode", ModuleCode)
            .Parameters.AddWithValue("@ApplicationControlCode", ApplicationControlCode)
            .Parameters.AddWithValue("@CharacterParameter", CharacterParameter)

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

        End Try

        Return retValue

    End Function


    Public Shared Function SendInviteEmail3(ByVal sAppUserID As String, ByVal sTo As String, ByVal sFrom As String, ByVal sSubject As String, Optional ByVal sCC As String = "", Optional ByVal sBCC As String = "", Optional ByVal sBetaPath As String = "") As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSendInviteEmail3"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@To", sTo)
            .Parameters.AddWithValue("@From", sFrom)
            .Parameters.AddWithValue("@Subject", sSubject)
            .Parameters.AddWithValue("@CC", sCC)
            .Parameters.AddWithValue("@BCC", sBCC)
            .Parameters.AddWithValue("@BetaPath", sBetaPath)
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

        End Try

        Return retValue

    End Function

    Public Shared Function SaveAddressBookAppUser3(ByVal sAppUserId As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sPagerNumber As String, ByVal sPagerTypeCode As String, ByVal sPrimaryCell As String, ByVal sCellCarrierCode As String, ByVal sPrimaryEmail As String, ByVal sPrimaryPhone As String, ByVal sPrimaryExt As String, ByVal sLoggedInAppUserID As String, sCoCode As String, sAcctNo As String, sAppCode As String, sAddressBookTypeCode As String, sAddressBookID As String, sEmailDisplayName As String, sCompany As String, sDepartment As String, sJobTitle As String, sStatementDescription As String, sEmail2Address As String, sHomePhone As String, sEmail3Address As String, sOtherPhone As String, sAssistantsName As String, sAssistantsPhone As String) As String

        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookAppUser3"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@PagerNumber", sPagerNumber)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerTypeCode)
            .Parameters.AddWithValue("@PrimaryCell", sPrimaryCell)
            .Parameters.AddWithValue("@CellCarrierCode", sCellCarrierCode)
            .Parameters.AddWithValue("@PrimaryEmail", sPrimaryEmail)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@PrimaryExt", sPrimaryExt)
            .Parameters.AddWithValue("@LoggedInAppUserID", sLoggedInAppUserID)

            .Parameters.AddWithValue("@CoCode", sCoCode)
            .Parameters.AddWithValue("@AcctNo", sAcctNo)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
            .Parameters.AddWithValue("@EmailDisplayName", sEmailDisplayName)
            .Parameters.AddWithValue("@Company", sCompany)
            .Parameters.AddWithValue("@Department", sDepartment)
            .Parameters.AddWithValue("@JobTitle", sJobTitle)
            .Parameters.AddWithValue("@StatementDescription", sStatementDescription)
            .Parameters.AddWithValue("@Email2Address", sEmail2Address)
            .Parameters.AddWithValue("@HomePhone", sHomePhone)
            .Parameters.AddWithValue("@Email3Address", sEmail3Address)
            .Parameters.AddWithValue("@OtherPhone", sOtherPhone)
            .Parameters.AddWithValue("@AssistantsName", sAssistantsName)
            .Parameters.AddWithValue("@AssistantsPhone", sAssistantsPhone)


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

        End Try

        Return retValue

    End Function


    Public Shared Function GetAddressBookClient(AddressBookTypeCode As String, ByVal sAppGroupID As String, ByVal sAppUserID As String, CoCode As String, AcctNo As String, PagerPhoneNo As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBookClient"
            .Parameters.AddWithValue("@AddressBookTypeCode", AddressBookTypeCode)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupID)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@CoCode", CoCode)
            .Parameters.AddWithValue("@AcctNo", AcctNo)
            .Parameters.AddWithValue("@PagerPhoneNo", PagerPhoneNo)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetAddressBookClient = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAppUsersACWithPager(Optional ByVal sAppUserId As String = "", Optional ByVal bAdminOnly As Boolean = False, Optional ByVal sSearchTerm As String = "", Optional ByVal sAppGroupID As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminAppUsersACWithPager"
            .Parameters.AddWithValue("@AppUserId", sAppUserId)
            .Parameters.AddWithValue("@AdminOnly", bAdminOnly)
            .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAppUsersACWithPager = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function BindLoadAddressBookTypes4(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal bIsConsultUser As Boolean) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBookType4"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@isConsultUser", bIsConsultUser)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookTypes4 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function SaveAddressBookContact4(ByVal sActionType As String, sAppCode As String, sAddressBookTypeCode As String, sAppUserID As String, sAppGroupId As String, sDomainName As String, sAddressBookID As String, sTitle As String, sFirstName As String, sMiddleName As String, sLastName As String, sSuffix As String, sEmailAddress As String, sEmailDisplayName As String, ByVal sCompany As String, ByVal sDepartment As String, ByVal sJobTitle As String, ByVal sBusinessStreet As String, ByVal sBusinessStreet2 As String, ByVal sBusinessStreet3 As String, ByVal sBusinessCity As String, ByVal sBusinessState As String, ByVal sBusinessPostalCode As String, ByVal sBusinessCountryRegion As String, ByVal sHomeStreet As String, ByVal sHomeStreet2 As String, ByVal sHomeStreet3 As String, ByVal sHomeCity As String, ByVal sHomeState As String, ByVal sHomePostalCode As String, ByVal sHomeCountryRegion As String, ByVal sOtherStreet As String, ByVal sOtherStreet2 As String, ByVal sOtherStreet3 As String, ByVal sOtherCity As String, ByVal sOtherState As String, ByVal sOtherPostalCode As String, ByVal sOtherCountryRegion As String, ByVal sAssistantsPhone As String, ByVal sBusinessFax As String, ByVal sBusinessPhone As String, ByVal sBusinessPhone2 As String, ByVal sCallback As String, ByVal sCarPhone As String, ByVal sCompanyMainPhone As String, ByVal sHomeFax As String, ByVal sHomePhone As String, ByVal sHomePhone2 As String, ByVal sISDN As String, ByVal sMobilePhone As String, ByVal sOtherFax As String, ByVal sOtherPhone As String, ByVal sPager As String, ByVal sPrimaryPhone As String, ByVal sRadioPhone As String, ByVal sTTYTDDPhone As String, ByVal sTelex As String, ByVal sAccount As String, ByVal sAnniversary As String, ByVal sAssistantsName As String, ByVal sBillingInformation As String, ByVal sBirthday As String, ByVal sBusinessAddressPOBox As String, ByVal sCategories As String, ByVal sChildren As String, ByVal sDirectoryServer As String, ByVal sEmailType As String, ByVal sEmail2Address As String, ByVal sEmail2Type As String, ByVal sEmail2DisplayName As String, ByVal sEmail3Address As String, ByVal sEmail3Type As String, ByVal sEmail3DisplayName As String, ByVal sGender As String, ByVal sGovernmentIDNumber As String, ByVal sHobby As String, ByVal sHomeAddressPOBox As String, ByVal sInitials As String, ByVal sInternetFreeBusy As String, ByVal sKeywords As String, ByVal sLanguage1 As String, ByVal sLocation As String, ByVal sManagersName As String, ByVal sMileage As String, ByVal sNotes As String, ByVal sOfficeLocation As String, ByVal sOrganizationalIDNumber As String, ByVal sOtherAddressPOBox As String, ByVal sPriority As String, ByVal sPrivate As String, ByVal sProfession As String, ByVal sReferredBy As String, ByVal sSensitivity As String, ByVal sSpouse As String, ByVal sUser1 As String, ByVal sUser2 As String, ByVal sUser3 As String, ByVal sUser4 As String, ByVal sWebPage As String, ByVal sSupervisor As String, ByVal sSupervisorPhone As String, ByVal sSupervisorEmail As String, ByVal sSupervisorAssistant As String, ByVal sSupervisorAssistantPhone As String, ByVal sSupervisorAssistantEmail As String, ByVal sDepartmentEscalationsContact As String, ByVal sDepartmentEscalationsContactNumber As String, ByVal sDepartmentEscalationsEmail As String, ByVal sCurrentEscalationsContact As String, ByVal sCurrentEscalationsContactPhoneNumber As String, ByVal sCurrentEscalationsEmail As String, ByVal sCurrentEscalationDateFrom As String, ByVal sCurrentEscalationDateTo As String, ByVal sSecondaryEscalationsContact As String, ByVal sSecondaryEscalationsContactPhoneNumber As String, ByVal sSecondaryEscalationsEmail As String, ByVal sSecondaryEscalationsDateFrom As String, ByVal sSecondaryEscalationsDateTo As String, ByVal sTemporaryForwardingEmailAddress1 As String, ByVal sTemporaryForwardingAddress1FromDate As String, ByVal sTemporaryForwardingAddress1ToDate As String, ByVal sTemporaryForwardingEmailAddress2 As String, ByVal sTemporaryForwardingAddress2FromDate As String, ByVal sTemporaryForwardingAddress2ToDate As String, ByVal sBestWaytoContactDuringBusinessHours As String, ByVal sBestWaytoContactAfterBusinessHours As String, ByVal sContactInformationNotes As String, ByVal sEscalationInformationNotes As String, ByVal sInCaseofEmergencyContactInformationName As String, ByVal sInCaseofEmergencyContactRelationship As String, ByVal sInCaseofEmergencyContactInformationPhoneNumber As String, ByVal sInCaseOfEmergencyContactInformationEmailAddress As String, ByVal sContactPriority1 As String, ByVal sContactPriority2 As String, ByVal sContactPriority3 As String, ByVal sFileAs As String, ByVal sPagerType As String, ByVal sDispatcherInfo As String, ByVal sCriticalMessagingAddress As String, ByVal sStatementDescription As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAddressBookContact4"
            .Parameters.AddWithValue("@ActionType", sActionType)
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@AppUserId", sAppUserID)
            .Parameters.AddWithValue("@AppGroupId", sAppGroupId)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AddressBookID", sAddressBookID)
            .Parameters.AddWithValue("@Title", sTitle)
            .Parameters.AddWithValue("@FirstName", sFirstName)
            .Parameters.AddWithValue("@MiddleName", sMiddleName)
            .Parameters.AddWithValue("@LastName", sLastName)
            .Parameters.AddWithValue("@Suffix", sSuffix)
            .Parameters.AddWithValue("@EmailAddress", sEmailAddress)
            .Parameters.AddWithValue("@EmailDisplayName", sEmailDisplayName)
            .Parameters.AddWithValue("@Company", sCompany)
            .Parameters.AddWithValue("@Department", sDepartment)
            .Parameters.AddWithValue("@JobTitle", sJobTitle)
            .Parameters.AddWithValue("@BusinessStreet", sBusinessStreet)
            .Parameters.AddWithValue("@BusinessStreet2", sBusinessStreet2)
            .Parameters.AddWithValue("@BusinessStreet3", sBusinessStreet3)
            .Parameters.AddWithValue("@BusinessCity", sBusinessCity)
            .Parameters.AddWithValue("@BusinessState", sBusinessState)
            .Parameters.AddWithValue("@BusinessPostalCode", sBusinessPostalCode)
            .Parameters.AddWithValue("@BusinessCountryRegion", sBusinessCountryRegion)
            .Parameters.AddWithValue("@HomeStreet", sHomeStreet)
            .Parameters.AddWithValue("@HomeStreet2", sHomeStreet2)
            .Parameters.AddWithValue("@HomeStreet3", sHomeStreet3)
            .Parameters.AddWithValue("@HomeCity", sHomeCity)
            .Parameters.AddWithValue("@HomeState", sHomeState)
            .Parameters.AddWithValue("@HomePostalCode", sHomePostalCode)
            .Parameters.AddWithValue("@HomeCountryRegion", sHomeCountryRegion)
            .Parameters.AddWithValue("@OtherStreet", sOtherStreet)
            .Parameters.AddWithValue("@OtherStreet2", sOtherStreet2)
            .Parameters.AddWithValue("@OtherStreet3", sOtherStreet3)
            .Parameters.AddWithValue("@OtherCity", sOtherCity)
            .Parameters.AddWithValue("@OtherState", sOtherState)
            .Parameters.AddWithValue("@OtherPostalCode", sOtherPostalCode)
            .Parameters.AddWithValue("@OtherCountryRegion", sOtherCountryRegion)
            .Parameters.AddWithValue("@AssistantsPhone", sAssistantsPhone)
            .Parameters.AddWithValue("@BusinessFax", sBusinessFax)
            .Parameters.AddWithValue("@BusinessPhone", sBusinessPhone)
            .Parameters.AddWithValue("@BusinessPhone2", sBusinessPhone2)
            .Parameters.AddWithValue("@Callback", sCallback)
            .Parameters.AddWithValue("@CarPhone", sCarPhone)
            .Parameters.AddWithValue("@CompanyMainPhone", sCompanyMainPhone)
            .Parameters.AddWithValue("@HomeFax", sHomeFax)
            .Parameters.AddWithValue("@HomePhone", sHomePhone)
            .Parameters.AddWithValue("@HomePhone2", sHomePhone2)
            .Parameters.AddWithValue("@ISDN", sISDN)
            .Parameters.AddWithValue("@MobilePhone", sMobilePhone)
            .Parameters.AddWithValue("@OtherFax", sOtherFax)
            .Parameters.AddWithValue("@OtherPhone", sOtherPhone)
            .Parameters.AddWithValue("@Pager", sPager)
            .Parameters.AddWithValue("@PrimaryPhone", sPrimaryPhone)
            .Parameters.AddWithValue("@RadioPhone", sRadioPhone)
            .Parameters.AddWithValue("@TTYTDDPhone", sTTYTDDPhone)
            .Parameters.AddWithValue("@Telex", sTelex)
            .Parameters.AddWithValue("@Account", sAccount)
            .Parameters.AddWithValue("@Anniversary", sAnniversary)
            .Parameters.AddWithValue("@AssistantsName", sAssistantsName)
            .Parameters.AddWithValue("@BillingInformation", sBillingInformation)
            .Parameters.AddWithValue("@Birthday", sBirthday)
            .Parameters.AddWithValue("@BusinessAddressPOBox", sBusinessAddressPOBox)
            .Parameters.AddWithValue("@Categories", sCategories)
            .Parameters.AddWithValue("@Children", sChildren)
            .Parameters.AddWithValue("@DirectoryServer", sDirectoryServer)
            .Parameters.AddWithValue("@EmailType", sEmailType)
            .Parameters.AddWithValue("@Email2Address", sEmail2Address)
            .Parameters.AddWithValue("@Email2Type", sEmail2Type)
            .Parameters.AddWithValue("@Email2DisplayName", sEmail2DisplayName)
            .Parameters.AddWithValue("@Email3Address", sEmail3Address)
            .Parameters.AddWithValue("@Email3Type", sEmail3Type)
            .Parameters.AddWithValue("@Email3DisplayName", sEmail3DisplayName)
            .Parameters.AddWithValue("@Gender", sGender)
            .Parameters.AddWithValue("@GovernmentIDNumber", sGovernmentIDNumber)
            .Parameters.AddWithValue("@Hobby", sHobby)
            .Parameters.AddWithValue("@HomeAddressPOBox", sHomeAddressPOBox)
            .Parameters.AddWithValue("@Initials", sInitials)
            .Parameters.AddWithValue("@InternetFreeBusy", sInternetFreeBusy)
            .Parameters.AddWithValue("@Keywords", sKeywords)
            .Parameters.AddWithValue("@Language1", sLanguage1)
            .Parameters.AddWithValue("@Location", sLocation)
            .Parameters.AddWithValue("@ManagersName", sManagersName)
            .Parameters.AddWithValue("@Mileage", sMileage)
            .Parameters.AddWithValue("@Notes", sNotes)
            .Parameters.AddWithValue("@OfficeLocation", sOfficeLocation)
            .Parameters.AddWithValue("@OrganizationalIDNumber", sOrganizationalIDNumber)
            .Parameters.AddWithValue("@OtherAddressPOBox", sOtherAddressPOBox)
            .Parameters.AddWithValue("@Priority", sPriority)
            .Parameters.AddWithValue("@Private", sPrivate)
            .Parameters.AddWithValue("@Profession", sProfession)
            .Parameters.AddWithValue("@ReferredBy", sReferredBy)
            .Parameters.AddWithValue("@Sensitivity", sSensitivity)
            .Parameters.AddWithValue("@Spouse", sSpouse)
            .Parameters.AddWithValue("@User1", sUser1)
            .Parameters.AddWithValue("@User2", sUser2)
            .Parameters.AddWithValue("@User3", sUser3)
            .Parameters.AddWithValue("@User4", sUser4)
            .Parameters.AddWithValue("@WebPage", sWebPage)
            .Parameters.AddWithValue("@Supervisor", sSupervisor)
            .Parameters.AddWithValue("@SupervisorPhone", sSupervisorPhone)
            .Parameters.AddWithValue("@SupervisorEmail", sSupervisorEmail)
            .Parameters.AddWithValue("@SupervisorAssistant", sSupervisorAssistant)
            .Parameters.AddWithValue("@SupervisorAssistantPhone", sSupervisorAssistantPhone)
            .Parameters.AddWithValue("@SupervisorAssistantEmail", sSupervisorAssistantEmail)
            .Parameters.AddWithValue("@DepartmentEscalationsContact", sDepartmentEscalationsContact)
            .Parameters.AddWithValue("@DepartmentEscalationsContactNumber", sDepartmentEscalationsContactNumber)
            .Parameters.AddWithValue("@DepartmentEscalationsEmail", sDepartmentEscalationsEmail)
            .Parameters.AddWithValue("@CurrentEscalationsContact", sCurrentEscalationsContact)
            .Parameters.AddWithValue("@CurrentEscalationsContactPhoneNumber", sCurrentEscalationsContactPhoneNumber)
            .Parameters.AddWithValue("@CurrentEscalationsEmail", sCurrentEscalationsEmail)
            .Parameters.AddWithValue("@CurrentEscalationDateFrom", sCurrentEscalationDateFrom)
            .Parameters.AddWithValue("@CurrentEscalationDateTo", sCurrentEscalationDateTo)
            .Parameters.AddWithValue("@SecondaryEscalationsContact", sSecondaryEscalationsContact)
            .Parameters.AddWithValue("@SecondaryEscalationsContactPhoneNumber", sSecondaryEscalationsContactPhoneNumber)
            .Parameters.AddWithValue("@SecondaryEscalationsEmail", sSecondaryEscalationsEmail)
            .Parameters.AddWithValue("@SecondaryEscalationsDateFrom", sSecondaryEscalationsDateFrom)
            .Parameters.AddWithValue("@SecondaryEscalationsDateTo", sSecondaryEscalationsDateTo)
            .Parameters.AddWithValue("@TemporaryForwardingEmailAddress1", sTemporaryForwardingEmailAddress1)
            .Parameters.AddWithValue("@TemporaryForwardingAddress1FromDate", sTemporaryForwardingAddress1FromDate)
            .Parameters.AddWithValue("@TemporaryForwardingAddress1ToDate", sTemporaryForwardingAddress1ToDate)
            .Parameters.AddWithValue("@TemporaryForwardingEmailAddress2", sTemporaryForwardingEmailAddress2)
            .Parameters.AddWithValue("@TemporaryForwardingAddress2FromDate", sTemporaryForwardingAddress2FromDate)
            .Parameters.AddWithValue("@TemporaryForwardingAddress2ToDate", sTemporaryForwardingAddress2ToDate)
            .Parameters.AddWithValue("@BestWaytoContactDuringBusinessHours", sBestWaytoContactDuringBusinessHours)
            .Parameters.AddWithValue("@BestWaytoContactAfterBusinessHours", sBestWaytoContactAfterBusinessHours)
            .Parameters.AddWithValue("@ContactInformationNotes", sContactInformationNotes)
            .Parameters.AddWithValue("@EscalationInformationNotes", sEscalationInformationNotes)
            .Parameters.AddWithValue("@InCaseofEmergencyContactInformationName", sInCaseofEmergencyContactInformationName)
            .Parameters.AddWithValue("@InCaseofEmergencyContactRelationship", sInCaseofEmergencyContactRelationship)
            .Parameters.AddWithValue("@InCaseofEmergencyContactInformationPhoneNumber", sInCaseofEmergencyContactInformationPhoneNumber)
            .Parameters.AddWithValue("@InCaseOfEmergencyContactInformationEmailAddress", sInCaseOfEmergencyContactInformationEmailAddress)
            .Parameters.AddWithValue("@ContactPriority1", sContactPriority1)
            .Parameters.AddWithValue("@ContactPriority2", sContactPriority2)
            .Parameters.AddWithValue("@ContactPriority3", sContactPriority3)
            .Parameters.AddWithValue("@FileAs", sFileAs)
            .Parameters.AddWithValue("@PagerTypeCode", sPagerType)
            .Parameters.AddWithValue("@DispatcherInformationNotes", sDispatcherInfo)
            .Parameters.AddWithValue("@CriticalMessagingAddress", sCriticalMessagingAddress)
            .Parameters.AddWithValue("@StatementDescription", sStatementDescription)
        End With

        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        SaveAddressBookContact4 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function SaveAppGroupAddressBookType(ByVal sAppGroupID As String, ByVal sAddressBookTypeCode As String, ByVal sSecurityGroupName As String, ByVal bDisplayOnDevice As Boolean, ByVal bDisplayOnMessageManager As Boolean) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveAppGroupAddressBookType"
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@SecurityGroupName", sSecurityGroupName)
            .Parameters.AddWithValue("@DisplayOnDevice", bDisplayOnDevice)
            .Parameters.AddWithValue("@DisplayOnMessageManager", bDisplayOnMessageManager)

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

        End Try

        Return retValue

    End Function

    Public Shared Function ResetAddressBookTopStatus(ByVal sAppUserID As String) As String
        Dim retValue As String = ""
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ResetAddressBookTopStatus"
            .Parameters.AddWithValue("@AppUserID", sAppUserID)

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

        End Try

        Return retValue

    End Function

    Public Shared Function BindLoadAddressBookTypesPI(ByVal sAppCode As String, ByVal sAppUserID As String, Optional ByVal sAddressBookTypeCode As String = "") As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBookTypePI"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookTypesPI = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function


    Public Shared Function SaveImportAddressBookPI(ByVal ActionType As String, ByVal AppCode As String, ByVal ImportType As String, ByVal ImportAppUserId As String, ByVal LineNumber As Long, ByVal EmailAddress As String, ByVal EmailType As String, ByVal EmailDisplayName As String, ByVal AddressBookTypeCode As String, ByVal DomainName As String, ByVal AppGroupID As String, ByVal AppUserID As String, ByVal Title As String, ByVal FirstName As String, ByVal MiddleName As String, ByVal LastName As String, ByVal Suffix As String, ByVal Company As String, ByVal Department As String, ByVal JobTitle As String, ByVal BusinessStreet As String, ByVal BusinessStreet2 As String, ByVal BusinessStreet3 As String, ByVal BusinessCity As String, ByVal BusinessState As String, ByVal BusinessPostalCode As String, ByVal BusinessCountryRegion As String, ByVal HomeStreet As String, ByVal HomeStreet2 As String, ByVal HomeStreet3 As String, ByVal HomeCity As String, ByVal HomeState As String, ByVal HomePostalCode As String, ByVal HomeCountryRegion As String, ByVal OtherStreet As String, ByVal OtherStreet2 As String, ByVal OtherStreet3 As String, ByVal OtherCity As String, ByVal OtherState As String, ByVal OtherPostalCode As String, ByVal OtherCountryRegion As String, ByVal AssistantsPhone As String, ByVal BusinessFax As String, ByVal BusinessPhone As String, ByVal BusinessPhone2 As String, ByVal Callback As String, ByVal CarPhone As String, ByVal CompanyMainPhone As String, ByVal HomeFax As String, ByVal HomePhone As String, ByVal HomePhone2 As String, ByVal ISDN As String, ByVal MobilePhone As String, ByVal OtherFax As String, ByVal OtherPhone As String, ByVal Pager As String, ByVal PrimaryPhone As String, ByVal RadioPhone As String, ByVal TTYTDDPhone As String, ByVal Telex As String, ByVal Account As String, ByVal Anniversary As String, ByVal AssistantsName As String, ByVal BillingInformation As String, ByVal Birthday As String, ByVal BusinessAddressPOBox As String, ByVal Categories As String, ByVal Children As String, ByVal DirectoryServer As String, ByVal Email2Address As String, ByVal Email2Type As String, ByVal Email2DisplayName As String, ByVal Email3Address As String, ByVal Email3Type As String, ByVal Email3DisplayName As String, ByVal Gender As String, ByVal GovernmentIDNumber As String, ByVal Hobby As String, ByVal HomeAddressPOBox As String, ByVal Initials As String, ByVal InternetFreeBusy As String, ByVal Keywords As String, ByVal Language1 As String, ByVal Location As String, ByVal ManagersName As String, ByVal Mileage As String, ByVal Notes As String, ByVal OfficeLocation As String, ByVal OrganizationalIDNumber As String, ByVal OtherAddressPOBox As String, ByVal Priority As String, ByVal sPrivate As String, ByVal Profession As String, ByVal ReferredBy As String, ByVal Sensitivity As String, ByVal Spouse As String, ByVal User1 As String, ByVal User2 As String, ByVal User3 As String, ByVal User4 As String, ByVal WebPage As String, ByVal PagerTypeCode As String, ByVal CriticalMessagingAddress As String, ByVal StatementDescription As String, ByVal DoNotImport As Boolean) As String
        'PI = Pager/IntelliMsg
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminSaveImportAddressBookPI"
            .Parameters.AddWithValue("@ActionType", ActionType)
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@ImportType", ImportType)
            .Parameters.AddWithValue("@ImportAppUserId", ImportAppUserId)
            .Parameters.AddWithValue("@LineNumber", LineNumber)
            .Parameters.AddWithValue("@EmailAddress", EmailAddress)
            .Parameters.AddWithValue("@EmailType", EmailType)
            .Parameters.AddWithValue("@EmailDisplayName", EmailDisplayName)
            .Parameters.AddWithValue("@AddressBookTypeCode", AddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", DomainName)
            .Parameters.AddWithValue("@AppGroupID", AppGroupID)
            .Parameters.AddWithValue("@AppUserID", AppUserID)
            .Parameters.AddWithValue("@Title", Title)
            .Parameters.AddWithValue("@FirstName", FirstName)
            .Parameters.AddWithValue("@MiddleName", MiddleName)
            .Parameters.AddWithValue("@LastName", LastName)
            .Parameters.AddWithValue("@Suffix", Suffix)
            .Parameters.AddWithValue("@Company", Company)
            .Parameters.AddWithValue("@Department", Department)
            .Parameters.AddWithValue("@JobTitle", JobTitle)
            .Parameters.AddWithValue("@BusinessStreet", BusinessStreet)
            .Parameters.AddWithValue("@BusinessStreet2", BusinessStreet2)
            .Parameters.AddWithValue("@BusinessStreet3", BusinessStreet3)
            .Parameters.AddWithValue("@BusinessCity", BusinessCity)
            .Parameters.AddWithValue("@BusinessState", BusinessState)
            .Parameters.AddWithValue("@BusinessPostalCode", BusinessPostalCode)
            .Parameters.AddWithValue("@BusinessCountryRegion", BusinessCountryRegion)
            .Parameters.AddWithValue("@HomeStreet", HomeStreet)
            .Parameters.AddWithValue("@HomeStreet2", HomeStreet2)
            .Parameters.AddWithValue("@HomeStreet3", HomeStreet3)
            .Parameters.AddWithValue("@HomeCity", HomeCity)
            .Parameters.AddWithValue("@HomeState", HomeState)
            .Parameters.AddWithValue("@HomePostalCode", HomePostalCode)
            .Parameters.AddWithValue("@HomeCountryRegion", HomeCountryRegion)
            .Parameters.AddWithValue("@OtherStreet", OtherStreet)
            .Parameters.AddWithValue("@OtherStreet2", OtherStreet2)
            .Parameters.AddWithValue("@OtherStreet3", OtherStreet3)
            .Parameters.AddWithValue("@OtherCity", OtherCity)
            .Parameters.AddWithValue("@OtherState", OtherState)
            .Parameters.AddWithValue("@OtherPostalCode", OtherPostalCode)
            .Parameters.AddWithValue("@OtherCountryRegion", OtherCountryRegion)
            .Parameters.AddWithValue("@AssistantsPhone", AssistantsPhone)
            .Parameters.AddWithValue("@BusinessFax", BusinessFax)
            .Parameters.AddWithValue("@BusinessPhone", BusinessPhone)
            .Parameters.AddWithValue("@BusinessPhone2", BusinessPhone2)
            .Parameters.AddWithValue("@Callback", Callback)
            .Parameters.AddWithValue("@CarPhone", CarPhone)
            .Parameters.AddWithValue("@CompanyMainPhone", CompanyMainPhone)
            .Parameters.AddWithValue("@HomeFax", HomeFax)
            .Parameters.AddWithValue("@HomePhone", HomePhone)
            .Parameters.AddWithValue("@HomePhone2", HomePhone2)
            .Parameters.AddWithValue("@ISDN", ISDN)
            .Parameters.AddWithValue("@MobilePhone", MobilePhone)
            .Parameters.AddWithValue("@OtherFax", OtherFax)
            .Parameters.AddWithValue("@OtherPhone", OtherPhone)
            .Parameters.AddWithValue("@Pager", Pager)
            .Parameters.AddWithValue("@PrimaryPhone", PrimaryPhone)
            .Parameters.AddWithValue("@RadioPhone", RadioPhone)
            .Parameters.AddWithValue("@TTYTDDPhone", TTYTDDPhone)
            .Parameters.AddWithValue("@Telex", Telex)
            .Parameters.AddWithValue("@Account", Account)
            .Parameters.AddWithValue("@Anniversary", Anniversary)
            .Parameters.AddWithValue("@AssistantsName", AssistantsName)
            .Parameters.AddWithValue("@BillingInformation", BillingInformation)
            .Parameters.AddWithValue("@Birthday", Birthday)
            .Parameters.AddWithValue("@BusinessAddressPOBox", BusinessAddressPOBox)
            .Parameters.AddWithValue("@Categories", Categories)
            .Parameters.AddWithValue("@Children", Children)
            .Parameters.AddWithValue("@DirectoryServer", DirectoryServer)
            .Parameters.AddWithValue("@Email2Address", Email2Address)
            .Parameters.AddWithValue("@Email2Type", Email2Type)
            .Parameters.AddWithValue("@Email2DisplayName", Email2DisplayName)
            .Parameters.AddWithValue("@Email3Address", Email3Address)
            .Parameters.AddWithValue("@Email3Type", Email3Type)
            .Parameters.AddWithValue("@Email3DisplayName", Email3DisplayName)
            .Parameters.AddWithValue("@Gender", Gender)
            .Parameters.AddWithValue("@GovernmentIDNumber", GovernmentIDNumber)
            .Parameters.AddWithValue("@Hobby", Hobby)
            .Parameters.AddWithValue("@HomeAddressPOBox", HomeAddressPOBox)
            .Parameters.AddWithValue("@Initials", Initials)
            .Parameters.AddWithValue("@InternetFreeBusy", InternetFreeBusy)
            .Parameters.AddWithValue("@Keywords", Keywords)
            .Parameters.AddWithValue("@Language1", Language1)
            .Parameters.AddWithValue("@Location", Location)
            .Parameters.AddWithValue("@ManagersName", ManagersName)
            .Parameters.AddWithValue("@Mileage", Mileage)
            .Parameters.AddWithValue("@Notes", Notes)
            .Parameters.AddWithValue("@OfficeLocation", OfficeLocation)
            .Parameters.AddWithValue("@OrganizationalIDNumber", OrganizationalIDNumber)
            .Parameters.AddWithValue("@OtherAddressPOBox", OtherAddressPOBox)
            .Parameters.AddWithValue("@Priority", Priority)
            .Parameters.AddWithValue("@Private", sPrivate)
            .Parameters.AddWithValue("@Profession", Profession)
            .Parameters.AddWithValue("@ReferredBy", ReferredBy)
            .Parameters.AddWithValue("@Sensitivity", Sensitivity)
            .Parameters.AddWithValue("@Spouse", Spouse)
            .Parameters.AddWithValue("@User1", User1)
            .Parameters.AddWithValue("@User2", User2)
            .Parameters.AddWithValue("@User3", User3)
            .Parameters.AddWithValue("@User4", User4)
            .Parameters.AddWithValue("@WebPage", WebPage)
            .Parameters.AddWithValue("@PagerTypeCode", PagerTypeCode)
            .Parameters.AddWithValue("@CriticalMessagingAddress", CriticalMessagingAddress)
            .Parameters.AddWithValue("@StatementDescription", StatementDescription)
            .Parameters.AddWithValue("@DoNotImport", DoNotImport)

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

        End Try

        Return retValue

    End Function


    Public Shared Function UpdateImportAddressBookPI(ByVal sAppCode As String, ByVal sImportType As String, ByVal sImportAppUserId As String, ByVal sAddressBookTypeCode As String, ByVal sDomainName As String, ByVal sAppGroupID As String, ByVal sAppUserID As String) As String
        Dim retValue As String
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand

        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "ImportAddressBook_UpdatePI"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportType", sImportType)
            .Parameters.AddWithValue("@ImportAppUserId", sImportAppUserId)
            .Parameters.AddWithValue("@AddressBookTypeCode", sAddressBookTypeCode)
            .Parameters.AddWithValue("@DomainName", sDomainName)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)

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

        End Try

        Return retValue

    End Function

    Public Shared Function GetImportAddressBookEntry3(ByVal sAppCode As String, ByVal sImportType As String, ByVal sAppUserId As String, ByVal sLineNumber As String) As DataTable
        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetImportAddressBookEntry3"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@ImportType", sImportType)
            .Parameters.AddWithValue("@ImportAppUserId", sAppUserId)
            .Parameters.AddWithValue("@LineNumber", sLineNumber)
        End With
        cn.Open()
        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetImportAddressBookEntry3 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function BindLoadAddressBookTypes5(ByVal sAppCode As String, ByVal sAppUserID As String, ByVal sAppGroupID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetAddressBookType5"
            .Parameters.AddWithValue("@AppCode", sAppCode)
            .Parameters.AddWithValue("@AppUserID", sAppUserID)
            .Parameters.AddWithValue("@AppGroupID", sAppGroupID)
        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        BindLoadAddressBookTypes5 = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function



    Public Shared Function GetUserQuickMessageGrid(ByVal AppCode As String, ByVal AppGroupID As String, ByVal AppUserID As String, ByVal QuickMessageID As Long, ByVal LoggedInAppUserID As String) As DataTable

        Dim myConn As openAppCn = New openAppCn
        Dim cn As New SqlConnection(myConn.cnString)
        Dim cmdSQL As New SqlCommand
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter
        With cmdSQL
            .Connection = cn
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = "AdminGetQuickMessageGrid"
            .Parameters.AddWithValue("@AppCode", AppCode)
            .Parameters.AddWithValue("@AppGroupId", AppGroupID)
            .Parameters.AddWithValue("@AppUserId", AppUserID)
            .Parameters.AddWithValue("@QuickMessageID", QuickMessageID)
            .Parameters.AddWithValue("@LoggedInAppUserId", LoggedInAppUserID)

        End With

        cn.Open()

        da.SelectCommand = cmdSQL
        da.Fill(dt)

        GetUserQuickMessageGrid = dt

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

        Try
            cn.Close()
            cn.Dispose()
            cn = Nothing
        Catch ex As Exception

        End Try

    End Function

End Class
