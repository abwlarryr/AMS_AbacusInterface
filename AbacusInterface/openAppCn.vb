Imports Microsoft.VisualBasic

Public Class openAppCn
    Public cnString As String

    Sub New()
        'Dim sDevTestProd As String = System.Configuration.ConfigurationManager.AppSettings("DevTestProd").ToString.Trim.ToUpper
        'Dim sCnString As String = 
        ''If sDevTestProd = "PROD" Then
        ''    sCnString = System.Configuration.ConfigurationManager.ConnectionStrings("ProdConnString").ConnectionString
        ''ElseIf sDevTestProd = "TEST" Then
        ''    sCnString = System.Configuration.ConfigurationManager.ConnectionStrings("TestConnString").ConnectionString
        ''Else 'Dev
        ''    sCnString = System.Configuration.ConfigurationManager.ConnectionStrings("DevConnString").ConnectionString
        ''End If

        'If sCnString <> [String].Empty Then
        '    cnString = sCnString
        'Else
        '    cnString = ""
        'End If
        cnString = frmMain.gDatabaseConnectionString
    End Sub
End Class
