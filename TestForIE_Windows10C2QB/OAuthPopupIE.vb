









Imports DevDefined.OAuth.Consumer
Imports DevDefined.OAuth.Framework
Imports System
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Configuration
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Net
Imports System.Web
Friend Class IppRealmOAuthProfile
End Class
Partial Public Class OAuthPopupIE
    Inherits Form

    Dim WithEvents oauthBrowser As SHDocVw.InternetExplorer = CreateObject("InternetExplorer.Application")

#Region " Member Variables "

    Private Const _dummyProtocol As String = "http://"
    Private Const _dummyHost As String = "www.example.com"
    Private _ippRealmOAuthProfile As IppRealmOAuthProfile
    Private _requestToken As IToken
    Private _consumerKey As String = ""
    Private _consumerSecret As String = ""
    Private _oauthVerifier As String = ""
    Private _caughtCallback As Boolean = False

#End Region

    Public Sub New(ippRealmOAuthProfile As IppRealmOAuthProfile)
            InitializeComponent()
            'AddHandler oauthBrowser.NavigateComplete2, AddressOf oauthBrowser_Navigated
            _ippRealmOAuthProfile = ippRealmOAuthProfile
            readConsumerKeysFromConfiguration()
            startOAuthHandshake()
        End Sub

        Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        End Sub

#Region " App Configuration "

        Private Sub readConsumerKeysFromConfiguration()
        _consumerKey = "qyprdTofPrkvPkTvyrhmhGkQVeCkk7" '  ConfigurationSettings.AppSettings("consumerKey")
        _consumerSecret = "B4aPhgMLu91cL5ONHLvH6ulpxa8yxYGkuKfEhWLk"  '  ConfigurationSettings.AppSettings("consumerSecret")
    End Sub

#End Region

#Region " Start OAuth Handshake - Get Request Token and Redirect to IPP for User Authorization "

    Private Sub startOAuthHandshake()
        _requestToken = getOauthRequestTokenFromIpp(_consumerKey, _consumerSecret)
        redirectToIppForUserAuthorization(_requestToken)
    End Sub

#Region " Get Request Token "

    Private Function getOauthRequestTokenFromIpp(consumerKey As String, consumerSecret As String) As IToken
        Dim oauthSession As IOAuthSession = createDevDefinedOAuthSession(consumerKey, consumerSecret)
        Return oauthSession.GetRequestToken()
    End Function

#End Region

#Region " Redirect to Service Provider - IPP "

    Private Sub redirectToIppForUserAuthorization(requestToken As IToken)
        Dim oauthUserAuthorizeUrl = ConfigurationSettings.AppSettings("oauthUserAuthUrl")
        Dim myPostData As String = ""
        Dim vHeaders As String = "" ' "Content-Type: application/x-www-form-urlencoded " & vbCrLf
        Dim myURL As String = oauthUserAuthorizeUrl
        Dim myPostDataBytes() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(myPostData)
        'FindMePostData
        'If MyTesting Then
        '    Stop
        '    myPostData = "oauth_token=" + requestToken.Token + "&oauth_callback=" + UriUtility.UrlEncode(_dummyProtocol & _dummyHost)
        '    'vHeaders = "Content-Type: application/x-www-form-urlencoded " & vbCrLf
        '    'vHeaders = "Content-Type: text " & vbCrLf
        '    oauthBrowser.Navigate(myURL, False, postData:=myPostDataBytes, additionalHeaders:=vHeaders)
        '    'oauthBrowser.Navigate(myURL, NIL, myPostDataBytes, vHeaders)
        '    Dim null As Object = Nothing
        '    'oauthBrowser.Navigate(myURL, null, null, null, null)
        'Else
        '    myURL += "?oauth_token=" + requestToken.Token + "&oauth_callback=" + UriUtility.UrlEncode(_dummyProtocol & _dummyHost)
        '    oauthBrowser.Navigate(myURL)
        'End If
        myURL += "?oauth_token=" + requestToken.Token + "&oauth_callback=" + UriUtility.UrlEncode(_dummyProtocol & _dummyHost)
        fnDebug("Redirect URL=" & myURL)
        oauthBrowser.Navigate2(myURL)

        oauthBrowser.Visible = True
        ' oauthBrowser.ScriptErrorsSuppressed = True  ' [[ 2016-02-20 B1.1.8.2 ]]  ' As per William Intuit 
    End Sub

#End Region

#End Region

#Region " Capture Callback, Exchange Request Token for Access Token and Close Dialog "

    Private Sub oauthBrowser_Navigated(sender As Object, ByRef e As Object) ' , url  As WebBrowserNavigatedEventArgs)
        Dim i As Integer
        ' IE_NavigateComplete2
        Dim pOutReason As String = NIL
        Dim myURL As String = e.ToString
        On Error GoTo ErrX
        fnDebug("oauth Browser Navigated Entered host=" & myURL)   '  & e.Url.Host)    ' [[ 2016-02-10 B1.1.8.2 ]]
        If myURL.Contains(_dummyHost) AndAlso Not _caughtCallback Then
            fnDebug("oauth Browser Navigated Dummy host=" & myURL)
            Dim query As NameValueCollection = System.Web.HttpUtility.ParseQueryString(myURL)   'HttpUtility.ParseQueryString(myURL.Query)
            _oauthVerifier = query("oauth_verifier")
            If _oauthVerifier IsNot Nothing Then
                _ippRealmOAuthProfile.realmId = query("realmId")
                _ippRealmOAuthProfile.dataSource = query("dataSource")
                _caughtCallback = True
                fnDebug("oauth Browser Navigated Dummy Navigate blank - Callback set to true ")
                oauthBrowser.Navigate("about:blank")
            ElseIf myURL.Contains("oauth_verifier") Then
                retB = MsgBox("oauth Browser Navigated has oauth_verifier but query did not work. End? ", vbYesNo) : If retB = vbYes Then CloseEnd()
            Else
                fnDebug("oauth Browser Navigated Dummy host=" & myURL & " Sleep 1 second")
                DoEvents()
                Sleep(1000)
            End If
        ElseIf _caughtCallback Then
            fnDebug("oauth Browser Navigated Call Back")
            Dim accessToken As IToken = exchangeRequestTokenForAccessToken(_consumerKey, _consumerSecret, _requestToken)
            _ippRealmOAuthProfile.accessToken = accessToken.Token
            _ippRealmOAuthProfile.accessSecret = accessToken.TokenSecret
            _ippRealmOAuthProfile.expirationDateTime = DateTime.Now.AddMonths(6)
            Me.DialogResult = DialogResult.OK
            If oauthBrowser IsNot Nothing Then oauthBrowser.Quit()
            fnDebug("oauth Browser Navigated closing")
            Me.Close()
        Else
            fnDebug("oauth Browser Navigated else " & myURL)
            'Dim accessToken As IToken = exchangeRequestTokenForAccessToken(_consumerKey, _consumerSecret, _requestToken)
            '_ippRealmOAuthProfile.accessToken = accessToken.Token
            '_ippRealmOAuthProfile.accessSecret = accessToken.TokenSecret
            '_ippRealmOAuthProfile.expirationDateTime = DateTime.Now.AddMonths(6)
            'Me.DialogResult = DialogResult.OK

            'Me.Close()
        End If
        Return

ErrX:
        pOutReason = "oauthBrowser_Navigated err=" & Err.Number & Sp & Err.Description
        Dim Ex As Exception = Err.GetException
        If Ex IsNot Nothing Then
            pOutReason &= Sp & Ex.Message
            If Ex.InnerException IsNot Nothing Then pOutReason &= Sp & Ex.InnerException.Message
        End If
        If MyTesting Then
            Dim l As Integer
            If pOutReason.Contains("Cross-thread operation not valid: Control") Then
                i = i
                Return
            End If
            Stop
            If l = 0 Then
                Resume
            ElseIf l = 1 Then
                Resume Next
            ElseIf l = 3 Then
                Return

            End If
        End If
        If pOutReason.Contains("Cross-thread operation not valid: Control") Then
            fnDebug(pOutReason)
        Else
            DoErrorsNoReturnCancel(pOutReason, ErrorC & Ts3, pOutReason)
        End If

    End Sub

#Region " Exhange Request Token for Access Token "


    Public Function exchangeRequestTokenForAccessToken(consumerKey As String, consumerSecret As String, requestToken As IToken) As IToken
        Dim pOutReason As String = NIL
        On Error GoTo ErrX
        Dim oauthSession As IOAuthSession = createDevDefinedOAuthSession(consumerKey, consumerSecret)
        Return oauthSession.ExchangeRequestTokenForAccessToken(requestToken, _oauthVerifier)
ErrX:
        pOutReason = "exchangeRequestTokenForAccessToken err=" & Err.Number & Sp & Err.Description
        Dim Ex As Exception = Err.GetException
        If Ex IsNot Nothing Then
            pOutReason &= Sp & Ex.Message
            If Ex.InnerException IsNot Nothing Then pOutReason &= Sp & Ex.InnerException.Message
        End If
        If MyTesting Then
            Dim l As Integer
            Stop
            If l = 0 Then
                Resume
            ElseIf l = 1 Then
                Resume Next

            End If
        End If
        MsgBox(pOutReason)
    End Function

#End Region

#End Region

#Region " DevDefined Helper Functions "

    Private Function createDevDefinedOAuthSession(consumerKey As String, consumerSecret As String) As IOAuthSession
        Dim pOutReason As String = NIL
        On Error GoTo ErrX
        Dim oauthRequestTokenUrl As String = ConfigurationSettings.AppSettings("oauthBaseUrl") + ConfigurationSettings.AppSettings("oauthRequestTokenEndpoint")
        Dim oauthAccessTokenUrl As String = ConfigurationSettings.AppSettings("oauthBaseUrl") + ConfigurationSettings.AppSettings("oauthAccessTokenEndpoint")
        Dim oauthUserAuthorizeUrl As String = ConfigurationSettings.AppSettings("oauthUserAuthUrl")

        Dim consumerContext As New OAuthConsumerContext() With {
             .consumerKey = TheConsumerKey,
             .consumerSecret = TheConsumerSecret,
             .SignatureMethod = SignatureMethod.HmacSha1
        }

        Return New OAuthSession(consumerContext, oauthRequestTokenUrl, oauthUserAuthorizeUrl, oauthAccessTokenUrl)
ErrX:
        pOutReason = "oauthBrowser_Navigated err=" & Err.Number & Sp & Err.Description
        Dim Ex As Exception = Err.GetException
        If Ex IsNot Nothing Then
            pOutReason &= Sp & Ex.Message
            If Ex.InnerException IsNot Nothing Then pOutReason &= Sp & Ex.InnerException.Message
        End If
        If MyTesting Then
            Dim l As Integer
            Stop
            If l = 0 Then
                Resume
            ElseIf l = 1 Then
                Resume Next

            End If
        End If
        MsgBox(pOutReason)
    End Function

#End Region

    Private Sub OAuthPopupIE_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Me.oauthBrowser.Navigated += New System.Windows.Forms.WebBrowserNavigatedEventHandler(AddressOf Me.oauthBrowser_Navigated)
        AddHandler oauthBrowser.NavigateComplete2, AddressOf oauthBrowser_Navigated

        'Me.Load += New System.EventHandler(Me.OAuthPopup_Load)
    End Sub

End Class
'End Name space


