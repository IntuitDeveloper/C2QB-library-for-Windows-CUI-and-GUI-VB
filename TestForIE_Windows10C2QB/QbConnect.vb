
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports DevDefined.OAuth.Consumer
Imports DevDefined.OAuth.Framework
Imports System.Diagnostics
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Web

Namespace Intuit.lib.C2QB
    Public NotInheritable Class QbConnect
        Implements IConnect(Of QbConfig)
        Private _oauthVerifier As String = String.Empty
        Private Const _tempProtocol As String = "http://"
        Private Const _tempHost As String = "www.example.com"
        Private _requestToken As IToken
        Private _qbResponse As QbResponse = Nothing
        Private _qbConfiguration As QbConfig = Nothing
        Private _internetExplorer As SHDocVw.InternetExplorer = Nothing
        Private _isReadytoExit As Boolean = False

        Public Sub New()
        End Sub
        Private Async Sub ProcessAuthAsync()
            If _qbConfiguration IsNot Nothing Then
                Dim task As Task(Of QbResponse) = StartOAuthHandshake(_qbConfiguration)
                _qbResponse = Await task
                If _qbConfiguration.SelectedMode = QbConfig.AppMode.Console Then
                    Console.WriteLine("Access Token: " + _qbResponse.AccessToken)
                End If
            End If
        End Sub



        Private Async Function StartOAuthHandshake(objectType As QbConfig) As Task(Of QbResponse)
            _requestToken = getOauthRequestTokenFromIpp(objectType)
            Return Await redirectToIppForUserAuthorization(_requestToken, objectType)
        End Function
        Private Async Function getOauthRequestTokenFromIpp(objectType As QbConfig) As Task(Of IToken)
            Dim oauthSession As IOAuthSession = createDevDefinedOAuthSession(objectType)
            Return oauthSession.GetRequestToken()
        End Function
        Private Async Function redirectToIppForUserAuthorization(requestToken As IToken, objectType As QbConfig) As Task(Of QbResponse)
            Dim navigateUrl As String = String.Format("{0}?oauth_token={1}&oauth_callback={2}", objectType.OauthUserAuthUrl, requestToken.Token, UriUtility.UrlEncode(_tempProtocol & _tempHost))
            Return Await WinSys(navigateUrl, objectType)
        End Function
        Private Async Function WinSys(url As String, qbConfig As QbConfig) As Task(Of QbResponse)
            Return Await WinApiAsync(url, qbConfig)
        End Function

        Private Async Function WinApiAsync(url__1 As String, qbConfig__2 As QbConfig) As Task(Of QbResponse)
            Try
                _qbResponse = New QbResponse()
                _qbConfiguration = qbConfig__2




                Dim IE = New SHDocVw.InternetExplorer()
                Dim URL__3 As Object = url__1
                IE.ToolBar = 0
                IE.StatusBar = False
                IE.MenuBar = False
                IE.Width = 1022
                IE.Height = 782
                IE.Visible = True
                IE.Navigate2(url__1, Nothing, Nothing, Nothing, Nothing)
                If _qbConfiguration.SelectedMode = QbConfig.AppMode.Console Then
                    Console.WriteLine("Begin Loading...")
                End If
                While IE.ReadyState <> SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE
                    If _qbConfiguration.SelectedMode = QbConfig.AppMode.Console Then
                        Console.WriteLine("Loading the web page...")
                    End If
                    If IE.ReadyState = SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE Then
                        AddHandler IE.NavigateComplete2, AddressOf IE_NavigateComplete2
                        If _qbConfiguration.SelectedMode = QbConfig.AppMode.Console Then
                            Console.WriteLine("Loaded!")
                        End If
                        _internetExplorer = IE
                    End If
                End While
            Catch generatedExceptionName As Exception

                Throw
            End Try
            While True
                If _isReadytoExit Then
                    Return _qbResponse
                End If
            End While


        End Function
        Private Sub IE_NavigateComplete2(pDisp As Object, ByRef URL As Object)
            Dim hostUrl As String = TryCast(URL, String)

            If hostUrl.Contains(_tempHost) Then
                Dim query As NameValueCollection = HttpUtility.ParseQueryString(hostUrl)
                _oauthVerifier = query("oauth_verifier")
                _qbResponse.RealmId = query("realmId")
                _qbResponse.DataSource = query("dataSource")
                Dim accessToken As IToken = exchangeRequestTokenForAccessToken(_qbConfiguration, _requestToken)
                _qbResponse.AccessToken = accessToken.Token
                _qbResponse.AccessSecret = accessToken.TokenSecret
                _qbResponse.ExpirationDateTime = DateTime.Now.AddMonths(6)
                _internetExplorer.Quit()
                Debug.WriteLine(String.Format("Access Token : {0}", _qbResponse.AccessToken))
                Debug.WriteLine(String.Format("Access Secret : {0}", _qbResponse.AccessSecret))
                If _qbConfiguration.SelectedMode = QbConfig.AppMode.Console Then
                    Console.WriteLine(String.Format("Access Token : {0}", _qbResponse.AccessToken))
                    Console.WriteLine(String.Format("Access Secret : {0}", _qbResponse.AccessSecret))
                End If
                _isReadytoExit = True
            End If
            If _qbConfiguration.SelectedMode = QbConfig.AppMode.Console Then
                Console.WriteLine(TryCast(URL, String))
            End If
        End Sub
        Private Function createDevDefinedOAuthSession(objectType As QbConfig) As IOAuthSession

            Dim oauthRequestTokenUrl = objectType.OauthBaseUrl + objectType.OauthRequestTokenEndpoint
            Dim oauthAccessTokenUrl = objectType.OauthBaseUrl + objectType.OauthAccessTokenEndpoint
            Dim oauthUserAuthorizeUrl = objectType.OauthUserAuthUrl
            Dim consumerContext As New OAuthConsumerContext() With {
                .ConsumerKey = objectType.ConsumerKey,
                .ConsumerSecret = objectType.ConsumerSecret,
                .SignatureMethod = SignatureMethod.HmacSha1
            }
            Return New OAuthSession(consumerContext, oauthRequestTokenUrl, oauthUserAuthorizeUrl, oauthAccessTokenUrl)
        End Function
        Public Function exchangeRequestTokenForAccessToken(objectType As QbConfig, requestToken As IToken) As IToken
            Dim oauthSession As IOAuthSession = createDevDefinedOAuthSession(objectType)
            Return oauthSession.ExchangeRequestTokenForAccessToken(requestToken, _oauthVerifier)
        End Function

        Public Function IConnect_Validate(qbConfiguration As QbConfig) As Boolean Implements IConnect(Of QbConfig).Validate
            Throw New NotImplementedException()
        End Function

        Public Sub IConnect_Connect(qbConfiguration As QbConfig, ByRef qbResponse As QbResponse) Implements IConnect(Of QbConfig).Connect
            Try
                _qbConfiguration = qbConfiguration
                Dim task As New Task(AddressOf ProcessAuthAsync)
                task.Start()
                task.Wait()
            Catch generatedExceptionName As Exception

                Throw
            End Try
            qbResponse = _qbResponse
        End Sub


    End Class
End Namespace


