
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace Intuit.lib.C2QB
    Public Class QbConfig
        Public Property ApplicationToken() As String
            Get
                Return m_ApplicationToken
            End Get
            Set
                m_ApplicationToken = Value
            End Set
        End Property
        Private m_ApplicationToken As String
        Public Property ConsumerKey() As String
            Get
                Return m_ConsumerKey
            End Get
            Set
                m_ConsumerKey = Value
            End Set
        End Property
        Private m_ConsumerKey As String
        Public Property ConsumerSecret() As String
            Get
                Return m_ConsumerSecret
            End Get
            Set
                m_ConsumerSecret = Value
            End Set
        End Property
        Private m_ConsumerSecret As String
        Public Property OauthRequestTokenEndpoint() As String
            Get
                Return m_OauthRequestTokenEndpoint
            End Get
            Set
                m_OauthRequestTokenEndpoint = Value
            End Set
        End Property
        Private m_OauthRequestTokenEndpoint As String
        Public Property OauthAccessTokenEndpoint() As String
            Get
                Return m_OauthAccessTokenEndpoint
            End Get
            Set
                m_OauthAccessTokenEndpoint = Value
            End Set
        End Property
        Private m_OauthAccessTokenEndpoint As String
        Public Property OauthBaseUrl() As String
            Get
                Return m_OauthBaseUrl
            End Get
            Set
                m_OauthBaseUrl = Value
            End Set
        End Property
        Private m_OauthBaseUrl As String
        Public Property OauthUserAuthUrl() As String
            Get
                Return m_OauthUserAuthUrl
            End Get
            Set
                m_OauthUserAuthUrl = Value
            End Set
        End Property
        Private m_OauthUserAuthUrl As String
        <Flags>
        Public Enum AppMode
            Desktop
            Console
            WindowsPhone
        End Enum
        Public Property SelectedMode() As AppMode
            Get
                Return m_SelectedMode
            End Get
            Set
                m_SelectedMode = Value
            End Set
        End Property
        Private m_SelectedMode As AppMode
    End Class
End Namespace

