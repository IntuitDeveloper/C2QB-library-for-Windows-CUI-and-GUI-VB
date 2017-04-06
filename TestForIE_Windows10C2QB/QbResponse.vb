
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace Intuit.lib.C2QB
    Public Class QbResponse
        Public Property AccessToken() As String
            Get
                Return m_AccessToken
            End Get
            Set
                m_AccessToken = Value
            End Set
        End Property
        Private m_AccessToken As String
        Public Property AccessSecret() As String
            Get
                Return m_AccessSecret
            End Get
            Set
                m_AccessSecret = Value
            End Set
        End Property
        Private m_AccessSecret As String
        Public Property RealmId() As String
            Get
                Return m_RealmId
            End Get
            Set
                m_RealmId = Value
            End Set
        End Property
        Private m_RealmId As String
        Public Property DataSource() As String
            Get
                Return m_DataSource
            End Get
            Set
                m_DataSource = Value
            End Set
        End Property
        Private m_DataSource As String
        Public Property ExpirationDateTime() As DateTime
            Get
                Return m_ExpirationDateTime
            End Get
            Set
                m_ExpirationDateTime = Value
            End Set
        End Property
        Private m_ExpirationDateTime As DateTime

    End Class
End Namespace


