Imports System.Collections.Generic
Imports System.Linq
Imports System.Text


Interface IConnect(Of T)
        Function Validate(objectType As T) As Boolean
        Sub Connect(objectType As T, ByRef qbResponse As QbResponse)
    End Interface

