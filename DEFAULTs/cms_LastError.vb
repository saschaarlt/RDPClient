
Imports System.Runtime.Serialization



<Serializable()> <DataContract()>
Public MustInherit Class cms_LastError

#Region "VARs"

    'Public VARs for direct Access (Fields):

    'Private VARs (Access only via Properties):
    Private _LastError As Exception = Nothing

#End Region

#Region "Functions"

    Public Sub clearLastError()
        _LastError = Nothing
    End Sub

    Public Sub setLastError(ex As Exception)
        _LastError = ex
    End Sub

    Public Function getLastError() As Exception
        If _LastError Is Nothing Then
            Return Nothing
        End If
        Return _LastError
    End Function

#End Region

End Class
