
'Imports Microsoft.Win32.Registry



<Serializable()> <DataContract()>
Public Class ALECS_RDP
    Inherits ALECS_Template_Entry_FAKE
    Implements IALECS_Entry

#Region "VARs"

    'Public VARs for direct Access (Fields):

    Public Structure RDPLoginResult

        Public Successful As Boolean?
        Public [Error] As Exception         '# maybe special or user exception type
        Public Start As Date?
        Public TimeToLogin As TimeSpan?

    End Structure

    'Private VARs (Access only via Properties):

#End Region

#Region "Constructors"

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New()
        MyBase.New(GetType(ALECS_RDP))    'inited in BASE Constructor
        INIT()
    End Sub

    '# INITed by UnitTests
    Public Sub New(ByRef objMain As Main)
        MyBase.New(objMain, GetType(ALECS_RDP))    'inited in BASE Constructor
        INIT()
    End Sub

    Private Sub INIT()
        clearLastError()
        Try



        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        End Try
    End Sub
#End Region

#Region "Functions"

#Region "IALECS_Entry Functions"

    Public Overrides Function getDisplayName() As String Implements IALECS_Entry.getDisplayName
        If String.IsNullOrEmpty(Name) Then Return ""
        Return Name
    End Function

    Public Overrides Sub DebugMe(Optional strPREfix As String = "") Implements IALECS_Entry.DebugMe
        Try
            myMain.Log.EnterMethod(Me)
            myMain.Log.SIS.LogColored(Color.LightGray, "DEBUGME: " & getDisplayName())
            myMain.Log.SIS.LogObject(strPREfix & ChildObjectName & ": " & getDisplayName(), Me)

#Region "Details"
            'Dim strDetails As String = "DETAILS"
            'myMain.Log.EnterMethod(strDetails)
            'If String.IsNullOrEmpty(strPREfix) Then strPREfix = "ASS "

            'myMain.Log.LeaveMethod(strDetails)
#End Region


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
        Finally
            myMain.Log.LeaveMethod(Me)
        End Try
    End Sub

#End Region

    Public Function LogIn(myCredentials As Net.NetworkCredential, strConnectionFile As String) As RDPLoginResult
        LogIn = Nothing
        Dim timTiming As New Timing
        clearLastError()
        Try
            myMain.Log.EnterMethod(Me)
            myMain.Log.SIS.LogObject(Level.Debug, "myCredentials", myCredentials)
            myMain.Log.SIS.LogString(Level.Debug, "strConnectionFile", strConnectionFile)


            'put code here...
            Throw New NotImplementedException("please code here")



            myMain.Log.EndsSuccessfully(Me)


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me, timTiming)
        End Try
    End Function

#End Region

End Class
