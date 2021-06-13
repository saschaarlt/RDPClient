
Imports System
Imports System.Data
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Text



Public Class returnValue

#Region "VARS"

    'Private VARs (Access only via Properties):
    Private _Success As Boolean = False
    Private _ExitCode As Integer = -1
    Private _Value As Object = Nothing
    Private _Exception As Exception = Nothing
    Private _Message As New ArrayList

#End Region

#Region "Properties"

    Public Property Success As Boolean
        Get
            Return _Success
        End Get
        Set(value As Boolean)
            _Success = value
            If value AndAlso _ExitCode < 0 Then _ExitCode = 0
            If Not value AndAlso _ExitCode = 0 Then _ExitCode = -1
        End Set
    End Property

    Public Property ExitCode As Integer
        Get
            Return _ExitCode
        End Get
        Set(value As Integer)
            _ExitCode = value
        End Set
    End Property

    Public Property Value As Object
        Get
            Return _Value
        End Get
        Set(value As Object)
            _Value = value
        End Set
    End Property

    Public Property Exception As Exception
        Get
            Return _Exception
        End Get
        Set(value As Exception)
            _Exception = value
        End Set
    End Property

    Public Property Message As ArrayList
        Get
            Return _Message
        End Get
        Set(value As ArrayList)
            _Message = value
        End Set
    End Property

#End Region

#Region "Functions"

    Public Sub CheckSuccess()
        If Not Success Then
            If Exception Is Nothing Then
                Exception = New NotImplementedException("This ERROR occurs, because we use the returnValue class and it was NOT successfull. But the Exception in returnValue is also <NOTHING> so we have a desaster :-(. Please inform development about this to make it better.")
            End If
            Throw Exception
        End If
    End Sub

#End Region

End Class





<Serializable()> <DataContract()>
Public Class Timing

#Region "VARS"

    Private myMain As Main

    <NonSerialized>
    Private _StopWatch As New Stopwatch
    Private _Pause As SortedList(Of Int16, Stopwatch)
    Private _State As TimingMode = TimingMode.Idle

    <Serializable()> <DataContract()>
    Public Enum TimingMode As Byte
        Idle = 0
        Run = 1
        Paused = 2
        Stopped = 3
    End Enum

    <Serializable()> <DataContract()>
    Public Enum TimingFormat As Byte
        Ticks = 0
        Milliseconds = 1
        Runtime = 3
        'ms = 13
    End Enum

#End Region

#Region "Constructors"

    Public Sub New()
        MyBase.New()
        _State = TimingMode.Run
        _StopWatch.Start()
    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property Result As Stopwatch
        Get
            Return _StopWatch
        End Get
    End Property

    Public ReadOnly Property State As TimingMode
        Get
            Return _State
        End Get
    End Property

    Public ReadOnly Property Duration(TimeFormat As TimingFormat) As String
        Get
            ' Get the elapsed time as a TimeSpan value. 
            Dim ts As TimeSpan = _StopWatch.Elapsed
            Select Case TimeFormat
                Case TimingFormat.Ticks
                    Return _StopWatch.ElapsedTicks

                Case TimingFormat.Milliseconds
                    ' Format and display the TimeSpan value. 
                    Return String.Format("{0:00}{1:00}{2:00}{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds)

                Case Else
                    ' Format and display the TimeSpan value. 
                    Return String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)

            End Select
        End Get
    End Property

    Public ReadOnly Property Runtime() As String
        Get
            Dim ts As TimeSpan = _StopWatch.Elapsed
            Return String.Format("{0:00}{1:00}{2:00}{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds)
        End Get
    End Property

    Public ReadOnly Property RuntimeSeconds() As Long
        Get
            Dim ts As TimeSpan = _StopWatch.Elapsed
            Return ts.TotalSeconds
        End Get
    End Property

    Public ReadOnly Property RuntimeMilliSeconds() As Double
        Get
            Dim ts As TimeSpan = _StopWatch.Elapsed
            Return ts.TotalMilliseconds
        End Get
    End Property

#End Region

#Region "Functions"

    Public Sub Start()
        _StopWatch.Restart()
        '_StopWatch.Start()
    End Sub

    Public Sub [Stop]()
        _StopWatch.Stop()
    End Sub

    <Obsolete("This function is depricated. Please use STOP function instead!", True)>
    Public Sub Halt()
        _StopWatch.Stop()
    End Sub

    Public Sub Paused()
        _StopWatch.Stop()
    End Sub

    Public Sub [Resume]()
        _StopWatch.Start()
    End Sub

#End Region

End Class





#Region "Cache HELPER Class"

<Serializable()> <DataContract()>
Public Class NET_Cache
    Inherits cms_LastError

#Region "VARs"

    'Dim myMain As Main

    'Public VARs for direct Access (Fields):
    <DataMember(Order:=1)>
    Public Property UID As Guid
    <DataMember(EmitDefaultValue:=False, Order:=2)>
    Public Property ID As String

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public TimeStamp As Date = Nothing

    <DataMember(Order:=99)>
    Public isLoading As Boolean = False

    'Public bolValue As Boolean = False
    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Value As Object = Nothing
    Public Exception As Exception = Nothing
    Public Filled As Boolean = False

    'Private VARs (Access only via Properties):

#End Region

#Region "Constructors"

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "Functions"

#Region "cms_EntryInterfaces Functions"

    Public Function getDisplayName() As String
        Return "Cache filled=" & Filled.ToString & " with value=""" & Value & """"
    End Function

    Public Sub DebugMe(ByRef myMain As Main, Optional strPREfix As String = "", Optional bolForce As Boolean = False)
        Try
            myMain.Log.SIS.LogColored(Color.LightGray, "DEBUGME: " & getDisplayName())
            myMain.Log.SIS.LogObject(strPREfix & "ALECS_Cache: " & getDisplayName(), Me)

#Region "Details"
            Dim strDetails As String = "DETAILS"
            myMain.Log.SIS.EnterMethod(strDetails)
            If String.IsNullOrEmpty(strPREfix) Then strPREfix = "AHE "
            'EX:
            If Exception IsNot Nothing Then
                myMain.Log.SIS.LogObject("..._Exception", Exception)
                'Exception.DebugMe(strPREfix)
            Else
                'myMain.Log.add("Exception is Nothing", Me)
            End If
            myMain.Log.SIS.LeaveMethod(strDetails)
#End Region


        Catch ex As Exception

        End Try
    End Sub

#End Region

#End Region

End Class

#End Region


Public Module myStandard

    Public Function returnPathWithBackslash(strPath As String) As String
        returnPathWithBackslash = strPath
        Try
            If Right(strPath, 1) <> "\" Then
                Return strPath & "\"
            Else
                Return strPath
            End If

        Catch ex As Exception

        End Try
    End Function

    Public Function returnPathWithoutBackslash(strPath As String) As String
        returnPathWithoutBackslash = strPath
        Try
            If strPath.EndsWith("\") Then
                Return strPath.Substring(0, strPath.Length - 1)
            Else
                Return strPath
            End If

        Catch ex As Exception

        End Try
    End Function

End Module