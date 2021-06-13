
Imports System.Runtime.Serialization
Imports Microsoft.Win32
Imports Gurock.SmartInspect



<Serializable()> <DataContract()>
Public Class ALECS_Host_Entry
    Inherits ALECS_Template_Entry_FAKE
    Implements IALECS_Entry, IALECS_Entry.ILoadable

#Region "VARs"

    'Public VARs for direct Access (Fields):
    <DataMember(EmitDefaultValue:=False, Order:=22)>
    Public SystemName As String = Nothing
    'in PROP: <DataMember(EmitDefaultValue:=False, Order:=23)>
    Public SystemType As SystemType2? = Nothing

    <DataMember(EmitDefaultValue:=False, Order:=51)>
    Public PrimaryIP As String = Nothing
    'without <DataMember>, because is used in Property Stages:
    Public Stage As Main.StageType? = Nothing

    <DataMember(EmitDefaultValue:=False, Order:=83)>
    Public Domain As String = Nothing

    Public LastBoot As DateTime? = Nothing

    <Serializable()> <DataContract()>
    Public Enum SystemType2
        client

        'server
        remotedesktop
        remoteapp
        domain
        web
        database
        app

        'router
        cloudgateway
        'peripheral

        nas

        'hosts / ewms
        host

    End Enum


    'Private VARs (Access only via Properties):
    Private _PROPCache As New NET_Cache
    Private _ContainerCacheOld As New NET_Cache
    Private _OSCache As New NET_Cache
    Private _AIL As New NET_Cache
    Private _ARDSC As New NET_Cache

    '-- private fields for ALECS_TemplateClass_DB:
    Private _BaseObjectName As String = GetType(ALECS_Host_Entry).Name
    <NonSerialized>
    Private drItem As DataRow = Nothing

#End Region

#Region "Properties"

    ''' <summary>
    ''' Here we use the full one, because in Dahsboard only has to be placed the REAL-SPEECH one (Stages)
    ''' </summary>
    '''' <returns>a functionality string in ALECS_JobS Dashboard</returns>
    <DataMember(EmitDefaultValue:=False, Order:=71)>
    Public Property Stages As String
        Get
            Return Stage.ToString
        End Get
        Set(value As String)
            Stage = myMain.PropertyStages(value)
        End Set
    End Property

#End Region

#Region "Constructors"

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New()
        MyBase.New(GetType(ALECS_Host_Entry))
        INIT()
    End Sub

    '# INITed by ALECS_JOB Queue or DeserializeObject (of Childs)
    Public Sub New(objType As Type)
        MyBase.New(objType)
        INIT()
    End Sub

    '# INITed by UnitTests
    Public Sub New(ByRef objMain As Main, objType As Type)
        MyBase.New(objMain, objType)
        INIT()
    End Sub

    ''' <summary>
    ''' QUICK aheItem constructor. NOT for Child classes!
    ''' </summary>
    Public Sub New(ByRef objMain As Main, objUID As Guid)
        MyBase.New(objMain, GetType(ALECS_Host_Entry), objUID)
        INIT()
    End Sub

    ''' <summary>
    ''' ChildClass constructor
    ''' </summary>
    Public Sub New(ByRef objMain As Main, objType As Type, objUID As Guid)
        MyBase.New(objMain, objType, objUID)
        INIT()
    End Sub

    Private Sub INIT()
        'something special

    End Sub

#End Region

#Region "Functions"

#Region "IALECS_Entry Functions"

    Public Overrides Function Load() As Boolean Implements IALECS_Entry.ILoadable.Load
        Load = False
        Dim timTiming As New Timing
        clearLastError()
        Try


            setLoaded()

            Load = True

            myMain.Log.SIS.LogObject("LOADED " & _BaseObjectName & " object: " & getDisplayName(), Me)


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Overrides Function getDisplayName() As String Implements IALECS_Entry.getDisplayName
        getDisplayName = SystemName
        If String.IsNullOrEmpty(getDisplayName) Then getDisplayName = Name
        If Not String.IsNullOrEmpty(PrimaryIP) Then
            getDisplayName += " (" & PrimaryIP & ")"
        Else
            getDisplayName += " (" & UID.ToString() & ")"
        End If
    End Function

    Public Overrides Sub DebugMe(Optional strPREfix As String = "") Implements IALECS_Entry.DebugMe
        Try
            myMain.Log.SIS.LogColored(Color.LightGray, "DEBUGME: " & getDisplayName())
            myMain.Log.SIS.LogObject(strPREfix & ChildObjectName & ": " & getDisplayName(), Me)

#Region "Details"
            Dim strDetails As String = "DETAILS"
            myMain.Log.SIS.EnterMethod(strDetails)
            If String.IsNullOrEmpty(strPREfix) Then strPREfix = "AHE "

            myMain.Log.SIS.LeaveMethod(strDetails)
#End Region


        Catch ex As Exception

        End Try
    End Sub

#End Region

    Public Function detectMySelf() As Boolean
        detectMySelf = False
        clearLastError()
        Try
            UID = New Guid("4bb1b48b-2d32-4f53-97d0-5433fa985d8e")
            Name = "TestHost"
            PrimaryIP = "192.168.1.1"
            Stage = Main.StageType.DEV


            If Not Load() Then getLastError()

            myMain.Log.SIS.LogObject(Level.Debug, "detectMySelf returns: " & getDisplayName(), Me)

            detectMySelf = True


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    ''' <summary>
    ''' check (via PING) if a host is reachable
    ''' </summary>
    ''' <param name="intRetries"></param>
    Public Function checkReachability(Optional intRetries As Integer = 10) As Boolean
        checkReachability = False
        clearLastError()
        Try
            Dim myNIC As New System.Net.NetworkInformation.Ping, retPING As System.Net.NetworkInformation.PingReply, intPingCount As Integer = 0
            For i As Integer = 1 To intRetries
                retPING = myNIC.Send(PrimaryIP)
                If retPING.Status = Net.NetworkInformation.IPStatus.Success Then
                    intPingCount += 1
                    myMain.Log.SIS.LogLong("retPING.RoundtripTime", retPING.RoundtripTime)
                    Return True
                End If
                System.Threading.Thread.Sleep(1000)
            Next


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

#End Region

End Class




<Serializable()> <DataContract()>
Public Class ALECS_HOST_Selection
    Inherits ALECS_Template_Entry_FAKE
    Implements IALECS_Entry

#Region "VARs"

    'Public VARs for direct Access (Fields):




    <Serializable()> <DataContract()>
    Public Enum GroupType
        manual

        preset

    End Enum


    'Private VARs (Access only via Properties):
    Private _Host As New ALECS_Host_Entry(myMain, GetType(ALECS_Host_Entry))

    '-- private fields for ALECS_TemplateClass_DB:
    Private _BaseObjectName As String = GetType(ALECS_HOST_Selection).Name

#End Region

#Region "Properties"

    'overwrite this (so can build it automaticly)
    <DataMember(EmitDefaultValue:=False, Order:=11)>
    Public Overridable Shadows ReadOnly Property Name As String 'Implements IALECS_Entry.Name
        Get
            Select Case True

                Case myMain.HELPER.NotAnEmptyValue(_Host.UID)
                    Name = "Selection by single Host (UID)"

                Case Not String.IsNullOrEmpty(_Host.SystemName)
                    If _Host.SystemName.Contains("%") Then
                        Name = "Selection by a Part (%) of SystemName"
                    Else
                        Name = "Selection by SystemName"
                    End If

                Case Else
                    Name = "EMPTY (new?)"

            End Select
        End Get
    End Property

    'OFF <DataMember(EmitDefaultValue:=False, Order:=31)>
    Public Property Stage As Main.StageType?
        Get
            'If _Host Is Nothing Then Return Nothing
            Return _Host.Stage
        End Get
        Set(value As Main.StageType?)
            _Host.Stage = value
        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=31)>
    Public Property Stages As String
        Get
            'If _Host Is Nothing Then Return Nothing
            Return _Host.Stages
        End Get
        Set(value As String)
            If Not String.IsNullOrEmpty(value) Then
                Stage = myMain.PropertyStages(value)
            Else
                Stage = Nothing
            End If
        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=32)>
    Public Property SystemName As String
        Get
            'If _Host Is Nothing Then Return Nothing
            Return _Host.SystemName
        End Get
        Set(value As String)
            '# replace * by % in SQL Syntax
            If Not String.IsNullOrEmpty(_Host.SystemName) AndAlso _Host.SystemName.Contains("*") Then _Host.SystemName = System.Text.RegularExpressions.Regex.Replace(_Host.SystemName, "[*]", "%").ToLower
            _Host.SystemName = value
        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=33)>
    Public Property Domain As String
        Get
            'If _Host Is Nothing Then Return Nothing
            Return _Host.Domain
        End Get
        Set(value As String)
            _Host.Domain = value
        End Set
    End Property

#End Region

#Region "Constructors"

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New()
        MyBase.New(GetType(ALECS_HOST_Selection))    'inited in BASE Constructor
    End Sub

#End Region

#Region "Functions"

#Region "IALECS_Entry Functions"

    Public Overrides Function getDisplayName() As String Implements IALECS_Entry.getDisplayName
        Return Name
        'getDisplayName = Name & "[" & _Host.getDisplayName & "]"
        'If Not String.IsNullOrEmpty(Comment) Then getDisplayName += " " & Comment
    End Function

    Public Overrides Sub DebugMe(Optional strPREfix As String = "") Implements IALECS_Entry.DebugMe
        Try
            myMain.Log.EnterMethod(Me)
            myMain.Log.SIS.LogColored(Color.LightGray, "DEBUGME: " & getDisplayName())
            myMain.Log.SIS.LogObject(strPREfix & _BaseObjectName & ": " & getDisplayName(), Me)

            '#Region "Details"
            '                Dim strDetails As String = "DETAILS"
            '                myMain.Log.EnterMethod(strDetails)
            '                If String.IsNullOrEmpty(strPREfix) Then strPREfix = "AHE "
            '                'ACE:
            '                If _ACE IsNot Nothing Then
            '                    myMain.Log.SIS.LogObject("..._ACE", _ACE)
            '                    _ACE.DebugMe(strPREfix)
            '                Else
            '                    myMain.Log.add("_ACE is Nothing", Me)
            '                End If
            '                myMain.Log.LeaveMethod(strDetails)
            '#End Region


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
        Finally
            myMain.Log.LeaveMethod(Me)
        End Try
    End Sub

#End Region

#End Region

End Class