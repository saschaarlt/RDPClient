
Imports System.Runtime.Serialization
Imports Gurock.SmartInspect



<Serializable()> <DataContract()>
Public Class ALECS_X_Action
    Inherits ALECS_Template_Entry_FAKE
    Implements IALECS_Entry, IALECS_Entry.ILoadable, IALECS_Entry.ISaveable, IALECS_Entry.IStateable, IALECS_Entry.IDeletable, IALECS_Entry.ICreateable

#Region "VARs"

    'Public VARs for direct Access (Fields):
    <DataMember(EmitDefaultValue:=False, Order:=21)>
    Public Category As String = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=22)>
    Public Script As String = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=23)>
    Public Type As XType? = Nothing
    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Group As ALECS_Group_Entry = Nothing
    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Icon As String = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=89)>
    Public Timeout As TimeSpan? = Nothing


    '# the ALECS_References (default) in correct order
    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=70)>
    Public Argument_Collection As Boolean? = Nothing
    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=72)>
    Public Argument_RDSC As Boolean? = Nothing
    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=73)>
    Public Argument_Host As Boolean? = Nothing
    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=74)>
    Public Argument_APP As Boolean? = Nothing
    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=75)>
    Public Argument_Tenant As Boolean? = Nothing
    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=76)>
    Public Argument_User As Boolean? = Nothing

    <DataMember(EmitDefaultValue:=False, Order:=86)>
    Public Argument001 As String = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=87)>
    Public Argument002 As String = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=88)>
    Public Argument003 As String = Nothing

    '# the following vars are for jobs nad not stored in db:
    <DataMember(EmitDefaultValue:=False, Order:=80)>
    Public Collection As ALECS_Collection = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=81)>
    Public RDSC As ALECS_RDSC_Entry = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=82)>
    Public HostSelection As ALECS_HOST_Selection = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=83)>
    Public AppSelection As Dictionary(Of Guid, ALECS_APP_Entry) = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=84)>
    Public TenantSelection As Dictionary(Of Guid, ALECS_Tenant_Entry) = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=85)>
    Public UserSelection As Dictionary(Of Guid, ALECS_User_Entry) = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Reason As String = Nothing


    <Serializable()> <DataContract()>
    Public Enum XType
        ALECS_Job
        ALECS_PROCESS_Windows
        CMD
        'EXE
        ACMP
        PowerShellCommand
        PowerShellScript

    End Enum

    'Private VARs (Access only via Properties):
    'Private Const _PSEPath As String = Network_Storages.IRZ13.ALECS_Storage & "PROD\PowerShellExecution\"
    Private Const _CategorySeperator As String = "|"

    '-- private fields for ALECS_TemplateClass_DB:
    Private _BaseObjectName As String = GetType(ALECS_X_Action).Name
    <NonSerialized>
    Private drItem As DataRow = Nothing

#End Region

#Region "Properties"

    ''' <summary>
    ''' THIS Property in only for better Dashboard-View and has NO functionality
    ''' it must be removed in Constructor.INIT
    ''' </summary>
    ''' <returns>a functionality string in ALECS_JobS Dashboard</returns>
    <DataMember(EmitDefaultValue:=False, Order:=13)>
    Public Property Types As String
        Get
            Return Type.ToString
        End Get
        Set(value As String)
            Type = myMain.HELPER.Enum(Of XType)(value)
        End Set
    End Property

#End Region

#Region "Constructors"

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New()
        MyBase.New(GetType(ALECS_X_Action))    'inited in BASE Constructor
        INIT()
    End Sub

    '# INITed by UnitTests
    Public Sub New(ByRef objMain As Main)
        MyBase.New(objMain, GetType(ALECS_X_Action))    'inited in BASE Constructor
        INIT()
    End Sub

    Public Sub New(ByRef objMain As Main, uidGuid As Guid)
        MyBase.New(objMain, GetType(ALECS_X_Action), uidGuid)
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
            myMain.Log.EnterMethod(Me)

            '#1 load BASE object from database (Inherited ALECS_TemplateClass_DB from ALECS_Class_Template.vb):
            drItem = BaseLoad()
            If drItem Is Nothing Then Throw getLastError()

            'set Categories to UpperCase:
            Category = Category.ToUpper


            setLoaded()

            Load = True

            myMain.Log.SIS.LogObject(Level.Debug, "LOADED " & _BaseObjectName & " object: " & getDisplayName(), Me)

            myMain.Log.EndsSuccessfully(Me)


        Catch ex As OperationCanceledException
            myMain.Log.OperationCanceledException(Me)
            setLastError(ex)
            Throw
        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me, timTiming)
        End Try
    End Function

    Public Overrides Function Save() As Boolean Implements IALECS_Entry.ISaveable.Save
        Save = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            myMain.Log.EnterMethod(Me)

            '#1 build special Fields "CATEGORY" for:
            Category = Category.ToUpper

            '#2 save BASE object to database (Inherited ALECS_TemplateClass_DB from cms_EntryInterfaces.vb):
            If Not BaseSave() Then Throw getLastError()

            setLoaded()

            Save = True

            myMain.Log.SIS.LogObject(Level.Debug, "SAVED " & _BaseObjectName & " object: " & getDisplayName(), Me)

            myMain.Log.EndsSuccessfully(Me)


        Catch ex As OperationCanceledException
            myMain.Log.OperationCanceledException(Me)
            setLastError(ex)
            Throw
        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me, timTiming)
        End Try
    End Function

    Public Overrides Function getDisplayName() As String Implements IALECS_Entry.getDisplayName
        If String.IsNullOrEmpty(Name) Then Return String.Empty
        Return Name
    End Function

    Public Overrides Sub DebugMe(Optional strPREfix As String = "") Implements IALECS_Entry.DebugMe
        Try
            myMain.Log.EnterMethod(Me)
            myMain.Log.SIS.LogColored(Color.LightGray, "DEBUGME: " & getDisplayName())
            myMain.Log.SIS.LogObject(strPREfix & _BaseObjectName & ": " & getDisplayName(), Me)

#Region "Details"
            Dim strDetails As String = "DETAILS"
            myMain.Log.EnterMethod(strDetails)
            If String.IsNullOrEmpty(strPREfix) Then strPREfix = "AXA "

            myMain.Log.DebugMeDefaultReferences(Me, strPREfix)
            'additional special HostSelection:
            Try
                If HostSelection IsNot Nothing Then
                    'Dim aheDebug As ALECS_Host_Entry = Host
                    'myMain.Log.SIS.LogObject("...HostSelection", HostSelection)
                    HostSelection.DebugMe()
                End If
            Catch ex As Exception
                'NO Host
            End Try
            'additional special AppSelection:
            Try
                If AppSelection IsNot Nothing Then
                    myMain.Log.SIS.LogDictionary("...AppSelection", AppSelection)
                End If
            Catch ex As Exception
                'NO APP
            End Try
            'additional special TenantSelection:
            Try
                If TenantSelection IsNot Nothing Then
                    myMain.Log.SIS.LogDictionary("...TenantSelection", TenantSelection)
                End If
            Catch ex As Exception
                'NO TENANT
            End Try
            'additional special UserSelection:
            Try
                If UserSelection IsNot Nothing Then
                    myMain.Log.SIS.LogDictionary("...UserSelection", UserSelection)
                End If
            Catch ex As Exception
                'NO TENANT
            End Try
            myMain.Log.LeaveMethod(strDetails)
#End Region


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me)
        End Try
    End Sub

#End Region

    Public Function Execute(aheDestination As ALECS_Host_Entry) As Boolean
        Execute = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            myMain.Log.EnterMethod(Me)

            If Not isLoaded() Then If Not Load() Then Throw getLastError()

            If Not Active Then Throw New InvalidOperationException("You try to .Execute " & getDisplayName() & ", but this Xaction is NOT Active at the moment. So please .Activate the 'ALECS_X_Action' first and then retry.")
            If Not Type.HasValue Then Throw New ArgumentNullException("Type", "The TYPE of areItem (ALECS_X_Action) is NOTHING. Please specify Type first!")

            myMain.Log.SIS.LogObject("Execute: " & getDisplayName(), Me)


            Select Case Type

                Case XType.ALECS_PROCESS_Windows
                    Execute = EXEcuteAPW(aheDestination)

                    'Case ExecutionType.CMD
                    '    Exit Select

                Case XType.PowerShellScript, XType.PowerShellCommand
                    Execute = EXEcutePSE()

                Case XType.ACMP
                    Execute = EXEcuteACMP()

                Case XType.ALECS_Job
                    Throw New ArgumentException("Type", "The given TYPE of areItem (ALECS_X_Action) is '" & Type.ToString & "'. That's an type which is NOT used by .Execution. Please check why we've ended here.")

                Case Else
                    Throw New ArgumentOutOfRangeException("Type", "The given TYPE of areItem (ALECS_X_Action) is '" & Type.ToString & "'. That's an UNKNOWN Type. Please specify a KNOW Type.")

            End Select

            myMain.Log.EndsSuccessfully(Me)


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me, timTiming)
        End Try
    End Function

#Region "EXEcution HELPER (APW, ACMP, etc.)"

    ''' <summary>
    ''' executes a ALECS_PROCESS_Windows (from DB)
    ''' </summary>
    ''' <returns>boolean if worked</returns>
    Private Function EXEcuteAPW(aheDestination As ALECS_Host_Entry) As Boolean
        EXEcuteAPW = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            myMain.Log.EnterMethod(Me)

            Throw New NotImplementedException("")

            EXEcuteAPW = True

            myMain.Log.EndsSuccessfully(Me)


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me, timTiming)
        End Try
    End Function

    ''' <summary>
    ''' executes a ACMP Script via RunCC.exe
    ''' </summary>
    ''' <returns>boolean if worked</returns>
    Private Function EXEcuteACMP() As Boolean
        EXEcuteACMP = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            myMain.Log.EnterMethod(Me)

            Throw New NotImplementedException("")

            EXEcuteACMP = True

            myMain.Log.EndsSuccessfully(Me)


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me, timTiming)
        End Try
    End Function

    Private Function EXEcutePSE() As Boolean
        EXEcutePSE = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            myMain.Log.EnterMethod(Me)

            Throw New NotImplementedException("")

            EXEcutePSE = True

            myMain.Log.EndsSuccessfully(Me)


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me, timTiming)
        End Try
    End Function

#End Region

#End Region

End Class
