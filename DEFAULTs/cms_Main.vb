
Imports Gurock.SmartInspect
Imports System.Runtime.Serialization



<Serializable()> <DataContract()>
Public Class Main

#Region "VARs"

    'Public VARs for direct Access (Fields):
    <DataMember(EmitDefaultValue:=False, Order:=11)>
    Public Stage As Main.StageType? = Nothing

    '<DataMember(EmitDefaultValue:=False, Order:=99)>
    <NonSerialized>
    Public CTN As New Threading.CancellationToken

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public [Error] As cms_Error = Nothing

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Log As ALECS_LOG = Nothing

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public HELPER As cms_Helper = Nothing


    <Serializable()> <DataContract()>
    Public Enum ApplicationType
        Forms
        Console
        Web
        Service
    End Enum

    <Serializable()> <DataContract()>
    Public Enum StageType
        ALL     ' or better None
        DEV
        TEST
        RC
        PROD
        VIP
        'more Stages?
    End Enum


    'Private VARs (Access only via Properties): 

    'Config
    ''Private _Config As xml_Config = Nothing

    'AppPath:
    Private _AppPath As IO.DirectoryInfo = Nothing

    Private _DevelopmentArea As Boolean? = Nothing


    'Application Type:
    Private _ApplicationType As Main.ApplicationType = ApplicationType.Console
    Private _ApplicationUID As Guid? = Nothing
    Private _UnitTest As Boolean? = False

    <NonSerialized>
    Private _ThreadList As New System.Collections.Generic.List(Of Threading.Thread)

#End Region

#Region "Properties"

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public ReadOnly Property ProductiveClient As Boolean
        Get
            Return Not isDevelopmentArea()
        End Get
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Property AppType As ApplicationType
        Get
            Return _ApplicationType
        End Get
        Set(value As ApplicationType)
            _ApplicationType = value
        End Set
    End Property

    Public ReadOnly Property ApplicationUID As Guid
        Get
            If _ApplicationUID.HasValue Then Return _ApplicationUID

            Dim uidAttribute As System.Runtime.InteropServices.GuidAttribute = CType(System.Reflection.Assembly.GetExecutingAssembly.GetCustomAttributes(GetType(System.Runtime.InteropServices.GuidAttribute), True)(0), System.Runtime.InteropServices.GuidAttribute)
            If Not String.IsNullOrEmpty(uidAttribute.Value) Then
                _ApplicationUID = New Guid(uidAttribute.Value)
            Else
                _ApplicationUID = Guid.Empty
            End If
            Return _ApplicationUID
        End Get
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public ReadOnly Property AppPath As IO.DirectoryInfo
        Get
            AppPath = Nothing
            If HELPER IsNot Nothing Then
                Return New IO.DirectoryInfo(HELPER.AppPath())
            End If
        End Get
    End Property

    ''' <summary>
    ''' specifies if this Run is a UnitTest
    ''' --> Only has to set to true in UnitTests in developer IDE's
    ''' </summary>
    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Property UnitTest As Boolean
        Get
            If _UnitTest.HasValue Then
                Return _UnitTest
            Else
                Return False
            End If
        End Get
        Set(value As Boolean)
            _UnitTest = value
#Disable Warning BC40000 ' Type or member is obsolete
            If value Then
                '# TRUE
                Log.SIS.LogSeparator()
                'AppInDebugMode = True
                'Log.AppInDebugMode2 = True
                'Log.setLogLevelPipe(Level.Debug)
                'Log.Clear()
            Else
                '# FALSE
                'AppInDebugMode = False
                'Log.AppInDebugMode2 = False
                'Log.setLogLevelPipe(Level.Message)
                'ALL = Level.Debug 
                'Log.add("The UnitTest ist set to FALSE by code. UnitTest may only activated by code but NEVER deactivated by code. Please check Code!", Me, DefaultLog.ActWarning)
            End If
#Enable Warning BC40000 ' Type or member is obsolete
        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public ReadOnly Property ThreadList As System.Collections.Generic.List(Of System.Threading.Thread)
        Get
            Return _ThreadList
        End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
        CTN = New Threading.CancellationToken
        INIT()
    End Sub

    Public Sub New(dirAppPath As IO.DirectoryInfo)
        _AppPath = dirAppPath
        INIT()
    End Sub

    ''' IMPORTANT: Level.Debug must stay here!!! Main Only!
    Public Sub New(dirAppPath As IO.DirectoryInfo, Optional bolDebug As Boolean = False, Optional bolClear As Boolean = False, Optional bolService As Boolean = False)
        _AppPath = dirAppPath
        INIT()
    End Sub

    Private Sub INIT()
        Dim timMain As New Timing()
        Try
            CTN = New Threading.CancellationToken

            SiAuto.Main.EnterMethod(Level.Debug, "start Initializing cms_HELPER --->")
            HELPER = New cms_Helper(Me)
            SiAuto.Main.LeaveMethod(Level.Debug, "finished Initializing cms_HELPER <---")

            Log = New ALECS_LOG(Me)

            Dim silConnectionsBuilder As New ConnectionsBuilder

            '# the Defaults: LEVEL
            Log.SIS.Level = Level.Message
            SiAuto.Si.Level = Level.Message
            Log.SIS.AddCheckpoint(Level.Warning, "setting LOG for ALECS_JobS!")
            Log.SIS.Level = Level.Debug
            SiAuto.Si.Level = Level.Debug
            SiAuto.Si.SessionDefaults.Level = Level.Debug



            ' NAME
            SiAuto.Si.AppName = ProductName()

            '# Make a TEMP Init for Gurock.SI:

            'SiAuto.Si.Enabled = True
            ''myMain.Log.SIS.ClearAll()
            'QUIET in normal Mode: myMain.Log.SIS.LogMessage("SI is initialized and ready!")


            'INIT cms_HELPER Class: on Place 0
            Log.SIS.EnterMethod(Level.Debug, "start Initializing cms_HELPER --->")
            HELPER = New cms_Helper(Me)
            Log.SIS.LeaveMethod(Level.Debug, "finished Initializing cms_HELPER <---")


            Log.SIS.LogDebug("APPType=" & _ApplicationType.ToString & ", ApplicationType=" & _ApplicationType)


            'INIT Log Class first:
            Log.SIS.EnterMethod(Level.Debug, "start Initializing ALECS_LOG --->")
            Log = New ALECS_LOG(Me)
            Log.SIS.LeaveMethod(Level.Debug, "finished Initializing ALECS_LOG <---")

            'INIT Error Class second:
            Log.EnterMethod("start Initializing ALECS_Error --->", Level.Debug)
            [Error] = New cms_Error(Me)
            Log.LeaveMethod("finished Initializing ALECS_Error <---", Level.Debug)

            ''INIT Config Class third:
            'If AppInDebugMode Then Log.EnterMethod(Level.Debug, "start Initializing cms_Config --->")
            '_Config = New xml_Config(Me)
            'If AppInDebugMode Then Log.LeaveMethod(Level.Debug, "finished Initializing cms_Config <---")
            '_Config.INIT()

            ' Log.add("Hello and welcome to CMSLOG :-)", Me, Level.Debug, System.Drawing.Color.LightBlue)

            '# Stage:
            ''Stage = detectStage(Me)
            Stage = Main.StageType.DEV


            Log.add("* INIT * Main Inited by itself!" & vbNewLine & "Runtime=" & timMain.Duration(Timing.TimingFormat.Runtime), Me, Level.Debug, Drawing.Color.Aquamarine)


        Catch ex As Exception
            If Not [Error] Is Nothing Then
                [Error].getException(ex, "cms_Main:New", Level.Fatal)
            Else
                Debug.WriteLine(ex.Message)
                Debug.WriteLine("----------------------")
                'Log2File(dirAppPath.FullName & myMain.ProductName() & ".log", "!!! FATAL !!!", ex.Message.ToString, , , "cms_Main:New")
            End If
        Finally
            Console.WriteLine("the MAIN class is NEWed")
            'myMain.Log.SIS.LeaveMethod(Me, "New")
        End Try
    End Sub

#End Region


#Region "Functions"

    ''' <summary>
    ''' checks if Machine is in the 'DEVELOPMENT' Area
    ''' </summary>
    ''' <returns>True=is a development machine OR False=no, it's a productive system</returns>
    Public Function isDevelopmentArea() As Boolean
        Return True
    End Function

    Public Sub MP(intNumber As Integer)
        DP(intNumber, Level.Message)
    End Sub

    Public Sub DP(intNumber As Integer, Optional logLevel2 As Level? = Nothing, Optional objME As Object = Nothing, Optional objVariable As Object = Nothing)
        Try
            If objME Is Nothing Then objME = String.Empty
            Dim strName As String = "DP" & intNumber.ToString

            Dim logLevel As Level = SiAuto.Si.DefaultLevel
            If Not logLevel2.HasValue Then logLevel = SiAuto.Si.DefaultLevel

            If objVariable Is Nothing Then
                Debug.WriteLine("DP " & intNumber.ToString)
            Else
                Debug.WriteLine("DP " & intNumber.ToString & ": " & objVariable.ToString)
                If Not String.IsNullOrEmpty(objVariable.ToString) Then strName = objVariable.GetType.ToString
                Select Case True
                    Case TypeOf (objVariable) Is Boolean
                        'Log.SIS.WatchBool(logLevel, strName, objVariable)
                        'Log.SIS.LogBool(logLevel, strName, objVariable)

                    Case TypeOf (objVariable) Is Byte
                        'Log.SIS.WatchByte(logLevel, strName, objVariable)
                        'Log.SIS.LogByte(logLevel, strName, objVariable)

                    Case TypeOf (objVariable) Is String
                        'Log.SIS.WatchString(logLevel, strName, objVariable)
                        'Log.SIS.LogString(logLevel, strName, objVariable)

                    Case Else
                        'Log.SIS.Watch(logLevel, strName, objVariable)
                        'Log.SIS.LogObject(logLevel, strName, objVariable)

                End Select
            End If
            'System.Threading.Thread.Sleep(0)


        Catch ex As Exception
            '[Error].getException(ex, Me)
        End Try
    End Sub

#Region "Stage(s)"

    ''' <summary>
    ''' converts Stages as STRING to a valid Stage
    ''' </summary>
    ''' <param name="strValue">the Stage as string</param>
    ''' <returns>Main.StageType or Nothing</returns>
    Public Function PropertyStages(strValue As String) As StageType?
        PropertyStages = Nothing
        Try
            If String.IsNullOrEmpty(strValue) Then
                PropertyStages = Nothing
            Else
                Select Case strValue.ToLower
                    Case "prod"
                        PropertyStages = Main.StageType.PROD
                    Case "rc"
                        PropertyStages = Main.StageType.RC
                    Case "test"
                        PropertyStages = Main.StageType.TEST
                    Case "dev"
                        PropertyStages = Main.StageType.DEV
                    Case "all", "none"
                        PropertyStages = Main.StageType.ALL
                    Case "vip"
                        PropertyStages = Main.StageType.VIP
                    Case Else
                        Throw New ArgumentOutOfRangeException("PropertyStages", "the given value: '" & strValue & "' is an INVALID string and can NOT convert to a valid/known Stage!")
                End Select
            End If


        Catch ex As Exception
            '[Error].getException(ex, Me)
        End Try
    End Function

#End Region

#Region "Product XXX"

    Public Function ProductName() As String
        Return System.Reflection.Assembly.GetExecutingAssembly.GetName.Name()
    End Function

    Public Function ProductVersion() As Version
        Return System.Reflection.Assembly.GetExecutingAssembly.GetName.Version
    End Function

    'Public Function ProductDate() As Date
    '    ProductDate = Nothing
    '    Try
    '        Dim myAge As IO.FileInfo = Nothing
    '        Select Case AppType
    '            Case Main.ApplicationType.Web
    '                myAge = New IO.FileInfo(HELPER.AppPath(True) & "bin\" & ProductName() & ".dll")
    '            Case Else
    '                myAge = New IO.FileInfo(HELPER.AppPath(True) & ProductName() & ".exe")
    '        End Select
    '        Return myAge.LastWriteTime.ToLocalTime()


    '    Catch ex As Exception
    '        '[Error].getException(ex, Me)
    '    End Try
    'End Function

#End Region

#End Region

End Class
