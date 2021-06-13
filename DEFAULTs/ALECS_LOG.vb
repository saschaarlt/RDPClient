
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Runtime.Serialization
Imports Gurock.SmartInspect



<Serializable()> <DataContract()>
Public Class ALECS_LOG
    Inherits cms_LastError

#Region "VARs"

    '# NO myMain usage in this class please
    'Dim myMain As Main

    'Public VARs for direct Access (Fields):
    <Serializable()> <DataContract()>
    Public Enum DefaultLog

        [Default]
        Successful
        Warning
        Failed

        ActStart
        ActEnd
        ActSuccessful
        ActWarning
        ActFailed

        MileStoneSuccessful
        MileStoneFailed

        UnitStart
        UnitEnd
        UnitSuccessful
        UnitFailed

    End Enum


    'Private VARs (Access only via Properties):
    Private _UID As Guid = Guid.NewGuid
    Private _AppInDebugMode As Boolean = False

    'Private _silLogLevel As Level = Level.Debug is NoW stored in:
    'myMain.Log.SIS.Level

    '#Fake VARs instead of myMain (initialized in New())
    Private newMain As Main = Nothing
    Private HELPER As cms_Helper = Nothing
    Private myError As cms_Error = Nothing

    <NonSerialized>
    Private _CBFile As New ConnectionsBuilder, _SILSession As Session = SiAuto.Main

#End Region

#Region "Properties"

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Property AppInDebugMode2 As Boolean
        Get
            Return _AppInDebugMode
        End Get
        Set(value As Boolean)
            If _AppInDebugMode.Equals(value) Then Exit Property

            _AppInDebugMode = value

#Disable Warning BC40000 ' Type or member is obsolete

            If value Then

                '# Later is Private!
                setLogLevelPipe(Level.Debug)

            Else

                '# Later is Private!
                setLogLevelPipe(Level.Message)

            End If

#Enable Warning BC40000 ' Type or member is obsolete

        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public ReadOnly Property SIS As Session
        Get
            Return _SILSession
        End Get
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=1)>
    Public ReadOnly Property UID As Guid
        Get
            Return _UID
        End Get
    End Property

    Public ReadOnly Property LogLevel As Level
        Get
            Return _SILSession.Level
        End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByRef objMain As Main)
        MyBase.New()
        newMain = objMain
        HELPER = objMain.HELPER

        'is not initialized, yet!
        myError = New cms_Error(objMain)

        '# Init ME (AUto)
        INIT()
    End Sub

    Public Sub INIT(Optional bolRuntimeFirst As Boolean = False)
        clearLastError()
        Try
            ' ##########################
            ' ### LOG Initialisieren ###
            ' ##########################


            '##### YES we're logging: #####

            _CBFile = New ConnectionsBuilder

            '# build FILE Logging:
            _CBFile.BeginProtocol("file")
            With _CBFile

                '# always append:
                .AddOption("append", True)

                '# always full logging for file:
                .AddOption("level", Level.Debug)

                '# Path + Filename:
                Dim strFile As String = HELPER.AppPath(True) & "LOG\" & newMain.ProductName & ".sil"
                Select Case True
                    Case Environment.MachineName.ToLower.Contains("web")
                        '# SPECIALs for WebServer:
                        'Dim dirPath As New ALECS_Directory_Entry(newMain, New IO.DirectoryInfo(System.AppDomain.CurrentDomain.BaseDirectory)), because LOG is NOT initialized, yet!
                        Dim ioPath As New IO.DirectoryInfo(System.AppDomain.CurrentDomain.BaseDirectory)
                        _SILSession.LogObject("dirPath", ioPath)
                        strFile = "d:\Logs\" & ioPath.Name & "\" & newMain.ProductName & ".sil"
                        _SILSession.LogString("special strFile for: " & Environment.MachineName.ToLower, strFile)

                End Select
                _SILSession.LogString(Level.Debug, "The ultimate Log File (with Path)", strFile)

                '# create LogPath if NOT exists
                If Not String.IsNullOrEmpty(strFile) Then

                    'check if LogPath exists:
                    'NO ALECS_FSO usage possible here Dim filLog As New ALECS_File_Entry(newMain, New IO.FileInfo(strFile)), because LOG is NOT initialized, yet!

                    Dim ioLogPath As New IO.DirectoryInfo(New IO.FileInfo(strFile).DirectoryName)
                    '_SILSession.LogObject("special Path for: " & Environment.MachineName.ToLower, ioLogPath)
                    If Not ioLogPath.Exists Then
                        ioLogPath.Create()
                    End If

                    .AddOption("filename", strFile)

                End If

                '# SIL File Size:
                Dim dblFileStorageSize As Double = HELPER.ByteSize("75mb", cms_Helper.cms_Byte_Size.KB)
                .AddOption("maxsize", dblFileStorageSize.ToString)

                '# Parts:
                .AddOption("maxparts", 15)

                '# Buffer:
                Dim retBuffer As Double = HELPER.ByteSize("10kb", cms_Helper.cms_Byte_Size.KB)
                .AddOption("buffer", retBuffer.ToString)

                '# Rotate:
                '.AddOption("rotate", FileRotate.Daily)

            End With
            _CBFile.EndProtocol()



            '# default LogLevel is debug:
            SiAuto.Si.DefaultLevel = Level.Debug

            '# set the Sessions MASTEr Log.UID
            _SILSession = SiAuto.Si.AddSession(_UID.ToString)


            '### ==>> Activate File + Pipe Logging:
            ActivateConfig()


            add("LOG Module initialized", Me, Level.Message, Color.Bisque)


        Catch ex As Exception
            myError.getException(ex, Me)
            ' only POS with:
            setLastError(ex)
        Finally
            '# result:
            _SILSession.LogMessage("my INITed LogLevel is """ & _SILSession.Level.ToString & """")
        End Try
    End Sub

#End Region

#Region "Functions"

#Region "Default LOG"

    Public Sub add(strText As String, objMe As Object, Optional enmDefaultLog As ALECS_LOG.DefaultLog = ALECS_LOG.DefaultLog.Default)

        addLogLine(strText, objMe, getDefaultLevel(enmDefaultLog), getDefaultColor(enmDefaultLog))

    End Sub

    Public Sub add(strText As String, objMe As Object, silLevel As Level, Optional rowColor As Color = Nothing)

        addLogLine(strText, objMe, silLevel, rowColor)

    End Sub

    Private Sub addLogLine(strText As String, objMe As Object, logLevel As Level, Optional RowColor As Color = Nothing)
        Try
            'Debug Ausgabe:
            'Debug.WriteLine(strText)
            If objMe IsNot Nothing Then Debug.WriteLine(objMe.ToString)

            If RowColor = Nothing Then RowColor = Color.Empty
            If RowColor <> Color.Empty Then _SILSession.Color = RowColor

            Select Case logLevel
                Case Level.Fatal
                    _SILSession.LogFatal(strText)
                Case Level.Error
                    _SILSession.LogError(strText)
                Case Level.Warning
                    _SILSession.LogWarning(strText)
                Case Level.Message
                    _SILSession.LogMessage(strText)
                Case Level.Verbose
                    _SILSession.LogVerbose(strText)
                Case Else
                    _SILSession.LogDebug(strText)
            End Select

            If RowColor <> Color.Empty Then _SILSession.ResetColor()


        Catch ex As Exception
            'cms_System.Log2File(_UID, "!!!EMERGENCY ERROR!!!", ex.Message, , "cms_Logging.vb:add")
            myError.getException(ex, Me)
        End Try
    End Sub

#End Region

    Public Sub addCheckpoint(strMessage As String, Optional strDetails As String = "", Optional logLevel As Level = Level.Message)
        Try
            If String.IsNullOrEmpty(strDetails) Then
                _SILSession.AddCheckpoint(logLevel, strMessage)
            Else
                _SILSession.AddCheckpoint(logLevel, strMessage, strDetails)
            End If


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Sub EmptyConstructor(objME As Object)
        Try
            Dim myME As New cms_ME(objME)

            add("NEW " & myME.Name & " INITed by " & SiAuto.Si.AppName & " as an EMPTY Object!", Me, Level.Verbose, Color.DarkSalmon)


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Sub Clear()
        SIS.ClearAll()
    End Sub

#Region "SPECIAL Objects: logObject, XML, JSON, Enums, ACE Credentials, etc."

    Public Sub logObject(objObject As Object, Optional strName As String = "", Optional logLevel As Level? = Nothing)
        Try
            If String.IsNullOrEmpty(strName) Then
                strName = TryCast(objObject, String)
                If String.IsNullOrEmpty(strName) AndAlso objObject IsNot Nothing Then
                    Dim myType As Type = objObject.GetType()
                    strName = myType.Name
                End If
            End If
            If String.IsNullOrEmpty(strName) Then strName = "logObject"

            'If TypeOf objObject Is String Then
            '    _SILSession.LogString(logLevel, strName, objObject)
            'ElseIf TypeOf objObject Is Integer Then
            '    _SILSession.LogInt(logLevel, strName, objObject)
            'Else
            '    _SILSession.LogObject(logLevel, strName, objObject)
            'End If

            If objObject IsNot Nothing Then

                '_SILSession.LogObject("Type:", objObject.GetType)

                Select Case True
                    Case TypeOf (objObject) Is Boolean
                        SIS.LogBool(detectLogLevel(logLevel), strName, objObject)
                    Case TypeOf (objObject) Is String
                        SIS.LogString(detectLogLevel(logLevel), strName, objObject)
                    Case TypeOf (objObject) Is Integer
                        SIS.LogInt(detectLogLevel(logLevel), strName, objObject)
                    Case TypeOf (objObject) Is Date
                        SIS.LogDateTime(detectLogLevel(logLevel), strName, objObject)
                    Case TypeOf (objObject) Is CollectionBase
                        SIS.LogCollection(detectLogLevel(logLevel), strName, objObject)

                        'Multiple Types!!!

                        'Try as next to last:
                    Case TypeOf (objObject.ToString) Is String
                        SIS.LogString(detectLogLevel(logLevel), strName, objObject.ToString)
                    Case Else
                        ' unknown (not programed) Type:
                        SIS.LogObject("detected Type", objObject.GetType)
                        _SILSession.LogWarning("Can't identified object!")
                        SIS.LogObject(detectLogLevel(logLevel), strName, objObject)
                End Select

            Else

                add(strName & " = <Nothing>", Me, detectLogLevel(logLevel))

            End If


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

#Region "XML + JSON logging"

    'Public Sub logXML(strName As String, strXML As String, Optional logLevel As Level = Level.Debug)
    '    Try
    '        _SILSession.LogSource(logLevel, strName, PrettyXML(strXML), SourceId.Xml)


    '    Catch ex As Exception
    '        myError.getException(ex, Me)
    '        '_SILSession.LogWarning("Can't log strXML")
    '    End Try
    'End Sub

    'Public Sub logXML(strName As String, xmlItem As System.Xml.Linq.XElement, Optional logLevel As Level = Level.Debug)
    '    Try
    '        _SILSession.LogSource(logLevel, strName, PrettyXML(xmlItem.ToString), SourceId.Xml)


    '    Catch ex As Exception
    '        myError.getException(ex, Me)
    '        _SILSession.LogWarning("Can't log XElement")
    '    End Try
    'End Sub

    'Public Sub logXML(strName As String, xmlDocument As System.Xml.XmlDocument, Optional logLevel As Level = Level.Debug)
    '    Try
    '        _SILSession.LogSource(logLevel, strName, PrettyXML(xmlDocument.OuterXml), SourceId.Xml)


    '    Catch ex As Exception
    '        myError.getException(ex, Me)
    '        _SILSession.LogWarning("Can't log XmlDocument")
    '    End Try
    'End Sub

    Public Sub logJson(strName As String, strJson As String, Optional logLevel As Level = Level.Debug)
        Try
            _SILSession.LogSource(logLevel, strName, PrettyJSON(strJson), SourceId.Xml)


        Catch ex As Exception
            myError.getException(ex, Me)
            _SILSession.LogWarning("Can't log json")
        End Try
    End Sub

    Private Function PrettyJSON(strJSON As String) As String
        PrettyJSON = Newtonsoft.Json.Linq.JToken.Parse(strJSON).ToString(Newtonsoft.Json.Formatting.Indented)
    End Function

#End Region

    Public Sub logAssert(strDescription As String, Optional logLevel As Level? = Nothing)
        _SILSession.LogAssert(detectLogLevel(logLevel), False, strDescription)
    End Sub

    <Obsolete("This function is depricated. Please use logAssert function instead!", True)>
    Public Sub logExclamationmark(strDescription As String)
        _SILSession.LogAssert(False, strDescription)
    End Sub

    Public Sub logRuntimeSI(myTiming2 As Timing, Optional strName As String = "Runtime", Optional logLevel As Level? = Nothing)
        Try
            'Dim myME As New cms_ME(objMe)
            If myTiming2 IsNot Nothing Then
                _SILSession.LogString(detectLogLevel(logLevel), strName, FormatNumber(myTiming2.RuntimeMilliSeconds, 0, TriState.False, TriState.UseDefault, TriState.True) & " ms - " & vbNewLine & myTiming2.Duration(Timing.TimingFormat.Runtime))
            Else
                _SILSession.LogString(detectLogLevel(logLevel), strName, FormatNumber(myTiming.RuntimeMilliSeconds, 0, TriState.False, TriState.UseDefault, TriState.True) & " ms - " & vbNewLine & myTiming.Duration(Timing.TimingFormat.Runtime))
            End If


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    ''' <summary>
    ''' This add (on the quick way) an ALECS_Runtime_Entry for documentation and statistics
    ''' </summary>
    ''' <param name="myTiming2">an Timing object</param>
    ''' <param name="uidJob">every Runtime sequence has his own uid for statistics analytics IMPORTANT</param>
    ''' <param name="strName">statistics is not everything :-) sometimes a human has to read it (OPTIONAL, could be added per runjob later)</param>
    ''' <param name="bolResult">gives a boolean value if the Result of this runtime is successfully</param>
    Public Sub logRuntime(objMe As Object, myTiming2 As Timing, uidJob As Guid, Optional strName As String = Nothing, Optional bolResult As Boolean = True, Optional logLevel As Level? = Nothing)
        add("Sorry, logACE is NOT implemented in this project!", Me, DefaultLog.ActWarning)
    End Sub

    Public Sub logDataRow(drRow As DataRow, Optional strName As String = Nothing, Optional logLevel As Level? = Nothing)
        Try
            'Name?
            If String.IsNullOrEmpty(strName) Then
                If Not String.IsNullOrEmpty(drRow.Table.TableName) Then
                    'SIS.LogDataTable(drRow.Table.TableName, drRow.Table)
                Else
                    strName = String.Empty
                    'Dim dtItem As New DataTable()
                    'dtItem.Rows.Add(drRow)
                    'SIS.LogDataTable(strName, dtItem)
                End If
                _SILSession.LogDataTable(detectLogLevel(logLevel), drRow.Table.TableName, drRow.Table)
            Else
                _SILSession.LogDataTable(detectLogLevel(logLevel), strName, drRow.Table)
            End If

            'Dim dtItem As New DataTable()
            'dtItem.Rows.Add(drRow)
            'SIS.LogDataTable(strName, dtItem)


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Sub logGivenParameter(objParameter As Object, Optional strTitle As String = "", Optional logLevel As Level? = Nothing)
        Try
            If strTitle.Equals("") Then
                strTitle = TryCast(objParameter, String)
                If strTitle Is Nothing AndAlso objParameter IsNot Nothing Then
                    Dim myType As Type = objParameter.GetType()
                    strTitle = "given Parameter: " & myType.Name
                End If
            End If
            If strTitle = "" Then strTitle = "given Parameter: (unidentified)"

            If objParameter IsNot Nothing Then

                Select Case True
                    Case TypeOf (objParameter) Is Boolean
                        _SILSession.LogBool(detectLogLevel(logLevel), strTitle, objParameter)
                    Case TypeOf (objParameter) Is String
                        _SILSession.LogString(detectLogLevel(logLevel), strTitle, objParameter)
                    Case TypeOf (objParameter.ToString) Is String
                        _SILSession.LogString(detectLogLevel(logLevel), strTitle, objParameter.ToString)
                    Case TypeOf (objParameter) Is Integer
                        _SILSession.LogInt(detectLogLevel(logLevel), strTitle, objParameter)
                    Case TypeOf (objParameter) Is Date
                        _SILSession.LogDateTime(detectLogLevel(logLevel), strTitle, objParameter)

                        'Multiple Types!!!

                    Case Else
                        ' unknown (not programed) Type:
                        _SILSession.LogObject("detected Type", objParameter.GetType)
                        _SILSession.LogWarning("Can't identified parameter!")
                        _SILSession.LogObject(detectLogLevel(logLevel), strTitle, objParameter)
                End Select

            Else

                add(strTitle & " = (Nothing)", Me, detectLogLevel(logLevel))

            End If


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Sub logListOfObject(lstOfObjects As Object, Optional strTitle As String = "", Optional logLevel As Level? = Nothing)
        Dim bolEnteredMethod As Boolean = False
        Try
            If lstOfObjects Is Nothing Then Exit Sub
            If strTitle.Equals(String.Empty) Then
                strTitle = TryCast(lstOfObjects, String)
                If strTitle Is Nothing Then
                    Dim myType As Type = lstOfObjects.GetType()
                    strTitle = myType.Name
                End If
            End If
            If String.IsNullOrEmpty(strTitle) Then strTitle = "log List of objects"

            EnterMethod(strTitle)
            bolEnteredMethod = True
            For Each objItem As Object In lstOfObjects
                Dim myType As Type = objItem.GetType()
                Dim strName As String = TryCast(objItem.Name, String), strUID As String = Nothing
                Try
                    'is it a ALECS object with an UID?
                    strUID = " (" & TryCast(objItem.UID.ToString, String) & ")"
                Catch ex As Exception
                    strUID = String.Empty
                End Try
                If String.IsNullOrEmpty(strName) Then
                    _SILSession.LogObject(detectLogLevel(logLevel), myType.Name & strUID, objItem)
                Else
                    _SILSession.LogObject(detectLogLevel(logLevel), myType.Name & " " & strName.ToUpper & strUID, objItem)
                End If
            Next


        Catch ex As Exception
            myError.getException(ex, Me)
        Finally
            If bolEnteredMethod Then LeaveMethod(strTitle)
        End Try
    End Sub

    Public Function logDictionary(Of T)(dicObject As Dictionary(Of Guid, T), Optional bolFullObject As Boolean = False, Optional strName As String = "", Optional logLevel As Level? = Nothing) As Boolean
        logDictionary = False
        Try
            If dicObject Is Nothing Then add("The given dicObject for logDictionary is <NOTHING>", Me, DefaultLog.ActWarning)
            Dim lstObject As IList(Of T) = dicObject.Values.ToList
            '_SILSession.LogCollection("lstObject (Collection)", lstObject)
            '_SILSession.LogDataTable("lstObject (Table)", dicObject.Values)

            If Not lstObject.Any Then
                'nothing in it?
                Return False
            End If

            Dim strListItemName As String = dicObject.Values.FirstOrDefault.ToString
            strListItemName = strListItemName.Substring(strListItemName.IndexOf(".") + 1)

            '# Name:
            If String.IsNullOrEmpty(strName) Then
                strName = "Dictionary (Of " & strListItemName & ")"
            Else
                strName = "Dictionary: " & strName & " (Of " & strListItemName & ")"
            End If
            If String.IsNullOrEmpty(strName) Then strName = "logDictionary"


            'Serialize to JSON:
            Dim strJson As String = Newtonsoft.Json.JsonConvert.SerializeObject(lstObject)

            '# now convert to XML:
            Dim strDocument = String.Format("{{ " & strListItemName & ": {0} }}", strJson)
            Dim xmlDoc As System.Xml.XmlDocument = Newtonsoft.Json.JsonConvert.DeserializeXmlNode(strDocument, strListItemName & "_collection")

            '# read (from XML):
            Dim strReader As New IO.StringReader(xmlDoc.InnerXml), dsItem As DataSet = New DataSet()
            'myMain.Log.logXML("pretty XML string", strReader.ReadToEnd)
            dsItem.ReadXml(strReader)
            dsItem.Tables(0).TableName = strName

            '#logIT:
            If bolFullObject Then

                logJson("FULL-JSON-Output: " & strName, strJson, detectLogLevel(logLevel))

                'logXML("FULL-XML-Output: " & strName, xmlDoc.OuterXml, detectLogLevel(logLevel))
                'log all, even if DataSet contains multiple tables
                _SILSession.LogDataSet(detectLogLevel(logLevel), dsItem)

            Else

                _SILSession.LogDataTable(detectLogLevel(logLevel), dsItem.Tables(0))

            End If

            logDictionary = True


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Function

    ''' <summary>
    ''' logs as System.Collections.Specialized.NameValueCollection (NVC) as a Dictionary(Of String, String)
    ''' </summary>
    ''' <param name="nvcItem"></param>
    ''' <param name="strTitle"></param>
    ''' <param name="logLevel"></param>
    Public Sub logNameValueCollection(nvcItem As System.Collections.Specialized.NameValueCollection, Optional strTitle As String = Nothing, Optional logLevel As Level? = Nothing)
        Dim bolEnteredMethod As Boolean = False
        Try
            If nvcItem Is Nothing Then Exit Sub
            If String.IsNullOrEmpty(strTitle) Then
                strTitle = nvcItem.ToString
                If strTitle Is Nothing Then
                    Dim myType As Type = nvcItem.GetType()
                    strTitle = myType.Name
                End If
            End If
            If String.IsNullOrEmpty(strTitle) Then strTitle = "log NameValueCollection"

            EnterMethod(strTitle, detectLogLevel(logLevel))
            bolEnteredMethod = True
            Dim dicNVC As New Dictionary(Of String, String)
            For Each strKey As String In nvcItem.Keys
                dicNVC.Add(strKey, nvcItem(strKey))
            Next
            strTitle += ", counts: " & dicNVC.Count
            SIS.LogDictionary(detectLogLevel(logLevel), strTitle, dicNVC)


        Catch ex As Exception
            myError.getException(ex, Me)
            _SILSession.LogWarning("Can't log System.Collections.Specialized.NameValueCollection")
        Finally
            If bolEnteredMethod Then LeaveMethod(strTitle, detectLogLevel(logLevel))
        End Try
    End Sub

    Public Sub logTimeSpan(tspItem As TimeSpan, Optional strTitle As String = Nothing, Optional logLevel As Level? = Nothing)
        Try
            If Not logLevel.HasValue Then logLevel = SiAuto.Si.DefaultLevel

            'If kvpItem Is Nothing Then Exit Sub
            If String.IsNullOrEmpty(strTitle) Then
                'not so good: strTitle = kvpItem.ToString
                strTitle = "TimeSpan"
                If strTitle Is Nothing Then
                    Dim myType As Type = tspItem.GetType()
                    strTitle = myType.Name
                End If
            End If
            If String.IsNullOrEmpty(strTitle) Then strTitle = "log TimeSpan"
            '_SILSession.LogString("strTitle", strTitle)

            SIS.LogValue(detectLogLevel(logLevel), strTitle, tspItem)
            'SIS.LogLong(detectLogLevel(logLevel), strTitle, tspItem.Ticks)
            SIS.LogString(detectLogLevel(logLevel), strTitle, tspItem.Days & " days, " & tspItem.Hours & " hours, " & tspItem.Minutes & " minutes, " & tspItem.Seconds & " seconds.")


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Sub logKeyValuePair(kvpItem As KeyValuePair(Of String, String), Optional strTitle As String = Nothing, Optional logLevel As Level? = Nothing)
        Try
            'If kvpItem Is Nothing Then Exit Sub
            If String.IsNullOrEmpty(strTitle) Then
                'not so good: strTitle = kvpItem.ToString
                strTitle = "KeyValuePair"
                If strTitle Is Nothing Then
                    Dim myType As Type = kvpItem.GetType()
                    strTitle = myType.Name
                End If
            End If
            If String.IsNullOrEmpty(strTitle) Then strTitle = "log KeyValuePair"

            SIS.LogValue(detectLogLevel(logLevel), strTitle, kvpItem)


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Sub logSQL(sqlCommand As Data.Common.DbCommand, Optional strTitle As String = Nothing, Optional logLevel As Level? = Nothing)
        Try
            If sqlCommand Is Nothing Then Exit Sub
            If String.IsNullOrEmpty(strTitle) Then
                strTitle = sqlCommand.ToString
            End If
            If String.IsNullOrEmpty(strTitle) Then strTitle = "log DbCommand"

            SIS.LogSql(detectLogLevel(logLevel), strTitle, sqlCommand.CommandText)


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Function logConnectionString(objConnectionStringBuilder As Data.Common.DbConnectionStringBuilder, Optional strName As String = "", Optional logLevel As Level? = Nothing) As Boolean
        logConnectionString = False
        Try
            '# FAKE Password:
            Dim objFake As Object = Nothing

            Select Case True
                Case TypeOf (objConnectionStringBuilder) Is SqlConnectionStringBuilder
                    Dim sqlCSB As New SqlConnectionStringBuilder(objConnectionStringBuilder.ConnectionString)
                    sqlCSB.Password = "***** irrecognizable"
                    If String.IsNullOrEmpty(strName) Then
                        'definine name:
                        strName = sqlCSB.ApplicationName & " to """ & sqlCSB.InitialCatalog & """ DB on " & sqlCSB.DataSource
                        If Not sqlCSB.IntegratedSecurity Then strName += " with user '" & sqlCSB.UserID & "'"
                    End If
                    objFake = sqlCSB

                    'Case TypeOf (objConnectionStringBuilder) Is FirebirdSql.Data.FirebirdClient.FbConnectionStringBuilder
                    '    Dim sqlFB As New FirebirdSql.Data.FirebirdClient.FbConnectionStringBuilder(objConnectionStringBuilder.ConnectionString)
                    '    sqlFB.Password = "***** irrecognizable"
                    '    If String.IsNullOrEmpty(strName) Then
                    '        'definine name:
                    '        'strName = sqlFB.ApplicationName & " to """ & sqlFB.InitialCatalog & """ DB on " & sqlFB.DataSource
                    '        'If Not sqlFB.IntegratedSecurity Then strName += " with user '" & sqlFB.UserID & "'"
                    '    End If
                    '    objFake = sqlFB


                    'Todo: Implement mySQL/MariaDB ConnectionStringBuilder


                Case Else
                    '_SILSession.LogObject(logLevel, strName, objConnectionStringBuilder)
                    objFake = objConnectionStringBuilder

            End Select

            If String.IsNullOrEmpty(strName) Then
                Dim strType As String = objConnectionStringBuilder.GetType.Name  'TryCast(objConnectionStringBuilder, String)
                'If strType Is Nothing Then
                '    'Dim myType As Type = objConnectionStringBuilder.GetType()
                '    strType = objConnectionStringBuilder.GetType.Name
                'End If
                If Not String.IsNullOrEmpty(strType) Then
                    strName += " (Type=" & strType & ")"
                End If
            End If

            If strName = "" Then strName = "logConnectionString"

            '#logIt
            _SILSession.LogObject(detectLogLevel(logLevel), strName, objFake)

            logConnectionString = True


        Catch ex As Exception
            myError.getException(ex, Me)
            _SILSession.LogWarning("Can't log ConnectionString")
        End Try
    End Function

#End Region

    Public Sub OperationCanceledException(objMe As Object, Optional silLogging As Level = Level.Warning)
        Dim myME As New cms_ME(objMe)
        add(myME.Name & " operation was canceled", myME, silLogging, Color.LightGoldenrodYellow)
        logAssert("Operation is canceled. " & newMain.ProductName() & " is stopped!")
    End Sub

    Public Sub DebugMeDefaultReferences(ByRef objItem As Object, strPrefix As String)
        add("Sorry, DebugMeDefaultReferences is NOT implemented in this project!", Me, DefaultLog.ActWarning)
    End Sub

#Region "Enter/Leave Methods"

    Dim myTiming As Timing

    '# !!! NO multiple (verschaltelte) EnterMethod. cms_ME nutzt Ebenen (wenn wie bei LeaveMethod immer direkt programmieren, NIE eine die andere aufrufen!!! '#

    Public Sub EnterMethod(objMe As Object, Optional logLevel As Level? = Nothing)
        Try
            myTiming = New Timing
            Dim myME As New cms_ME(objMe)
            _SILSession.EnterMethod(detectLogLevel(logLevel), myME.myName)


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Sub LeaveMethod(objMe As Object, ByRef myTiming2 As Timing)
        Try
            'myTiming.Paused()
            Dim myME As New cms_ME(objMe)
            If myTiming2 IsNot Nothing Then
                _SILSession.LogString(Level.Verbose, "Runtime", FormatNumber(myTiming2.RuntimeMilliSeconds, 0, TriState.False, TriState.UseDefault, TriState.True) & " ms - " & vbNewLine & myTiming2.Duration(Timing.TimingFormat.Runtime))
            End If
            _SILSession.LeaveMethod(Level.Debug, myME.myName)


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    Public Sub LeaveMethod(ByRef objMe As Object, Optional logLevel As Level? = Nothing, Optional ByRef myTiming2 As Timing = Nothing)
        Try
            Dim myME As New cms_ME(objMe)
            If Not myTiming2 Is Nothing Then
                _SILSession.LogString(detectLogLevel(logLevel), "Runtime", FormatNumber(myTiming2.RuntimeMilliSeconds, 0, TriState.False, TriState.UseDefault, TriState.True) & " ms - " & vbNewLine & myTiming2.Duration(Timing.TimingFormat.Runtime))
            End If
            _SILSession.LeaveMethod(detectLogLevel(logLevel), myME.myName)


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

    'This is the good function and would be replaced by the nex one:
    Public Sub EndsSuccessfully(objMe As Object)
        Dim myME As New cms_ME(objMe)
        add(myME.Name & " ends successfully.", myME, detectLogLevel(Nothing))
    End Sub

#End Region

#Region "HELPER Classes: cms_ME Class, ReferenceDefault"

    Private Function detectLogLevel(Optional objLogLevel As Level? = Nothing) As Level
        detectLogLevel = SiAuto.Si.DefaultLevel
        Try
            If objLogLevel.HasValue Then
                detectLogLevel = objLogLevel
            End If


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Function

    Private Class cms_ME

#Region "VARs"

        'Public VARs for direct Access (Fields):
        Public Name As String = Nothing
        Public Type As System.Type = Nothing

        'Private VARs (Access only via Properties):
        Private _ME As Object = Nothing
        Private _Source As Reflection.MethodBase = Nothing, _SourceName As String = Nothing

#End Region

#Region "Constructors"

        Public Sub New(objME As Object)
            If TypeOf (objME) Is String Then
                _ME = New Object
                Name = objME
            Else
                _ME = objME
                Type = objME.GetType
                Name = Type.Name
            End If
            Dim myStackTrace As New System.Diagnostics.StackTrace()
            _Source = myStackTrace.GetFrame(2).GetMethod()
        End Sub

#End Region

#Region "Properties"

        Public ReadOnly Property myName As String
            Get
                If SourceName = "" Then
                    Return Name
                Else
                    Return Name & ":" & SourceName
                End If
            End Get
        End Property

        Public ReadOnly Property SourceName As String
            Get
                If _Source Is Nothing Then
                    Return _SourceName
                Else
                    Return _Source.Name
                End If
            End Get
        End Property

#End Region

    End Class

#End Region

#Region "SmartInspect Helper"

    <Obsolete("WARNING! Be careful with this! This is a very complex and expensive action, so please only use it economical", False)>
    Public Function setLogLevelPipe(Optional newLevel As Level? = Nothing) As Boolean
        setLogLevelPipe = False
        Dim oldLevel As Level = _SILSession.Level
        Try
            EnterMethod(Me, oldLevel)

            If Not newLevel.HasValue Then Return False
            If _SILSession.Level.Equals(newLevel) Then
                _SILSession.LogWarning("The newLevel EQUALS the old Level. So we do nothing!")
                Return True
            End If

            _SILSession.LogString(_SILSession.Level, "NOW we try to change the PIPE-logLevel from """ & oldLevel.ToString & """ TO", newLevel.ToString)

            'needed? myMain.Log.SIS.Level = newLevel
            SiAuto.Si.Level = newLevel
            _SILSession.Level = newLevel

            '### ==>> Activate File + Pipe Logging:
            ActivateConfig()

            _SILSession.LogString(newLevel, "LogLevel ist set to", newLevel.ToString)

            setLogLevelPipe = True


        Catch ex As Exception
            myError.getException(ex, Me)
        Finally
            LeaveMethod(Me, newLevel)
        End Try
    End Function

    Private Enum silConnectionType As Byte
        file = 1
        tcp = 2
        pipe = 3
    End Enum

    Private Sub ActivateConfig()
        Try
            '# SET --> CONSOLE Logging:
            Dim silConfigPipe As New ConnectionsBuilder
            With silConfigPipe
                .BeginProtocol("pipe")
                .AddOption("reconnect", True)
                .AddOption("reconnect.interval", "10s")
                .AddOption("level", SiAuto.Si.Level)
                .EndProtocol()
            End With

            ''# STOP ! (not needed!)
            'do NOT STOP this: _SILSession.Active = False
            'SiAuto.Si.Enabled = False
            'SiAuto.Si.SessionDefaults.Active = False

            SiAuto.Si.Connections = _CBFile.Connections & ", " & silConfigPipe.Connections
            'failed: myMain.Log.SIS.Connections = ...
            'failed _SILSession.Connections = ...

            ''# FIRE (start SmartInspect Logging by defaults)!
            _SILSession.Active = True       ' but start it!
            SiAuto.Si.Enabled = True
            SiAuto.Si.SessionDefaults.Active = True


        Catch ex As Exception
            myError.getException(ex, Me)
        End Try
    End Sub

#End Region

#Region "(Default) Level + Colors"

    Private Function getDefaultLevel(enmColor As DefaultLog) As Level
        Try
            Select Case enmColor
                Case DefaultLog.Default
                    Return Level.Debug
                Case DefaultLog.Successful
                    Return Level.Message
                Case DefaultLog.Warning
                    Return Level.Warning
                Case DefaultLog.Failed
                    Return Level.Message

                    'Act Colors dulled:
                Case DefaultLog.ActStart
                    Return Level.Verbose
                Case DefaultLog.ActEnd
                    Return Level.Verbose
                Case DefaultLog.ActSuccessful
                    Return Level.Message
                Case DefaultLog.ActWarning
                    Return Level.Warning
                Case DefaultLog.ActFailed
                    Return Level.Warning

                    'MileStone Colors very heavy:
                Case DefaultLog.MileStoneSuccessful
                    Return Level.Message
                Case DefaultLog.MileStoneFailed
                    Return Level.Error

                    'UNIT Color very light:
                Case DefaultLog.UnitStart
                    Return Level.Debug
                Case DefaultLog.UnitEnd
                    Return Level.Debug
                Case DefaultLog.UnitSuccessful
                    Return Level.Verbose
                Case DefaultLog.UnitFailed
                    Return Level.Warning

            End Select

            Return Level.Debug


        Catch ex As Exception
            myError.getException(ex, Me)
            Return Level.Fatal
        End Try
    End Function

    Private Function getDefaultColor(enmColor As DefaultLog) As Color
        Try
            Select Case enmColor
                Case DefaultLog.Default
                    Return Color.Empty
                Case DefaultLog.Successful
                    Return Color.DarkGreen
                Case DefaultLog.Warning
                    Return Color.Orange
                Case DefaultLog.Failed
                    Return Color.Red

                    'Act Colors dulled:
                Case DefaultLog.ActStart
                    Return Color.LightSkyBlue
                Case DefaultLog.ActEnd
                    Return Color.LightSkyBlue
                Case DefaultLog.ActSuccessful
                    Return Color.GreenYellow
                Case DefaultLog.ActWarning
                    Return Color.LightYellow
                Case DefaultLog.ActFailed
                    Return Color.LightPink

                    'MileStone Colors very heavy:
                Case DefaultLog.MileStoneSuccessful
                    Return Color.DarkOliveGreen
                Case DefaultLog.MileStoneFailed
                    Return Color.DarkRed

                    'UNIT Color very light:
                Case DefaultLog.UnitStart
                    Return Color.MediumPurple
                Case DefaultLog.UnitEnd
                    Return Color.MediumPurple
                Case DefaultLog.UnitSuccessful
                    'leicht
                    Return Color.LawnGreen
                Case DefaultLog.UnitFailed
                    'leicht
                    Return Color.IndianRed

            End Select
            Return Color.Empty


        Catch ex As Exception
            myError.getException(ex, Me)
            Return Color.Empty
        End Try
    End Function

#End Region

#End Region

End Class
