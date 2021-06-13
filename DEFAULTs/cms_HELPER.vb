
Imports System.Runtime.Serialization
Imports System.Xml.Linq
Imports Gurock.SmartInspect



<Serializable()> <DataContract()>
Public Class cms_Helper
    Inherits cms_LastError

#Region "VARs"

    Dim myMain As Main

    'Public VARs for direct Access (Fields):


    'Private VARs (Access only via Properties):

    Private _ActAsHost As ALECS_Host_Entry = Nothing

#End Region

#Region "Constructors"

    '# INIT by JOB Queue
    Public Sub New()
        myMain = New Main()
    End Sub

    '# INIT for addJob
    Public Sub New(ByRef objMain As Main)
        MyBase.New()
        myMain = objMain
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' here you can save the AHE item that's i'am working for. Different classes use this item for logging, events, etc.
    ''' be careful!!! If you do NOT save mySelf i'am acting as a diffrent host!
    ''' </summary>
    ''' <returns>ALECS_HOST_ENTRY. detectMySelf is nothing is specified. The SET AHE if it's set.</returns>
    Public Property Host As ALECS_Host_Entry
        Get
            'use saved value?
            If _ActAsHost IsNot Nothing Then Return _ActAsHost

            '#detect it:
            clearLastError()
            Try
                _ActAsHost = New ALECS_Host_Entry(myMain, GetType(ALECS_Host_Entry))
                If Not _ActAsHost.detectMySelf() Then Throw _ActAsHost.getLastError


            Catch ex As Exception
                setLastError(ex)
                _ActAsHost = Nothing
            End Try
            If _ActAsHost IsNot Nothing Then Return _ActAsHost Else Return Nothing
        End Get
        Set(value As ALECS_Host_Entry)
            _ActAsHost = value
        End Set
    End Property

#End Region

#Region "Functions: Values / Empty Values / CheckEmpty / etc..."

    Public Function CheckEmptyValues(Of T)(objValue As T, objReplacement As T) As T
        Try
            ''only for debugging purposes
            'If objValue Is Nothing Then
            '    myMain.Log.SIS.LogDebug("given objValue is NOTHING!")
            'Else
            '    myMain.Log.SIS.LogObject("given objValue to check if AnEmptyValue: '" & objValue.ToString & "'", objValue)
            'End If

            If AnEmptyValue(objValue) Then
                Return objReplacement
            Else
                Return objValue
            End If


        Catch ex As Exception
            Return objReplacement
        End Try
    End Function

    ''' <summary>
    ''' checks / validates the given object of be a guid (UnitTested)
    ''' </summary>
    ''' <param name="uidGUID"></param>
    ''' <returns>.HasValue shows if it's valid</returns>
    Public Function CheckEmptyGUID(uidGUID As Object) As Guid?
        Try
            ''only for debugging purposes
            'If uidGUID Is Nothing Then
            '    myMain.Log.SIS.LogDebug("given objValue is NOTHING!")
            'Else
            '    myMain.Log.SIS.LogObject("given objValue to check if AnEmptyValue: '" & uidGUID.ToString & "'", uidGUID)
            'End If

            If uidGUID Is Nothing Then Return Nothing

            If TypeOf uidGUID Is String Then
                Dim uidTemp As Guid = Nothing
                'if parse is OK then return uid else nothing
                If Guid.TryParse(uidGUID, uidTemp) Then uidGUID = uidTemp Else Return Nothing
            End If

            If uidGUID.Equals(Guid.Empty) Then
                Return Nothing
            ElseIf Not TypeOf uidGUID Is Guid Then
                Return Nothing
            Else
                Return uidGUID
            End If


        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Return an empty but NOT Nothing String or a wished replacement (default value)
    ''' </summary>
    ''' <typeparam name="T"></typeparam>         --> kann man übergeben, muss man aber nicht! I.d.R. bei Strings überfällig!
    Public Function CheckEmptyString(Of T)(strValue As T, Optional strReplacement As String = Nothing) As String
        Try
            ''only for debugging purposes
            'If strValue Is Nothing Then
            '    myMain.Log.SIS.LogDebug("given objValue is NOTHING!")
            'Else
            '    myMain.Log.SIS.LogObject("given objValue to check if AnEmptyValue: '" & strValue.ToString & "'", strValue)
            'End If

            If strValue Is Nothing OrElse IsDBNull(strValue) Then
                Return strReplacement
            ElseIf TypeOf strValue Is String AndAlso CObj(strValue) = vbNullChar Then
                Return strReplacement
            ElseIf strValue.Equals("") Then
                Return strReplacement
            Else
                Return strValue.ToString
            End If


        Catch ex As Exception
            Return strReplacement
        End Try
    End Function

    ''' <summary>
    ''' returned a EMPTY Value
    ''' </summary>
    ''' <typeparam name="T"></typeparam>         --> kann man übergeben, muss man aber nicht! I.d.R. bei Dates überfällig!
    Public Function CheckEmptyDate(Of T)(dteValue As T, Optional dteReplacement As DateTime = Nothing) As Date?
        Try
            ''only for debugging purposes
            'If dteValue Is Nothing Then
            '    myMain.Log.SIS.LogDebug("given objValue is NOTHING!")
            'Else
            '    myMain.Log.SIS.LogObject("given objValue to check if AnEmptyValue: '" & dteValue.ToString & "'", dteValue)
            'End If

            Dim bolReturnNothing As Boolean = False
            If dteReplacement = New DateTime(1, 1, 1) Then
                bolReturnNothing = True
            End If

            If dteValue Is Nothing OrElse IsDBNull(dteValue) Then
                If bolReturnNothing Then Return Nothing Else Return dteReplacement
                'Return dteReplacement
            ElseIf dteValue.ToString = "01.01.0001 00:00:00" Then
                If bolReturnNothing Then Return Nothing Else Return dteReplacement
            ElseIf dteValue.ToString = "01.01.0001" Then
                If bolReturnNothing Then Return Nothing Else Return dteReplacement
            Else
                Return Date.Parse(dteValue.ToString)
            End If


        Catch ex As Exception
            Return dteReplacement
        End Try
    End Function

    Public Function NotAnEmptyValue(objValue As Object) As Boolean
        NotAnEmptyValue = Not AnEmptyValue(objValue)
    End Function

    Public Function AnEmptyValue(objValue As Object) As Boolean
        AnEmptyValue = Nothing
        Try
            ''only for debugging purposes
            'If objValue Is Nothing Then
            '    myMain.Log.SIS.LogDebug("given objValue is NOTHING!")
            'Else
            '    myMain.Log.SIS.LogObject("given objValue to check if AnEmptyValue: '" & objValue.ToString & "'", objValue)
            'End If

            If objValue Is Nothing Then
                Return True
            ElseIf IsDBNull(objValue) Then
                Return True
            ElseIf IsNumeric(objValue) Then
                If objValue.Equals(Double.NaN) Then Return True
                If objValue.Equals(vbNull) Then Return True
                Return False
            ElseIf TypeOf objValue Is String AndAlso String.IsNullOrEmpty(objValue) Then
                Return True
            ElseIf TypeOf objValue Is Guid AndAlso objValue = Guid.Empty Then
                Return True
            ElseIf TypeOf objValue Is Guid AndAlso objValue = New Guid() Then
                'above is better??
                Return True
            ElseIf TypeOf objValue Is String AndAlso objValue.ToString.Contains("00000000-0000-0000-0000-000000000000") Then
                Return True
            ElseIf TypeOf objValue Is String AndAlso objValue = vbNullChar Then
                Return True
            ElseIf TypeOf objValue Is Date AndAlso objValue = New Date() Then
                Return True
            ElseIf TypeOf objValue Is TimeSpan AndAlso objValue = TimeSpan.FromMilliseconds(0) Then
                Return True
            Else
                Return False
            End If


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    ''' <summary>
    ''' Everytime you need an Loaded object you can use the shortener / equalizer
    ''' here's NO try because it throws direct errors!
    ''' </summary>
    ''' <param name="objALECs"></param>
    Public Sub CheckLoad(objALECs As ALECS_Template_Entry_FAKE)
        If objALECs.isLoaded Then

            'fine:
            Exit Sub

        Else

            If objALECs.Load Then

                'fine:
                Exit Sub

            Else

                Throw objALECs.getLastError

            End If

        End If
    End Sub

    ''' <summary>
    ''' Everytime you need an active object you can use the shortener / equalizer
    ''' here's NO try because it throws direct errors!
    ''' </summary>
    ''' <param name="objALECs"></param>
    Public Sub CheckActive(objALECs As ALECS_Template_Entry_FAKE)
        CheckLoad(objALECs)
        If objALECs.Active Then

            'fine:
            Exit Sub

        Else

            myMain.Log.SIS.LogObject(Level.Debug, "CheckActive fails on check: " & objALECs.getDisplayName, objALECs)
            Throw New ArgumentException("Sorry, but the object: " & objALECs.getDisplayName & " is NOT active! A DE-activated " & objALECs.ChildObjectName & " would NOT possible and ", "ACTIVE")

        End If
    End Sub

#End Region

#Region "Functions: Format / FormatDateTime / FormatSpace / etc..."

    Public Enum FormatDateTimeTypes
        [Default]
        Extended
        DateOnly
        DateOnlyShort
        TimeOnly
        TimeOnlyLong
        File
        FolderDay
    End Enum

    Public Function FormatDateTime(dteValue As DateTime, Optional FormatType As FormatDateTimeTypes = Nothing) As String
        FormatDateTime = String.Empty
        clearLastError()
        Try
            If AnEmptyValue(dteValue) Then Throw New ArgumentNullException("dteValue", "The given dteValue (DateValue) can NOT be NULL! Please specify a Date first.")

            Select Case FormatType
                Case FormatDateTimeTypes.File
                    'dteValue.ToString("yyyy-MM-dd_HH_mm_ss")
                    Return Format$(dteValue, "yyyyMMdd_HHmmss_fff")
                Case FormatDateTimeTypes.FolderDay
                    Return dteValue.ToString("yyyyMMdd")
                Case FormatDateTimeTypes.Extended
                    Return dteValue.ToString("dd.MM.yyyy HH:mm:ss.fff")
                Case FormatDateTimeTypes.DateOnly
                    Return dteValue.ToString("dd.MM.yyyy")
                Case FormatDateTimeTypes.DateOnlyShort
                    Return dteValue.ToString("dd.MM.yy")
                Case FormatDateTimeTypes.TimeOnly
                    Return dteValue.ToString("HH:mm:ss")
                Case FormatDateTimeTypes.TimeOnlyLong
                    Return dteValue.ToString("HH:mm:ss.fff")
                Case Else
                    Return dteValue.ToString("dd.MM.yyyy HH:mm:ss")
            End Select


        Catch ex As Exception
            setLastError(ex)
            Return ""
        End Try
    End Function

    Public Enum FormatNumberTypes
        [Default]
        LeadingDigit2
        LeadingDigit3
    End Enum

    Public Function FormatNumber(dlbValue As Double, Optional FormatType As FormatNumberTypes = Nothing) As String
        FormatNumber = String.Empty
        clearLastError()
        Try
            Select Case FormatType
                Case FormatNumberTypes.LeadingDigit2
                    Return Math.Floor(dlbValue).ToString("00")
                Case FormatNumberTypes.LeadingDigit3
                    Return Math.Floor(dlbValue).ToString("000")
                Case Else
                    Return Math.Floor(dlbValue).ToString("00000")
            End Select


        Catch ex As Exception
            setLastError(ex)
            Return ""
        End Try
    End Function

#Region "ByteSize Function"

    'ByteSize is a utility class that makes byte size representation in code easier by...
    'https://github.com/omar/ByteSize
    'in kleinen Teilen übernommen:

    Public Enum cms_Byte_Size As Long
        B = 8               'Byte
        KB = 1024           'Kilobyte
        MB = 1048576
        GB = 1073741824
        TB = 1099511627776
        PB = 1125899906842624
    End Enum

    Public Function ByteSize(strValueUnit As String, Optional DestinationSize As cms_Byte_Size = cms_Byte_Size.KB) As Double
        clearLastError()
        ByteSize = Double.NaN
        Try
            Dim dblValue As Double = Convert.ToDouble(System.Text.RegularExpressions.Regex.Replace(strValueUnit.ToUpper, "[^0-9]", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            Dim strUnit As String = System.Text.RegularExpressions.Regex.Replace(strValueUnit.ToUpper, "[^a-z]", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            Dim SourceSize As cms_Byte_Size = Nothing
            Select Case strUnit.ToUpper
                Case "B"
                    SourceSize = cms_Byte_Size.B
                Case "KB"
                    SourceSize = cms_Byte_Size.KB
                Case "MB"
                    SourceSize = cms_Byte_Size.MB
                Case "GB"
                    SourceSize = cms_Byte_Size.GB
                Case "TB"
                    SourceSize = cms_Byte_Size.TB
                Case "PB"
                    SourceSize = cms_Byte_Size.PB
                Case Else
                    Throw New Exception("Unit: " & strUnit & " is unknown. Please use: B, KB, MB, GB, TB, PB")
            End Select
            Dim dblBytes As Double = dblValue * SourceSize
            ByteSize = dblBytes / DestinationSize


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

#End Region

#End Region

#Region "Functions CONVERT..."

    Public Function convertObject(Of T)(objToConvert As Object) As T
        convertObject = Nothing
        Dim timTiming As New Timing
        clearLastError()
        Try
            convertObject = Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(Newtonsoft.Json.JsonConvert.SerializeObject(objToConvert))

            myMain.Log.SIS.LogObject("Converted object", convertObject)


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Function convertObject(objSource As Object, ByRef objDestination As Object, Optional bolFields As Boolean = True, Optional bolProperties As Boolean = True) As Boolean
        convertObject = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            '# Source Fields:
            If bolFields Then
                For Each fldItem As System.Reflection.FieldInfo In objSource.GetType.GetFields()
                    Dim objValue As Object = fldItem.GetValue(objSource)
                    If myMain.HELPER.NotAnEmptyValue(objValue) Then
                        Dim typDestination As Type = objDestination.[GetType]
                        Dim fldDestination As System.Reflection.FieldInfo = typDestination.GetField(fldItem.Name)
                        If Not fldDestination Is Nothing Then
                            fldDestination.SetValue(objDestination, objValue)
                        End If
                    End If
                Next
            End If

            '# Source Properties:
            If bolProperties Then
                For Each prpItem As System.Reflection.PropertyInfo In objSource.GetType().GetProperties()
                    If prpItem.CanWrite AndAlso prpItem.CanRead Then
                        Dim objValue As Object = prpItem.GetValue(objSource, Nothing)
                        If myMain.HELPER.NotAnEmptyValue(objValue) Then
                            Dim typDestination As Type = objDestination.[GetType]
                            Dim prpDestination As System.Reflection.PropertyInfo = typDestination.GetProperty(prpItem.Name)
                            If Not prpDestination Is Nothing Then
                                If prpDestination.CanWrite Then prpDestination.SetValue(objDestination, objValue)
                            End If
                        End If
                    End If
                Next
            End If

            convertObject = True


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    ''' <summary>
    ''' SAME as convertObject. So if you change here, change above
    ''' </summary>
    ''' <param name="objSource"></param>
    ''' <param name="objDestination"></param>
    Public Function convertObjectDirect(objSource As Object, ByRef objDestination As Object) As Boolean
        convertObjectDirect = False
        clearLastError()
        Try
            '# Source Fields:
            For Each fldItem As System.Reflection.FieldInfo In objSource.GetType.GetFields()
                Dim objValue As Object = fldItem.GetValue(objSource)
                If myMain.HELPER.NotAnEmptyValue(objValue) Then
                    Dim typDestination As Type = objDestination.[GetType]
                    Dim fldDestination As System.Reflection.FieldInfo = typDestination.GetField(fldItem.Name)
                    If Not fldDestination Is Nothing Then
                        fldDestination.SetValue(objDestination, objValue)
                    End If
                End If
            Next

            '# Source Properties:
            For Each prpItem As System.Reflection.PropertyInfo In objSource.GetType().GetProperties()
                If prpItem.CanWrite AndAlso prpItem.CanRead Then
                    Dim objValue As Object = prpItem.GetValue(objSource, Nothing)
                    If myMain.HELPER.NotAnEmptyValue(objValue) Then
                        Dim typDestination As Type = objDestination.[GetType]
                        Dim prpDestination As System.Reflection.PropertyInfo = typDestination.GetProperty(prpItem.Name)
                        If Not prpDestination Is Nothing Then
                            If prpDestination.CanWrite Then prpDestination.SetValue(objDestination, objValue)
                        End If
                    End If
                End If
            Next

            convertObjectDirect = True


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    ''' <summary>
    ''' http://bytutorial.com/blogs/aspnet/convert-enum-to-list-of-key-pairs-value-in-aspnet
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <returns>New List(Of KeyValuePair(Of String, Integer))</returns>
    ''' 
    Public Function ConvertEnum2List(Of T)() As List(Of KeyValuePair(Of String, Integer))
        ConvertEnum2List = New List(Of KeyValuePair(Of String, Integer))
        clearLastError()
        Try
            Dim lstKVP = New List(Of KeyValuePair(Of String, Integer))()
            For Each objItem As Object In System.[Enum].GetValues(GetType(T))
                lstKVP.Add(New KeyValuePair(Of String, Integer)(objItem.ToString(), CInt(objItem)))
            Next
            Return lstKVP


        Catch ex As Exception
            setLastError(ex)
            Return New List(Of KeyValuePair(Of String, Integer))
        End Try
    End Function

    ''' <summary>
    ''' convert ENUM values (both numbers and strings) to a valid ENUM value/type
    ''' </summary>
    ''' <typeparam name="T">NEVER a Nullable Type PLEASE!!!</typeparam>
    ''' <param name="strValue">the value as string or byte</param>
    ''' <param name="objReplacement">NOTHING is NOT given. otherwise the default value</param>
    ''' <returns>the ENUM as an Object instead of to be <Nothing>. T is alway enum=0</returns>
    Public Function [Enum](Of T)(strValue As String, Optional objReplacement As Object = Nothing) As Object '--> must be object instead of to be <Nothing>. T is alway enum=0
        [Enum] = objReplacement
        clearLastError()
        Try
            ''only for debugging purposes
            'If AppInDebugMode Then
            '    If String.IsNullOrEmpty(strValue) Then
            '        myMain.Log.SIS.LogDebug("given strValue is EMPTY!")
            '    Else
            '        myMain.Log.SIS.LogString("given strValue to convert to ENUM of Type: '" & GetType(T).ToString & "'. Converting.....", strValue)
            '    End If
            'End If

            '# check Nullables?
            Dim typType As Type = GetType(T), bolNullable As Boolean = False
            If Not typType.IsValueType Then bolNullable = True
            If Nullable.GetUnderlyingType(typType) IsNot Nothing Then bolNullable = True

            If String.IsNullOrEmpty(strValue) Then Exit Function

            'IMPORTANT Notice:
            '# did NOT work, cause: enmVariable can NOT be declared as T AndAlso NOT worked as object (then returns always the first value)
            'System.[Enum].TryParse(strValue, [Enum])
            'so we have to use PARSE (that's quicker) but FAILS :-(

            Try
                Dim enmTest As T = System.[Enum].Parse(GetType(T), strValue, True)
                If enmTest Is Nothing AndAlso objReplacement IsNot Nothing Then Return objReplacement
                [Enum] = enmTest


            Catch ex As Exception
                Exit Function
            End Try


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Function ConvertUmlauts(strContent As String) As String
        ConvertUmlauts = strContent
        clearLastError()
        Try
            'NULL / Nothing handling to prevent errors.
            If ConvertUmlauts Is Nothing Then Return Nothing

            ConvertUmlauts = ConvertUmlauts.Replace("ä", "ae")
            ConvertUmlauts = ConvertUmlauts.Replace("Ä", "Ae")
            ConvertUmlauts = ConvertUmlauts.Replace("ö", "oe")
            ConvertUmlauts = ConvertUmlauts.Replace("Ö", "Oe")
            ConvertUmlauts = ConvertUmlauts.Replace("ü", "ue")
            ConvertUmlauts = ConvertUmlauts.Replace("Ü", "Ue")
            ConvertUmlauts = ConvertUmlauts.Replace("ß", "ss")


        Catch ex As Exception
            setLastError(ex)
            Return Nothing
        End Try
    End Function

#End Region

#Region "Mail Stuff"

    Public Function checkEmailAddress(strEmailAddress As String) As Boolean
        checkEmailAddress = False
        clearLastError()
        Try
            myMain.Log.SIS.LogString("checking the given Emailaddress, if it's correct", strEmailAddress)

            If String.IsNullOrEmpty(strEmailAddress) Then
                checkEmailAddress = False
            Else
                'good pattern: https://regex101.com/r/KzgdK0/1
                Const pattern As String = "^((?!\.)([a-zA-Z0-9._\-öäüÖÄÜß](?<!\.\.))+(?<!\.)@[a-zA-Z0-9._\-öäüÖÄÜß]+\.[a-zA-Z0-9]+)$"

                '# options:
                Const options As System.Text.RegularExpressions.RegexOptions = System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.ExplicitCapture
                Dim rgxItem As System.Text.RegularExpressions.Regex = Nothing, matchTimeout As TimeSpan = TimeSpan.FromSeconds(2)
                rgxItem = New System.Text.RegularExpressions.Regex(pattern, options, matchTimeout)

                checkEmailAddress = rgxItem.IsMatch(strEmailAddress)

                strEmailAddress = strEmailAddress.ToLower()
                Dim emlAddress As System.Net.Mail.MailAddress = Nothing
                Try
                    emlAddress = New System.Net.Mail.MailAddress(strEmailAddress)


                Catch ex As Exception
                    checkEmailAddress = False
                End Try


                myMain.Log.SIS.LogBool(Level.Debug, "checkEmailAddress returns", checkEmailAddress)

            End If


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

#End Region

#Region "System things"

    Public Function AppPath(Optional withBackslash As Boolean = True) As String
        Dim tmpWert As String
        tmpWert = System.AppDomain.CurrentDomain.BaseDirectory
        If withBackslash Then
            Return returnPathWithBackslash(tmpWert)
        Else
            Return returnPathWithoutBackslash(tmpWert)
        End If
    End Function

    Public Function ThreadStart(ByRef thrThread As System.Threading.Thread, bolBackGroundThread As Boolean) As returnValue
        ThreadStart = New returnValue
        Dim ME_ As String = "cms_System:ProcessThreads"
        Try
            thrThread.IsBackground = bolBackGroundThread
            thrThread.Start()
            myMain.ThreadList.Add(thrThread)

            ThreadStart.Success = True


        Catch ex As Exception
            ThreadStart.Exception = ex
        End Try
    End Function

    Public Function Invoke(strClass As String, strMethod As String, Optional lstConstructor As List(Of Object) = Nothing, Optional lstParameter As List(Of Object) = Nothing) As Boolean
        Invoke = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            myMain.Log.EnterMethod(Me)

            Dim asmApplication As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly
            myMain.Log.SIS.LogString("asmApplication.GetName.Name", asmApplication.GetName.Name)

            Dim typClass As Type = asmApplication.GetType(asmApplication.GetName.Name & "." & strClass)

            If typClass IsNot Nothing Then

                myMain.Log.SIS.LogObject(Level.Verbose, "builded typClass: " & typClass.FullName, typClass)

                'Dim arrMethodInfoas As Reflection.MethodInfo() = modType.GetMethods()
                'For Each methodInfo As System.Reflection.MethodInfo In arrMethodInfoas
                '    myMain.Log.SIS.LogObject("methodInfo: " & methodInfo.Name, methodInfo)
                'Next

                'Dim miMethod As Reflection.MethodInfo = typClass.GetMethod(strMethod)
                ''miMethod = typClass.GetMethod(strMethod, Reflection.BindingFlags.Static Or Reflection.BindingFlags.Public)
                'myMain.Log.SIS.LogObject("working with miMethod: " & miMethod.Name, miMethod)

                ''Dim oType As Type = Type.GetTypeFromProgID(typClass.FullName)
                'Dim oType As Type = Type.GetTypeFromProgID("PDFConverter.PDFConverterX")

                '# build Constructor:
                Dim arrConstructor As Object() = Nothing, intCounter As Integer = 0
                If lstConstructor IsNot Nothing AndAlso lstConstructor.Count > 0 Then
                    arrConstructor = New Object(lstConstructor.Count - 1) {}
                    For Each objItem As Object In lstConstructor
                        arrConstructor(intCounter) = objItem
                        intCounter += 1
                    Next
                End If
                myMain.Log.SIS.LogArray(Level.Debug, "arrConstructor", arrConstructor)

                '# Create object:
                Dim objClass As Object = Activator.CreateInstance(typClass, arrConstructor)
                myMain.Log.SIS.LogObject("this is the object generated from " & strClass & " (objClass)", objClass)


                'Dim arrParameters As Object()
                'arrParameters = New Object(2) {}
                'arrParameters(0) = Nothing
                'arrParameters(1) = Nothing
                'arrParameters(2) = New ALECS_Constants.AHE(myMain).CCS_TSSHDEVMASTER

                '# build Parameter:
                Dim arrParameters As Object() = Nothing
                intCounter = 0
                If lstParameter IsNot Nothing AndAlso lstParameter.Count > 0 Then
                    arrParameters = New Object(lstParameter.Count - 1) {}
                    For Each objItem As Object In lstParameter
                        arrParameters(intCounter) = objItem
                        intCounter += 1
                    Next
                End If
                myMain.Log.SIS.LogArray(Level.Debug, "arrParameters", arrParameters)

                'good
                objClass.[GetType]().InvokeMember(strMethod, Reflection.BindingFlags.InvokeMethod, Nothing, objClass, arrParameters)
                'Conv.[GetType]().InvokeMember(strMethod, Reflection.BindingFlags.InvokeMethod, Nothing, Conv, arrParameters)

                'mi.Invoke(Nothing, New Object() {myMain})
                'mi.Invoke(Nothing, arrParameters)

            Else

                Throw New EntryPointNotFoundException("'" & asmApplication.GetName.Name & "." & strClass & "' was build, but NOT found in this app.")

            End If


            Invoke = True

            myMain.Log.EndsSuccessfully(Me)


        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
        Finally
            myMain.Log.LeaveMethod(Me, timTiming)
        End Try
    End Function

#End Region

#Region "XML Helper"

    Public Function xmlValue(xmlElement As XElement, strNode As String, Optional bolRequired As Boolean = False, Optional strDefault As String = Nothing) As String
        xmlValue = strDefault
        clearLastError()
        Try
            If xmlElement Is Nothing AndAlso bolRequired Then

                Throw New NullReferenceException("The requested xml element which should contain the node " & strNode & " does not exist in the configuration! Please create it!")

            ElseIf xmlElement Is Nothing Then

                Return strDefault

            Else
                Dim xmlSubElement As XElement = xmlElement.Element(strNode)

                If xmlSubElement Is Nothing AndAlso bolRequired Then

                    Throw New NullReferenceException("The node '" & strNode & "' does not exist in the configuration! Please create it!")

                ElseIf xmlSubElement Is Nothing Then

                    Return strDefault

                Else

                    xmlValue = xmlSubElement.Value

                End If

            End If


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Function xmlAttribute(xmlElement As XElement, strAttribute As String, Optional bolRequired As Boolean = False, Optional strDefault As String = Nothing) As String
        xmlAttribute = strDefault
        clearLastError()
        Try
            If xmlElement Is Nothing AndAlso bolRequired Then

                Throw New NullReferenceException("The requested xml element which should contain the attribute " & strAttribute & " does not exist in the configuration! Please create it!")

            ElseIf xmlElement Is Nothing Then

                Return strDefault

            Else

                Dim xmlAtt As XAttribute = xmlElement.Attribute(strAttribute)

                If xmlAtt Is Nothing AndAlso bolRequired Then

                    Throw New NullReferenceException("The requested xml Attribute " & strAttribute & " in xml element (" & xmlElement.Name.ToString & ") does not exist in the configuration! Please create the attribute!")

                ElseIf xmlAtt IsNot Nothing Then

                    xmlAttribute = xmlElement.Attribute(strAttribute).Value

                End If

            End If


        Catch ex As Exception
            If bolRequired Then
                setLastError(ex)
                Throw ex
            Else
                Return strDefault
            End If
        End Try
    End Function

    Public Function xmlAttributeAsBoolean(xmlElement As XElement, strAttribute As String, Optional bolDefault As Boolean = False) As Boolean
        xmlAttributeAsBoolean = bolDefault
        clearLastError()
        Try
            If xmlElement Is Nothing Then

                Return bolDefault

            Else

                Try
                    If xmlElement.Attribute(strAttribute) IsNot Nothing Then
                        Return xmlElement.Attribute(strAttribute).Value.ToLower().Equals("true")
                    Else
                        Return bolDefault
                    End If


                Catch ex As Exception
                    Return bolDefault
                End Try
            End If


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

#End Region

#Region "RANDOM"

    Public Enum CharTemplate
        [Default]
        Numeric
        LetterOnlyLowerCase
        LetterOnly
        File
        Password
        PasswordHard
    End Enum

    Public Function RandomString(Optional intStringLength As Integer = 32, Optional enmTemplate As CharTemplate = CharTemplate.Default) As String
        RandomString = String.Empty
        clearLastError()
        Try
            Dim strBuilder As New System.Text.StringBuilder()
            '# IMPORTANT Static usage: https://stackoverflow.com/questions/43652573/vb-net-random-number-generator-is-repeating-the-same-value
            Static rndRandom As New Random(), strDefault As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

            Dim strAvailable As String = String.Empty
            Select Case enmTemplate
                Case CharTemplate.Numeric
                    strAvailable = "0123456789"
                Case CharTemplate.LetterOnlyLowerCase
                    strAvailable = "abcdefghijklmnopqrstuvwxyz"
                Case CharTemplate.LetterOnly
                    strAvailable = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
                Case CharTemplate.File
                    strAvailable = strDefault
                Case CharTemplate.PasswordHard
                    strAvailable = strDefault & "<>|@€^°!""§$%&/()=?`*'_:;ß´+#äöüÖÜÄ-.,{[]}\~"
                Case Else
                    'Default (letter + some special chars):
                    strAvailable = strDefault & "*#+-_@$()=?ßäöüÖÜÄ[]"
            End Select

            For intIndex As Integer = 1 To intStringLength
                strBuilder.Append(strAvailable.Substring(rndRandom.Next(0, strAvailable.Length), 1))
            Next

            RandomString = strBuilder.ToString()


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

#End Region

#Region "Tools"

    Public Function TemporaryFile(Optional strFilenameWithOrWithoutPath As String = Nothing, Optional ByRef myCTN As System.Threading.CancellationToken = Nothing) As ALECS_File_Entry
        TemporaryFile = Nothing
        clearLastError()
        Try
            Dim fiTemp As IO.FileInfo = Nothing
            If String.IsNullOrEmpty(strFilenameWithOrWithoutPath) Then

                '#1 generate new FileName + path: be free
                fiTemp = New IO.FileInfo(IO.Path.GetTempFileName())

            Else

                '#1 check if strFilenameWithOrWithoutPath is with or without.....

                'Doku and example to how IO.Path.GetDirectoryName works: https://docs.microsoft.com/de-de/dotnet/api/system.io.path.getdirectoryname?view=netframework-4.8
                'also works with UNC Path's
                Dim strPath As String = IO.Path.GetDirectoryName(strFilenameWithOrWithoutPath)

                If Not String.IsNullOrEmpty(strPath) Then
                    '# with a given path
                    fiTemp = New IO.FileInfo(IO.Path.Combine(strPath, strFilenameWithOrWithoutPath))
                Else
                    '# withOUT a given path: but with a given Filename
                    fiTemp = New IO.FileInfo(IO.Path.Combine(IO.Path.GetTempPath(), strFilenameWithOrWithoutPath))
                End If
            End If

            TemporaryFile = New ALECS_File_Entry(myMain, fiTemp)


        Catch ex As Exception
            setLastError(ex)
            Return Nothing
        End Try
    End Function

#End Region

End Class
