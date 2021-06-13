
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Security.AccessControl
Imports Microsoft.Win32



<Serializable()> <DataContract()>
Public Class ALECS_FSO_Base
    Inherits ALECS_Template_Entry_FAKE
    Implements IALECS_Entry

#Region "VARs"

    'Public VARs for direct Access (Fields):

    <DataMember(EmitDefaultValue:=False, Order:=11)>
    Public Storage As ALECS_FSO_Base.ALECS_FSO_Storage? = Nothing
    <DataMember(EmitDefaultValue:=False, Order:=12)>
    Public Host As ALECS_Host_Entry = Nothing


    <Serializable()> <DataContract()>
    Public Enum ALECS_FSO_Types
        File
        Directory
    End Enum

    <Serializable()> <DataContract()>
    Public Enum ALECS_FSO_Storage
        Drive
        UNC
        HTTP
        HTTPS
        ' and many more, later
    End Enum

    <Serializable()> <DataContract()>
    Public Enum ALECS_FSO_SystemInfo
        CreationTime
        LastAccess
        LastWrite
    End Enum

#End Region

#Region "Constructors"

    'private, because cms_FSO is only for internal use
    Public Sub New(objType As Type)
        MyBase.New(objType)
    End Sub

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New(ByRef objMain As Main, objType As Type)
        MyBase.New(objMain, objType)
    End Sub

    'private, because cms_FSO is only for internal use
    Public Sub New(ByRef objMain As Main, objType As Type, Optional objAHE As ALECS_Host_Entry = Nothing)
        MyBase.New(objMain, objType)
        Host = objAHE
    End Sub

#End Region

#Region "Properties"

    <DataMember(EmitDefaultValue:=False, Order:=13)>
    Public ReadOnly Property StorageAsString As String
        Get
            If Not Storage.HasValue Then Return Nothing
            Return Storage.ToString
        End Get
    End Property

#End Region

#Region "Functions"

    'NO Unit Test!
    Public Function isFileOrDirectory(strFullname As String) As ALECS_FSO_Types?
        isFileOrDirectory = Nothing
        clearLastError()
        Try
            If (IO.File.GetAttributes(strFullname) And IO.FileAttributes.Directory) = FileAttributes.Directory Then
                Return ALECS_FSO_Types.Directory
            Else
                Return ALECS_FSO_Types.File
            End If


        Catch ex As Exception
            'NO error handling here. If it NOT exist we do NOT want an error in the log!
            'myMain.Error.handleException(ex, Me)
            setLastError(ex)
        End Try
    End Function

    Public Sub setStorage(strFullname As String)
        Try
            strFullname = strFullname.ToLower
            Select Case True
                Case strFullname.StartsWith("\\")
                    Storage = ALECS_FSO_Storage.UNC

                Case strFullname.Substring(1, 2) = ":\"
                    Storage = ALECS_FSO_Storage.Drive

                Case strFullname.StartsWith("https://")
                    Storage = ALECS_FSO_Storage.HTTPS

                Case strFullname.StartsWith("http://")
                    Storage = ALECS_FSO_Storage.HTTP

                Case Else
                    Storage = Nothing
            End Select


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' returns the "HOSTNAME" from a UNC Path
    ''' </summary>
    ''' <param name="strFullname"></param>
    Public Function getHost(strFullname As String) As String
        getHost = Nothing
        Dim timTiming As New Timing
        clearLastError()
        Try
            If Storage <> ALECS_FSO_Storage.UNC Then
                Throw New NotImplementedException("NOT worked yet for other Storages than: UNC")
            End If

            getHost = strFullname.Replace("\\", "")
            getHost = getHost.Split("\")(0)

            myMain.Log.SIS.LogObject("getHost results", getHost)


        Catch ex As Exception
            setLastError(ex)
            Return Nothing
        End Try
    End Function

#End Region

End Class





<Serializable()> <DataContract()>
Public Class ALECS_File_Entry
    Inherits ALECS_FSO_Base
    Implements IALECS_Entry, IALECS_Entry.IDeletable, IALECS_Entry.ICreateable

#Region "VARs"

    'specified in cms_FSO DIM myMain As Main

    'Public VARs for direct Access (Fields):
    <DataMember(EmitDefaultValue:=False, Order:=15)>
    <System.Xml.Serialization.XmlIgnoreAttribute>
    Public FileUri As Uri = Nothing


    'Private VARs (Access only via Properties):
    Private _File As IO.FileInfo = Nothing
    '# Helper VARs for Caching:
    Private _ExistsCache As New NET_Cache
    Private _AccessCache As New NET_Cache

    '# only for Files: containing logical objects like ALECS_Config, cml_Config, etc.
    Private _Value As Object = Nothing

#End Region

#Region "Constructors"

    '# for INIT by ALECS_JobS
    Public Sub New()
        MyBase.New(GetType(ALECS_File_Entry))
    End Sub

    Public Sub New(ByRef objMain As Main, filSource As IO.FileInfo, Optional objAHE As ALECS_Host_Entry = Nothing)
        MyBase.New(objMain, GetType(ALECS_File_Entry), objAHE)
        FSO = filSource
        Host = objAHE
    End Sub

#End Region

#Region "Properties"

    'not needed <DataMember(EmitDefaultValue:=False)>
    <System.Xml.Serialization.XmlIgnoreAttribute>
    Public Property FSO As IO.FileInfo
        Get
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    Exit Select
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
            Return _File
        End Get
        Set(value As IO.FileInfo)
            _File = value
            If Not _File Is Nothing Then
                setStorage(value.FullName)
            Else
                'Fake: Drive settings (for now)
                Storage = ALECS_FSO_Storage.Drive
            End If
        End Set
    End Property

    'not needed: <DataMember(EmitDefaultValue:=False, Order:=11)>
    Public Overridable Shadows ReadOnly Property Name As String 'Implements IALECS_Entry.Name
        Get
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    If _File Is Nothing Then Return Nothing
                    Return _File.Name
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Return String.Empty

                Case Else
                    Return Nothing
            End Select
        End Get
    End Property

    '# IMPORTANT: Not Readonly, because csm_FSO.File can not READ from JSON
    <DataMember(EmitDefaultValue:=False, Order:=21)>
    Public Property Fullname As String
        Get
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    If _File Is Nothing Then Return Nothing
                    Return _File.FullName
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    If FileUri Is Nothing Then Return Nothing
                    Return FileUri.AbsoluteUri
                Case Else
                    Return Nothing
            End Select
        End Get
        Set(value As String)
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    _File = New FileInfo(value)
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    FileUri = New Uri(value)
                Case Else
                    Throw New NotImplementedException("Fullname property (SET) not implemented for " & Storage.ToString & " based files")
            End Select
        End Set
    End Property

    'not needed: <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Property Value As Object
        Get
            Return _Value
        End Get
        Set(value As Object)
            _Value = value
        End Set
    End Property

    'not needed: <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public ReadOnly Property Version As FileVersionInfo
        Get
            If _File Is Nothing OrElse Not CheckIfExist() Then Return Nothing

            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    Return FileVersionInfo.GetVersionInfo(_File.FullName)
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property

    'do not export, only for IDE use: <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public ReadOnly Property Path As ALECS_Directory_Entry
        Get
            Select Case Storage
                Case ALECS_FSO_Storage.UNC, ALECS_FSO_Storage.Drive
                    If _File Is Nothing Then Return Nothing
                    Return New ALECS_Directory_Entry(myMain, New DirectoryInfo(_File.DirectoryName), Host)
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property

    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Overloads ReadOnly Property Created As DateTime? Implements IALECS_Entry.Created
        Get
            Return FSO.CreationTime.ToLocalTime
        End Get
    End Property

    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Overloads ReadOnly Property changed As DateTime? Implements IALECS_Entry.changed
        Get
            Return FSO.LastWriteTime.ToLocalTime
        End Get
    End Property

#End Region

#Region "Functions"

#Region "IALECS_Entry Functions"

    Public Overrides Function CheckIfExist(Optional bolUseCache As Boolean = True) As Boolean 'Implements IALECS_Entry.ILoadable.CheckIfExist
        CheckIfExist = False
        clearLastError()
        Try
            If Not _ExistsCache.Filled Or Not bolUseCache Then
                '# fill Cache:
                Try
                    Select Case Storage
                        Case ALECS_FSO_Storage.Drive
                            If Not checkAccess(FileAccess.Read, bolUseCache) Then
                                Try
                                    Throw _AccessCache.Exception
                                Catch ex As IOException
                                    'The File exists and i would have access if no one else has open it. But this happends now :-(
                                Catch ex As Exception
                                    Throw ex
                                End Try
                            End If
                            FSO.Refresh()
                            _ExistsCache.Value = FSO.Exists

                        Case ALECS_FSO_Storage.UNC
                            '# Try 1:
                            If Not checkAccess(FileAccess.Read, bolUseCache) Then
                                'myMain.Log.add("checkAccess can't read on it :-( so we must handle this error: " & _AccessCache.Exception.Message, Me)
                                Try
                                    'throw _AccessCache Exception for Exist (for log)
                                    Throw _AccessCache.Exception


                                Catch ex As FileNotFoundException
                                    'ups, file exists NOT!
                                    Throw ex
                                Catch ex As UnauthorizedAccessException
                                    'so this is a problem... we did NOT know if File exist and we only have NOT access :-(
                                    Throw ex
                                Catch ex As InvalidOperationException
                                    'Unerwarteter Methodenfehler xx. (handled in checkAccess)
                                    Throw ex

                                    'Catch ex As IOException
                                    '    myMain.DP(8110)
                                    '    myMain.Log.add(ex.Message, Me, DefaultLog.ActWarning)

                                Catch ex As Exception
                                    Throw New NotImplementedException(ex.Message)
                                End Try
                            End If
                            FSO.Refresh()
                            _ExistsCache.Value = FSO.Exists

                        Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS


                        Case Else
                            Throw New NotImplementedException("Exists check NOT implemented for " & Storage.ToString & " Type files")
                    End Select

                    '# so here we're: and the File did Exists :-)
                    _ExistsCache.Filled = True


                Catch ex As Exception
                    '# so here we're: and the File did NOT Exist :-(
                    _ExistsCache.Value = False
                    _ExistsCache.Filled = True
                    _ExistsCache.Exception = ex
                    setLastError(ex)
                End Try

            Else
                '# use Cache:

                'last line returned: Exists = _ExistCache.bolValue
                If _ExistsCache.Exception IsNot Nothing Then setLastError(_ExistsCache.Exception)

            End If

            myMain.Log.SIS.LogBool("CheckIfExist", _ExistsCache.Value)


            CheckIfExist = _ExistsCache.Value


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
            CheckIfExist = False
        End Try
    End Function

    Public Overrides Function Delete() As Boolean Implements IALECS_Entry.IDeletable.Delete
        Delete = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            'protect before UnitTest Errors / Fails and accidental delete:
            If Fullname.Equals("\\alecs_config.ccs.cmsystemhaus.net\alecs_storage\TEST\UnitTests\IamAUnitTestFile.test", StringComparison.CurrentCultureIgnoreCase) Then Throw New InvalidProgramException("OH NO, do NOT delete: \\alecs_config.ccs.cmsystemhaus.net\alecs_storage\TEST\UnitTests\ PLEASE :-o")

            Select Case Storage
                Case ALECS_FSO_Storage.Drive
                    '# Check if Operation Cancelled:
                    myMain.CTN.ThrowIfCancellationRequested()

                    FSO.Delete()

                Case ALECS_FSO_Storage.UNC
                    If Not checkAccess(FileAccess.Write, False, False) Then
                        If TypeOf getLastError() Is FileNotFoundException Then
                            Return True
                        End If
                        Throw getLastError()
                    End If

                    '# Check if Operation Cancelled:
                    myMain.CTN.ThrowIfCancellationRequested()

                    FSO.Delete()

                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS


                Case Else
                    Throw New NotImplementedException("Delete function not implemented for " & Storage.ToString & " Type files.")

            End Select

            Dispose()

            Delete = True


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Overrides Function Create() As Boolean Implements IALECS_Entry.ICreateable.Create
        Create = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    Dim ioFile As New IO.FileInfo(Fullname)
                    ioFile.Create().Close()
                    Create = True


                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Throw New NotImplementedException("ALECS_File_Entry.Create --> StorageType=" & Storage.ToString)

                Case Else
                    Throw New NotImplementedException("ALECS_File_Entry.Create --> StorageType=" & Storage.ToString)
            End Select

            'reset Cache:
            _ExistsCache = New NET_Cache


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Overrides Function getDisplayName() As String Implements IALECS_Entry.getDisplayName
        getDisplayName = String.Empty
        clearLastError()
        Try
            If _File IsNot Nothing Then
                getDisplayName = _File.Name & " in (" & Fullname & ")"
            End If
            getDisplayName += ", Type=File"
            If Storage.HasValue Then
                If Not String.IsNullOrEmpty(getDisplayName) Then getDisplayName += ", "
                getDisplayName += "Storage=" & Storage.ToString
            End If


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Overrides Sub DebugMe(Optional strPREfix As String = "") Implements IALECS_Entry.DebugMe
        Try
            myMain.Log.SIS.LogColored(Color.LightGray, "DEBUGME: " & getDisplayName())
            myMain.Log.SIS.LogObject(strPREfix & ChildObjectName & ": " & getDisplayName(), Me)

#Region "Details"
            Dim strDetails As String = "DETAILS"
            myMain.Log.SIS.EnterMethod(strDetails)
            If String.IsNullOrEmpty(strPREfix) Then strPREfix = "FIL "
            'Host:
            If Host IsNot Nothing Then
                myMain.Log.SIS.LogObject("...HOST", Host)
                'Host.DebugMe(strPREfix)
            Else
                'myMain.Log.add("HOST is Nothing", Me)
            End If
            myMain.Log.SIS.LeaveMethod(strDetails)
#End Region


        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Basic FSO Functions"

    <Obsolete("This function is depricated. Please use CheckIfExist function instead!", True)>
    Public Function Exists(Optional useCache As Boolean = True) As Boolean
        Return CheckIfExist(useCache)
    End Function

#Region "MOVE Functions"

    Public Function Move(ByRef filDestination As ALECS_File_Entry, Optional bolOverwrite As Boolean = True) As Boolean
        Move = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            Select Case Storage
                        'NO global Move, put back to fil+dir functions. Kein Vorteil von global.

                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    If Fullname.Equals(filDestination.Fullname, StringComparison.CurrentCultureIgnoreCase) Then Throw New InvalidOperationException("Source and Destination File are equal!")

                    'Check Access Source:
                    If Not checkAccess(FileAccess.Read, False, True) Then Throw _AccessCache.Exception

                    '# in case of problem with UNC access add the following lines of code to check / make access before!
                    Dim dirDestination As ALECS_Directory_Entry = filDestination.getAsDirectory
                    If Not dirDestination.CheckIfExist(True) Then If Not dirDestination.Create() Then Throw dirDestination.getLastError

                    '# Check if Operation Cancelled:
                    myMain.CTN.ThrowIfCancellationRequested()

                    If filDestination.CheckIfExist(True) AndAlso bolOverwrite Then
                        'Copy first
                        Dim fsoResult As IO.FileInfo = FSO.CopyTo(filDestination.Fullname, True)
                        myMain.Log.SIS.LogObject("fsoResult (positive?!)", fsoResult)
                        If fsoResult.FullName.Equals(filDestination.Fullname, StringComparison.CurrentCultureIgnoreCase) Then
                            Move = True
                        Else
                            Throw New IO.FileNotFoundException("The following FileName was returned by Move (clandestine copy) action but has to be: " & filDestination.Fullname, fsoResult.FullName)
                        End If
                        'then delete:
                        If Not Delete() Then Throw getLastError()
                    Else
                        'this command throws an error if destination already exists (what is OK when bolOverwrite is false)
                        FSO.MoveTo(filDestination.Fullname)
                        '2DO (Check if OK)
                        Move = True
                    End If

                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Throw New NotImplementedException("MOVE function from HTTP(S) is NOT unitTested @ the moment!")

                Case Else
                    Throw New NotImplementedException("MOVE function not implemented for " & Storage.ToString & " type Files")

            End Select
            _AccessCache = New NET_Cache
            _ExistsCache = New NET_Cache
            filDestination._AccessCache = New NET_Cache
            filDestination._ExistsCache = New NET_Cache


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
            Return False
        End Try
    End Function

#End Region

#Region "COPY Functions"

    Public Function Copy(ByRef filDestination As ALECS_File_Entry, Optional bolOverwrite As Boolean = True, Optional bolAsUploadtoSource As Boolean = False) As Boolean
        Copy = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            Select Case Storage                            '==> Source Storage

                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC

                    Select Case filDestination.Storage      '==> Destination Storage
                        Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                            If Fullname.Equals(filDestination.Fullname, StringComparison.CurrentCultureIgnoreCase) Then Throw New InvalidOperationException("Source and Destination File are equal!")

                            'Check Access Source:
                            If Not checkAccess(FileAccess.Read, False, True) Then Throw _AccessCache.Exception

                            '# in case of problem with UNC access add the following lines of code to check / make access before!
                            Dim dirDestination As ALECS_Directory_Entry = filDestination.getAsDirectory
                            If Not dirDestination.CheckIfExist() Then If Not dirDestination.Create() Then Throw dirDestination.getLastError
                            'Check Access Destination:
                            'did NOT know why, but it fails: If Not dirDestination.checkAccess(FileAccess.Write, False) Then Throw dirDestination.getLastError

                            '# Check if Operation Cancelled:
                            myMain.CTN.ThrowIfCancellationRequested()

                            ''myMain.Log.add("start copying File " & _File.FullName & " to " & filDestination.Fullname & ".....", Me)
                            '# Copy File:
                            If Name.IndexOf("*") > -1 Then
                                'Checkme: Missing error message. is this for multiple files?
                                Throw New NotImplementedException("Sorry copying files with a star (*) in NOT supported yet! Please contact development!")
                            Else
                                '# Check if Operation Cancelled:
                                myMain.CTN.ThrowIfCancellationRequested()

                                'Copy without Wildcard:
                                My.Computer.FileSystem.CopyFile(Fullname, filDestination.Fullname, bolOverwrite)
                            End If
                            Copy = True

                        Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                            Throw New NotImplementedException("File COPY function (File action) is NOT implemented for actions from Drive or UNC to " & filDestination.Storage.ToString)

                        Case Else
                            Throw New NotImplementedException("File COPY function (File action) is NOT implemented for actions from Drive or UNC to " & filDestination.Storage.ToString)

                    End Select

                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Select Case filDestination.Storage      '==> Destination Storage
                        Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                            'Copy = copyFileFromURItoDriveUNC(Fullname, filDestination, bolOverwrite, bolAsUploadtoSource)
                        Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                            Throw New NotImplementedException("NOT unittested yet!")
                        Case Else
                            Throw New NotImplementedException("Copy action from Http(s) to " & filDestination.Storage.ToString & "is NOT implemented yet")
                    End Select

                Case Else
                    Throw New NotImplementedException("COPY Function not implemented for " & Storage.ToString & " Type Files")
            End Select

            If Not Copy Then Throw getLastError()

            '_ExistsCache = New ALECS_Cache
            filDestination._AccessCache = New NET_Cache
            filDestination._ExistsCache = New NET_Cache


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
            Copy = False
        End Try
    End Function

    Public Function Copy(ByRef dirDestination As ALECS_Directory_Entry, Optional bolOverwrite As Boolean = True, Optional bolAsUploadtoSource As Boolean = False) As Boolean
        Copy = False
        clearLastError()
        Try
            Dim filDestination As ALECS_File_Entry = Nothing

            Select Case Storage                            '==> Source Storage

                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC

                    Select Case dirDestination.Storage      '==> Destination Storage
                        Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                            'specify destination filename by source + dirDestination:
                            filDestination = New ALECS_File_Entry(myMain, New IO.FileInfo(dirDestination.Fullname & Name), dirDestination.Host)

                        Case Else
                            Throw New NotImplementedException("The ALECS_File_Entry.COPY (2) function (File action) is NOT implemented for actions from Drive/UNC to " & dirDestination.Storage.ToString)

                    End Select

                Case Else
                    Throw New NotImplementedException("COPY Function not implemented for storage" & Storage.ToString & " Type Files")

            End Select

            If Not Copy(filDestination, bolOverwrite, bolAsUploadtoSource) Then
                Throw getLastError()
            Else
                Copy = True
            End If


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

#End Region

#End Region

#Region "checkFunctions"

    Public Function checkAccess(Optional AccessType As FileAccess = FileAccess.Read, Optional useCache As Boolean = True, Optional bolMustExist As Boolean = False) As Boolean
        checkAccess = False
        clearLastError()
        Try
            If Not _AccessCache.Filled Or Not useCache Then
                '# fill Cache:
                Try
                    Dim intRetry As Integer = 2
                    Select Case Storage

                        Case ALECS_FSO_Storage.Drive
                            Try
                                '# Solution from here (see below): https://stackoverflow.com/questions/130617/how-do-you-check-for-permissions-to-write-to-a-directory-or-file
                                Dim arcRule As System.Security.AccessControl.AuthorizationRuleCollection = IO.File.GetAccessControl(_File.FullName).GetAccessRules(True, True, GetType(System.Security.Principal.NTAccount))
                                For Each fsaRule As System.Security.AccessControl.FileSystemAccessRule In arcRule
                                    If fsaRule.AccessControlType = System.Security.AccessControl.AccessControlType.Allow Then
                                        myMain.Log.SIS.LogBool("hasAccess", True)
                                        Exit For
                                    End If
                                Next


                            Catch ex As FileNotFoundException
                                If bolMustExist Then
                                    myMain.Log.SIS.LogWarning("The file '" & _File.FullName & "' does NOT exist (bolMustExist=" & bolMustExist & "): " & vbNewLine & ex.Message)
                                    Throw ex
                                Else
                                    'myMain.Log.add("That's NOT bad :-). File must NOT exists so i returned: True", Me, Level.Verbose, Color.LightYellow)
                                    Return True
                                End If
                            End Try

                        Case ALECS_FSO_Storage.UNC
                            Try
startAccessUNC:
                                '# Solution from here (see below): https://stackoverflow.com/questions/130617/how-do-you-check-for-permissions-to-write-to-a-directory-or-file
                                Dim secFile As System.Security.AccessControl.FileSecurity = Nothing, intCounterRetrySecFile As Integer = 5
retrySecFile:
                                Try
                                    secFile = IO.File.GetAccessControl(_File.FullName)


                                Catch ex As FileNotFoundException
                                    Throw ex

                                Catch ex As Exception
                                    myMain.Log.SIS.AddCheckpoint("That's really BAD! We try a secFile INIT with NO success! Now we try arcRule, but this seems NOT to be better :-(")
                                End Try


                                Dim arcRule As System.Security.AccessControl.AuthorizationRuleCollection = secFile.GetAccessRules(True, True, GetType(System.Security.Principal.NTAccount))
                                For Each fsaRule As System.Security.AccessControl.FileSystemAccessRule In arcRule
                                    If fsaRule.AccessControlType = System.Security.AccessControl.AccessControlType.Allow Then
                                        myMain.Log.SIS.LogBool("checkAccess (" & _File.FullName & ")", True)
                                        Exit For
                                    End If
                                Next


                            Catch ex As FileNotFoundException
                                'handled in Catch block before! Only for jump over arcRule
                                Throw ex
                            Catch ex As UnauthorizedAccessException
                                'handled in Catch block before! Only for jump over arcRule
                                Throw ex
                            Catch ex As InvalidOperationException
                                'handled in Catch block before! Only for jump over arcRule
                                Throw ex
                            Catch ex As NullReferenceException
                                intRetry -= 1
                                If intRetry <= 0 Then
                                    Throw New NullReferenceException("The maximum retries of accessing '" & _File.FullName & "' exeeded!")
                                Else
                                    GoTo startAccessUNC
                                End If
                            Catch ex As Exception
                                myMain.Log.SIS.LogFatal("So whats NOW? We got a: " & ex.GetType.ToString & " exception and what next? (see following details)")
                                Throw ex
                            End Try

                        Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS


                        Case Else
                            Throw New NotImplementedException("checkAccess function not implemented for " & Storage & " type files")

                    End Select

                    '# so here we're: and we HAVE Access :-)
                    _AccessCache.Value = True
                    _AccessCache.Filled = True


                Catch ex As Exception
                    '# so here we're: and we have NO Access :-(
                    _AccessCache.Value = False
                    _AccessCache.Filled = True
                    _AccessCache.Exception = ex
                    setLastError(ex)
                    'really needed an eror? Throw
                End Try

            Else
                '# use Cache:

                'last line returned checkAccess = _AccessCache.bolValue
                If _AccessCache.Exception IsNot Nothing Then setLastError(_AccessCache.Exception)

            End If

            checkAccess = _AccessCache.Value

            myMain.Log.SIS.LogBool("checkAccess (" & _File.FullName & ")", checkAccess)


        Catch ex As Exception
            setLastError(ex)
            Return False
        End Try
    End Function

#End Region

#Region "OTHER"

    ''' <summary>
    ''' returns the FILENAME from a File
    ''' </summary>
    Public Function getFilenameWithoutExtension() As String
        getFilenameWithoutExtension = String.Empty
        clearLastError()
        Try
            Select Case Storage

                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    Return IO.Path.GetFileNameWithoutExtension(_File.FullName)

                    'Case cms_FSO_Storage.HTTP, cms_FSO_Storage.HTTPS
                    '    Return FileUri.AbsoluteUri

                Case Else
                    Throw New NotImplementedException("ALECS_File_Entry.getFilenameWithoutExtension --> Storage=" & Storage.ToString)

            End Select


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    ''' <summary>
    ''' returns the ROOT Directory (Share?) from a UNC File
    ''' </summary>
    ''' <returns></returns>
    Public Function getRootDirectory() As ALECS_Directory_Entry
        getRootDirectory = Nothing
        Dim timTiming As New Timing
        clearLastError()
        Try
            getRootDirectory = Path.Clone()
            getRootDirectory.FSO = New IO.DirectoryInfo(Path.Fullname).Root

            myMain.Log.SIS.LogObject("getRootDirectory results: " & getRootDirectory.getDisplayName, getRootDirectory)


        Catch ex As Exception
            setLastError(ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' returns the DIRECTORY from a Unc File
    ''' </summary>
    ''' <returns></returns>
    Public Function getAsDirectory() As ALECS_Directory_Entry
        getAsDirectory = Nothing
        Dim timTiming As New Timing
        clearLastError()
        Try
            getAsDirectory = Path.Clone()
            getAsDirectory.FSO = New IO.DirectoryInfo(Path.Fullname)
            getAsDirectory.Host = Host
            'getAsDirectory.ACE = ACE

            myMain.Log.SIS.LogObject("getAsDirectory results: " & getAsDirectory.getDisplayName, getAsDirectory)


        Catch ex As Exception
            setLastError(ex)
            Return Nothing
        End Try
    End Function

#End Region

    Public Overridable Sub Dispose()
        _AccessCache = New NET_Cache
        _ExistsCache = New NET_Cache
    End Sub

#End Region

End Class





<Serializable()> <DataContract()>
Public Class ALECS_Directory_Entry
    Inherits ALECS_FSO_Base
    Implements IALECS_Entry, IALECS_Entry.IDeletable, IALECS_Entry.ICreateable

#Region "VARs"

    'Public VARs for direct Access (Fields):
    <DataMember(EmitDefaultValue:=False, Order:=15)>
    <System.Xml.Serialization.XmlIgnoreAttribute>
    Public DirectoryUri As Uri = Nothing


    'Private VARs (Access only via Properties):
    Private _Directory As IO.DirectoryInfo = Nothing

    ' for Drive Connect
    Private _DriveLetter As String = Nothing
    Private _DriveConnected As Boolean = False

    '# Helper VARs for Caching:
    Private _ExistsCache As New NET_Cache
    Private _AccessCache As New NET_Cache

#End Region

#Region "Constructors"

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New()
        MyBase.New(GetType(ALECS_Directory_Entry))
    End Sub

    ''' <summary>
    ''' This is the MAIN (new) constructor. Please use this!
    ''' So when a connection is needed the automatic detected HOST Credentials via Host.getACE is used
    ''' </summary>
    ''' <param name="dirSource">the path for the directory. E.g. C:\Program Files\Java or for network ressources: "\\" & Host.PrimaryIP & "C$\Program Files\Java</param>
    ''' <param name="objAHE">the given host for a remote directory (definines the credentials)</param>
    Public Sub New(ByRef objMain As Main, dirSource As IO.DirectoryInfo, Optional objAHE As ALECS_Host_Entry = Nothing)
        MyBase.New(objMain, GetType(ALECS_Directory_Entry), objAHE)
        FSO = dirSource
        Host = objAHE
    End Sub

#End Region

#Region "Properties"

    'not needed <DataMember(EmitDefaultValue:=False)>
    <System.Xml.Serialization.XmlIgnoreAttribute>
    Public Property FSO As IO.DirectoryInfo
        Get
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    Exit Select
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
            Return _Directory
        End Get
        Set(value As DirectoryInfo)
            _Directory = value      ' when instance is web based, _Directory is nothing, so NOT an Exception
            If myMain.HELPER.NotAnEmptyValue(_Directory) Then
                If Not _Directory.FullName.EndsWith("\") Then _Directory = New IO.DirectoryInfo(_Directory.FullName & "\")
                'Check if any VARiable is in Destination Path:
                If _Directory.ToString.IndexOf("%%") > -1 Then
                    'Directory contains %% Variables. Check if Drive Letter is in Path:
                    If _Directory.ToString.Substring(1, 1) = ":" Then
                        _Directory = New DirectoryInfo(returnPathWithBackslash(_Directory.ToString.Substring(0, 3) & _Directory.ToString.Substring(3).ToString.Replace(":", "_")))
                    Else
                        _Directory = New DirectoryInfo(returnPathWithBackslash(_Directory.ToString.ToString.Replace(":", "_")))
                    End If
                End If
                setStorage(_Directory.FullName)
            End If
        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=11)>
    Public Overridable Shadows ReadOnly Property Name As String 'Implements IALECS_Entry.Name
        Get
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    Return _Directory.Name
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Return String.Empty
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property

    '# IMPORTANT: Not Readonly, because csm_FSO.Directory can not READ from JSON
    <DataMember(EmitDefaultValue:=False, Order:=21)>
    Public Property Fullname As String
        Get
            Select Case Storage
                Case ALECS_FSO_Storage.UNC, ALECS_FSO_Storage.Drive
                    Return returnPathWithBackslash(_Directory.FullName)
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Return DirectoryUri.AbsoluteUri
                Case Else
                    Return Nothing
            End Select
        End Get
        Set(value As String)
            Select Case Storage
                Case ALECS_FSO_Storage.UNC, ALECS_FSO_Storage.Drive
                    _Directory = New IO.DirectoryInfo(value)
                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    DirectoryUri = New Uri(value)
                Case Else
                    Throw New NotImplementedException("Fullname property (SET) not implemented for " & Storage.ToString & " type Directories")
            End Select
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns>the Drive letter the Directory is or should be Mapped to</returns>
    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Property Drive As String
        Get
            Return _DriveLetter
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then Exit Property
            _DriveLetter = value(0) & ":"
        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public ReadOnly Property MappedAsDrive As Boolean
        Get
            Return _DriveConnected
        End Get
    End Property

    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Overloads ReadOnly Property Created As DateTime? Implements IALECS_Entry.Created
        Get
            Return FSO.CreationTime.ToLocalTime
        End Get
    End Property

    'NO show in Hangfire DashBoard? <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Overloads ReadOnly Property changed As DateTime? Implements IALECS_Entry.changed
        Get
            Return FSO.LastWriteTime.ToLocalTime
        End Get
    End Property

#End Region

#Region "Functions"

#Region "IALECS_Entry Functions"

    Public Function CheckIfExist(Optional bolUseCache As Boolean = True) As Boolean 'Implements IALECS_Entry.ILoadable.CheckIfExist
        CheckIfExist = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            If Not _ExistsCache.Filled Or Not bolUseCache Then
                '# fill Cache:
                Try
                    Select Case Storage
                        Case ALECS_FSO_Storage.Drive
                            If Not checkAccess(FileAccess.Read, bolUseCache) Then
                                Try
                                    Throw _AccessCache.Exception
                                Catch ex As IOException
                                    'The Directory exists and i would have access if no one else has open it. But this happends now :-(
                                Catch ex As Exception
                                    Throw ex
                                End Try
                            End If
                            FSO.Refresh()
                            _ExistsCache.Value = FSO.Exists

                        Case ALECS_FSO_Storage.UNC
                            '# Try 1:
                            If Not checkAccess(FileAccess.Read, bolUseCache) Then
                                'myMain.Log.add("checkAccess can't read on it :-( so we must handle this error: " & _AccessCache.Exception.Message, Me)
                                Try
                                    'throw _AccessCache Exception for Exist (for log)
                                    Throw _AccessCache.Exception


                                Catch ex As Exception
                                    Throw ex
                                End Try
                            End If
                            FSO.Refresh()
                            _ExistsCache.Value = FSO.Exists

                        Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS


                        Case Else
                            Throw New NotImplementedException("Exists check NOT implemented for " & Storage.ToString & " Type files")

                    End Select

                    '# so here we're: and we File did Exists :-)
                    _ExistsCache.Filled = True


                Catch ex As Exception
                    '# so here we're: and the File did NOT Exist :-(
                    _ExistsCache.Value = False
                    _ExistsCache.Filled = True
                    _ExistsCache.Exception = ex
                    setLastError(ex)
                End Try

            Else
                '# use Cache:

                'last line returned: Exists = _ExistCache.bolValue
                If _ExistsCache.Exception IsNot Nothing Then setLastError(_ExistsCache.Exception)

            End If

            myMain.Log.SIS.LogBool("CheckIfExist", _ExistsCache.Value)

            Return _ExistsCache.Value


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            CheckIfExist = False
        End Try
    End Function

    Public Function Delete() As Boolean Implements IALECS_Entry.IDeletable.Delete
        Delete = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            'protect before UnitTest Errors / Fails and accidental delete:
            If Fullname.Equals("\\alecs_config.ccs.cmsystemhaus.net\alecs_storage\TEST\UnitTests\", StringComparison.CurrentCultureIgnoreCase) Then Throw New InvalidProgramException("OH NO, do NOT delete: \\alecs_config.ccs.cmsystemhaus.net\alecs_storage\TEST\UnitTests\ PLEASE :-o")

            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    If Storage = ALECS_FSO_Storage.UNC Then
                        If Not checkAccess(FileAccess.Write, False, False) Then
                            If TypeOf getLastError() Is DirectoryNotFoundException Then
                                Return True
                            End If
                            ''If Not DeleteWMI() Then Throw getLastError()
                        End If
                    End If

                    'check if already NOT exist:
                    _Directory.Refresh()
                    If Not _Directory.Exists Then Return True

                    '# Check if Operation Cancelled:
                    myMain.CTN.ThrowIfCancellationRequested()

                    Try
                        'try to delete normaly with DeleteWMI backup (catch) (ahe must be specified!)!
                        _Directory.Delete(True)
                        'alternative: My.Computer.FileSystem.DeleteDirectory(myDirectory.FullName, FileIO.DeleteDirectoryOption.DeleteAllContents)


                    Catch ex As Exception
                        ''If Not DeleteWMI() Then Throw getLastError()
                    End Try

                    _Directory.Refresh()
                    Delete = Not _Directory.Exists

                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Throw New NotImplementedException("Delete function not implemented for HTTP/HTTPS directories")

                Case Else
                    Throw New NotImplementedException("Delete function not implemented for " & Storage.ToString & " Type directories")
            End Select

            Dispose()
            '_Loaded = False


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Overloads Function Create() As Boolean Implements IALECS_Entry.ICreateable.Create
        Create = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    Dim ioDirectory As New DirectoryInfo(Fullname)
                    ioDirectory.Create()
                    Create = True

                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Throw New NotImplementedException("ALECS_Directory_Entry.Create --> StorageType=" & Storage.ToString)

                Case Else
                    Throw New NotImplementedException("ALECS_Directory_Entry.Create --> StorageType=" & Storage.ToString)
            End Select

            'reset Cache:
            _ExistsCache = New NET_Cache
            _AccessCache = New NET_Cache


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Overrides Function getDisplayName() As String Implements IALECS_Entry.getDisplayName
        getDisplayName = String.Empty
        clearLastError()
        Try
            If _Directory IsNot Nothing Then
                getDisplayName = _Directory.Name & " in (" & Fullname & ")"
            End If
            getDisplayName += ", Type=Directory"
            If Storage.HasValue Then
                If Not String.IsNullOrEmpty(getDisplayName) Then getDisplayName += ", "
                getDisplayName += "Storage=" & Storage.ToString
            End If


        Catch ex As Exception
            setLastError(ex)
        End Try
    End Function

    Public Overrides Sub DebugMe(Optional strPREfix As String = "") Implements IALECS_Entry.DebugMe
        Try
            myMain.Log.SIS.LogColored(Color.LightGray, "DEBUGME: " & getDisplayName())
            myMain.Log.SIS.LogObject(strPREfix & "ALECS_Directory_Entry: " & getDisplayName(), Me)

#Region "Details"
            Dim strDetails As String = "DETAILS"
            myMain.Log.SIS.EnterMethod(strDetails)
            If String.IsNullOrEmpty(strPREfix) Then strPREfix = "DIR "
            'Host:
            If Host IsNot Nothing Then
                myMain.Log.SIS.LogObject("...HOST", Host)
                'Host.DebugMe(strPREfix)
            Else
                'myMain.Log.add("HOST is Nothing", Me)
            End If
            myMain.Log.SIS.LeaveMethod(strDetails)
#End Region


        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Basic FSO Functions"

    <Obsolete("This function is depricated. Please use CheckIfExist function instead!", True)>
    Public Function Exists(Optional useCache As Boolean = True) As Boolean
        Return CheckIfExist(useCache)
    End Function

    Public Function DeleteContent(strFiles As String, Optional bolIgnoreErrors As Boolean = False) As Boolean
        DeleteContent = False
        Dim myTiming As New Timing
        clearLastError()
        Try
            If strFiles.Equals("*") Then
                '# Delete All SubDirectories in Folder:
                Select Case Storage
                    Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                        For Each objSubDirectory As DirectoryInfo In _Directory.GetDirectories()
                            Dim dir2Delete As New ALECS_Directory_Entry(myMain, objSubDirectory, Host)
                            If Not dir2Delete.Delete() AndAlso Not bolIgnoreErrors Then Throw getLastError()
                        Next
                        '# Delete All Files in Folder:
                        Dim delFiles As New ArrayList(IO.Directory.GetFiles(_Directory.FullName, strFiles))
                        myMain.Log.SIS.LogEnumerable("List of Files in Directory (to delete!)", delFiles)
                        For Each strFileName As String In delFiles
                            Dim fil2Delete As New ALECS_File_Entry(myMain, New IO.FileInfo(strFileName), Host)
                            If Not fil2Delete.Delete() AndAlso Not bolIgnoreErrors Then Throw getLastError()
                        Next

                    Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                        Throw New NotImplementedException("Content Deletion Actions for Web Based Directories are not implemented yet")

                    Case Else
                        Throw New NotImplementedException("DeleteContent function is not implemented for " & Storage.ToString & " type Directories")

                End Select
            Else
                '# Delete ONLY selected Types (Files+SubDirectories) in Folder:
                Select Case Storage
                    Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC

                        'CASE "*." handling ??

                        '# Delete Files in Folder:
                        Dim delFiles As New ArrayList(IO.Directory.GetFiles(_Directory.FullName, strFiles))
                        myMain.Log.SIS.LogEnumerable("List of Files in Directory (to delete? Because matches: " & strFiles & ")", delFiles)
                        For Each strFileName As String In delFiles
                            Dim fil2Delete As New ALECS_File_Entry(myMain, New IO.FileInfo(strFileName), Host)
                            If Not fil2Delete.Delete() AndAlso Not bolIgnoreErrors Then Throw getLastError()
                        Next

                    Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                        'TODO 2DO implement this
                        Throw New NotImplementedException("Content Deletion Actions for Web Based Directories are not implemented yet")

                    Case Else
                        Throw New NotImplementedException("DeleteContent function is not implemented for " & Storage.ToString & " type Directories")

                End Select

            End If

            DeleteContent = True


        Catch ex As Exception
            setLastError(ex)
            DeleteContent = False
        End Try
    End Function

    Public Function Copy(ByRef dirDestination As ALECS_Directory_Entry, Optional bolOverwrite As Boolean = True) As Boolean
        Copy = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            Select Case Storage
                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    If Fullname.Equals(dirDestination.Fullname, StringComparison.CurrentCultureIgnoreCase) Then Throw New InvalidOperationException("Source and Destination Directory are equal!")

                    'Check Access Source:
                    If Not checkAccess(FileAccess.Read, False, True) Then Throw getLastError()

                    '# in case of problem with UNC access add the following lines of code to check / make access before!
                    If Not dirDestination.CheckIfExist(True) Then If Not dirDestination.Create() Then Throw dirDestination.getLastError

                    '# Check if Operation Cancelled:
                    myMain.CTN.ThrowIfCancellationRequested()

                    ' Check if the Target Directory exists, if not, create it.
                    My.Computer.FileSystem.CopyDirectory(Fullname, dirDestination.Fullname, bolOverwrite)
                    Copy = True

                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    Throw New NotImplementedException("Copy Actions for web based Directories are not implemented yet")

                Case Else
                    Throw New NotImplementedException("COPY Function is Not implemented for " & Storage.ToString & " type directories.")

            End Select

            '_ExistCache = New ALECS_Cache
            dirDestination._ExistsCache = New NET_Cache()


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
            'detailed exception is in ex.DATA
            Copy = False
        End Try
    End Function

    Public Function Move(ByRef dirDestination As ALECS_Directory_Entry, Optional bolOverwrite As Boolean = True) As Boolean
        Move = False
        Dim timTiming As New Timing
        clearLastError()
        Try
            Select Case Storage

                Case ALECS_FSO_Storage.Drive, ALECS_FSO_Storage.UNC
                    If Fullname.Equals(dirDestination.Fullname, StringComparison.CurrentCultureIgnoreCase) Then Throw New InvalidOperationException("Source and Destination Directory are equal!")

                    'Check Access Source:
                    If Not checkAccess(FileAccess.Read, False, True) Then Throw getLastError()

                    '# in case of problem with UNC access add the following lines of code to check / make access before!
                    If Not dirDestination.CheckIfExist(True) Then If Not dirDestination.Create() Then Throw dirDestination.getLastError

                    '# Check if Operation Cancelled:
                    myMain.CTN.ThrowIfCancellationRequested()

                    My.Computer.FileSystem.MoveDirectory(Fullname, dirDestination.Fullname, bolOverwrite)
                    Move = True

                Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS
                    'Move = MyBase.Move(Fullname, dirDestination, bolOverwrite)
                    Throw New NotImplementedException("MOVE function from HTTP(S) is NOT unitTested @ the moment!")

                Case Else
                    Throw New NotImplementedException("MOVE Function not implemented for " & Storage.ToString & " Type directories.")

            End Select

            _ExistsCache = New NET_Cache
            dirDestination._ExistsCache = New NET_Cache


        Catch ex As OperationCanceledException
            setLastError(ex)
            Throw
        Catch ex As Exception
            setLastError(ex)
            Move = False
        End Try
    End Function

#End Region

#Region "checkFunctions"

    Public Function checkAccess(Optional AccessType As FileAccess = FileAccess.Read, Optional useCache As Boolean = True, Optional bolMustExist As Boolean = False) As Boolean
        checkAccess = False
        clearLastError()
        Try
            If Not _AccessCache.Filled Or Not useCache Then
                '# fill Cache:
                Try
                    Dim intRetry As Integer = 2
                    Select Case Storage

                        Case ALECS_FSO_Storage.Drive
                            Try
                                '# Solution from here (see below): https://stackoverflow.com/questions/130617/how-do-you-check-for-permissions-to-write-to-a-directory-or-file
                                Dim arcRule As System.Security.AccessControl.AuthorizationRuleCollection = System.IO.Directory.GetAccessControl(Fullname).GetAccessRules(True, True, GetType(System.Security.Principal.NTAccount))
                                For Each rule As System.Security.AccessControl.FileSystemAccessRule In arcRule
                                    If rule.AccessControlType = System.Security.AccessControl.AccessControlType.Allow Then
                                        'If AppInDebugMode Then myMain.Log.SIS.LogBool("checkAccess", True)
                                        Exit For
                                    End If
                                Next


                            Catch ex As DirectoryNotFoundException
                                If bolMustExist Then
                                    myMain.Log.SIS.LogWarning("The file '" & _Directory.FullName & "' does NOT exist (bolMustExist=" & bolMustExist & "): " & vbNewLine & ex.Message)
                                    Throw ex
                                Else
                                    Return True
                                End If
                            End Try

                        Case ALECS_FSO_Storage.UNC
                            Try
startAccessUNC:
                                '# Solution from here (see below): https://stackoverflow.com/questions/130617/how-do-you-check-for-permissions-to-write-to-a-directory-or-file
                                Dim secDirectory As System.Security.AccessControl.DirectorySecurity = Nothing, intCounterRetrySecDirectory As Integer = 5
retrySecFile:
                                Try
                                    secDirectory = IO.Directory.GetAccessControl(_Directory.FullName)


                                Catch ex As Exception
                                    myMain.DP(22841)
                                End Try

                                Dim arcRule As System.Security.AccessControl.AuthorizationRuleCollection = secDirectory.GetAccessRules(True, True, GetType(System.Security.Principal.NTAccount))
                                For Each fsaRule As System.Security.AccessControl.FileSystemAccessRule In arcRule
                                    If fsaRule.AccessControlType = System.Security.AccessControl.AccessControlType.Allow Then
                                        myMain.Log.SIS.LogBool("checkAccess (" & _Directory.FullName & ")", True)
                                        Exit For
                                    End If
                                Next


                            Catch ex As DirectoryNotFoundException
                                'handled in Catch block before! Only for jump over arcRule
                                Throw ex
                            Catch ex As UnauthorizedAccessException
                                'handled in Catch block before! Only for jump over arcRule
                                Throw ex
                            Catch ex As InvalidOperationException
                                'handled in Catch block before! Only for jump over arcRule
                                Throw ex
                            Catch ex As NullReferenceException
                                myMain.DP(22861)
                            Catch ex As Exception
                                Throw ex
                            End Try


                        Case ALECS_FSO_Storage.HTTP, ALECS_FSO_Storage.HTTPS


                        Case Else
                            Throw New NotImplementedException("checkAccess function not implemented for " & Storage & " type directories")

                    End Select

                    '# so here we're: and we HAVE Access :-)
                    _AccessCache.Value = True
                    _AccessCache.Filled = True


                Catch ex As Exception
                    '# so here we're: and we have NO Access :-(
                    _AccessCache.Value = False
                    _AccessCache.Filled = True
                    _AccessCache.Exception = ex
                    setLastError(ex)
                    'really needed an eror? Throw
                End Try

            Else
                '# use Cache:

                'last line returned checkAccess = _AccessCache.bolValue
                If _AccessCache.Exception IsNot Nothing Then setLastError(_AccessCache.Exception)

            End If

            checkAccess = _AccessCache.Value
            myMain.Log.SIS.LogBool("checkAccess (" & _Directory.FullName & ")", checkAccess)


        Catch ex As Exception
            setLastError(ex)
            Return False
        End Try
    End Function

#End Region

    Public Overridable Sub Dispose()
        'If _DriveConnected Then DriveDisconnect()
        'DriveDisconnect()
        _AccessCache = New NET_Cache
        _ExistsCache = New NET_Cache
    End Sub

#End Region

End Class
