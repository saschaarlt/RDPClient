
Imports System.Runtime.Serialization
Imports Gurock.SmartInspect



'# Best Practise for Exceptions:
'# https://msdn.microsoft.com/de-de/library/seyhszts(v=vs.110).aspx
<Serializable()> <DataContract()>
Public Class cms_Error
    'Inherits cms_LastError

#Region "VARs"

    'MAIN:
    Dim myMain As Main

    'Public VARs for direct Access (Fields):
    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Shared PublicErrorSub(1) As Object

    <Serializable()> <DataContract()>
    Public Enum MessageBoxOnError_Type As Byte
        NO = 0
        INFO = 1
        SELECT_ABORT = 2
    End Enum


    'Private VARs (Access only via Properties):
    Private _ServiceMode As Boolean = False
    Private _MessageBoxOnError As MessageBoxOnError_Type = MessageBoxOnError_Type.NO
    Private _QuitOnError As Boolean = False

#End Region

#Region "Properties"

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Property ServiceMode As Boolean
        Get
            Return _ServiceMode
        End Get
        Set(value As Boolean)
            _ServiceMode = value
        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Property MessageboxOnError As MessageBoxOnError_Type
        Get
            Return _MessageBoxOnError
        End Get
        Set(value As MessageBoxOnError_Type)
            _MessageBoxOnError = value
        End Set
    End Property

    <DataMember(EmitDefaultValue:=False, Order:=99)>
    Public Property QuitOnError As Boolean
        Get
            Return _QuitOnError
        End Get
        Set(value As Boolean)
            _QuitOnError = value
        End Set
    End Property

#End Region

#Region "Constructors"

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New()

    End Sub

    Public Sub New(ByRef objMain As Main)
        myMain = objMain
    End Sub

#End Region

#Region "Functions"

    '###
    'https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/concepts/caller-information

    Public Sub handleException(excError As Exception, objMe As Object,
        <System.Runtime.CompilerServices.CallerMemberName> Optional memberName As String = Nothing,
        <System.Runtime.CompilerServices.CallerFilePath> Optional sourcefilePath As String = Nothing,
        <System.Runtime.CompilerServices.CallerLineNumber()> Optional sourceLineNumber As Integer = 0)

        If excError.GetType() = GetType(OperationCanceledException) Then
            myMain.Log.OperationCanceledException(Me)
            Exit Sub
        Else
            '+++ Setze Fehlerbehandlung (intern) für den Service Modi:
            If _ServiceMode Then On Error Resume Next

            Dim strModulClassName As String = String.Empty, strSubFuncName As String = String.Empty

            If TypeOf objMe Is String AndAlso objMe.ToString.Contains(":") Then
                Dim arrMe As Array = objMe.ToString.Split(":")
                'getException(excError, arrMe(0), arrMe(1), "", ErrorCode, False)
                strModulClassName = arrMe(0)
                strSubFuncName = arrMe(1)
            Else
                '2DO:
                'Wie wird das Module und funktion ausgelesen?
                'getException(excError, objMe.ToString, objMe.ToString, "", ErrorCode, False)
                strModulClassName = objMe.ToString
                strSubFuncName = objMe.ToString
            End If


            'this works but spamed log, also Line is the line from myMain.Error.handleException(ex, Me)
            'better copy *.pdb File for exact LineNumbers!
            'myMain.Log.SIS.LogString("Exception in file", sourcefilePath)
            'myMain.Log.SIS.LogString("Exception in method", memberName)
            'myMain.Log.SIS.LogString("Exception at line numer", sourceLineNumber)
            myMain.Log.SIS.LogException(excError)



            'If excError.Message = "Ungültiger Leseversuch, wenn keine Daten vorhanden sind." Then
            '    'vergessen ein sdrxxxx.read zu machen!!!
            '    If CMSMsgBox(myMain, "Fehler in Modul " & strModulClassName & " (" & strSubFuncName & "):" & vbNewLine & excError.Message & vbNewLine & vbNewLine & excError.TargetSite.ToString & vbNewLine & vbNewLine & "TIPP: Sie haben wahrscheinlich nur vergessen ein sdr_XXXX.Read() zu machen!", "cms_Error: " & strSubFuncName, CMSMsgBoxButton.OKCancel, CMSMsgBoxIcon.Error_, CMSMsgBoxDefaultButton.Button1, CMSMsgBoxOptions.DefaultDesktopOnly, False) = CMSMsgBoxDialogResult.Cancel AndAlso MessageboxOnError = MessageBoxOnError_Type.SELECT_ABORT Then
            '        Environment.Exit(0)
            '    End If

            'ElseIf excError.Message = "Zeichenfolgen- oder Binärdaten würden abgeschnitten." & vbNewLine & "Die Anweisung wurde beendet." Then
            '    'Daten sind zum anfügen an die Tabelle zu lang!
            '    If CMSMsgBox(myMain, "Fehler in Modul " & strModulClassName & " (" & strSubFuncName & "):" & vbNewLine & excError.Message & vbNewLine & vbNewLine & excError.TargetSite.ToString & vbNewLine & vbNewLine & "TIPP: SQL-ERROR=Ein Wert ist zu Lang, und das Feld zu kurz! SQL-Statement fehlgeschlagen!", "cms_Error: " & strSubFuncName, CMSMsgBoxButton.OKCancel, CMSMsgBoxIcon.Error_, CMSMsgBoxDefaultButton.Button1, CMSMsgBoxOptions.DefaultDesktopOnly, False) = CMSMsgBoxDialogResult.Cancel AndAlso MessageboxOnError = MessageBoxOnError_Type.SELECT_ABORT Then
            '        Environment.Exit(0)
            '    End If
            'ElseIf Not _ServiceMode Then
            '    If CMSMsgBox(myMain, "Fehler in Modul " & strModulClassName & " (" & strSubFuncName & "):" & vbNewLine & excError.Message & vbNewLine & vbNewLine & excError.StackTrace.ToString, "cms_Error: " & strSubFuncName, CMSMsgBoxButton.OKCancel, CMSMsgBoxIcon.Error_, CMSMsgBoxDefaultButton.Button1, CMSMsgBoxOptions.DefaultDesktopOnly, False) = CMSMsgBoxDialogResult.Cancel AndAlso MessageboxOnError = MessageBoxOnError_Type.SELECT_ABORT Then
            '        Environment.Exit(0)
            '    End If

            '    'ElseIf myMain.Error.MessageboxOnError Then
            '    '    If CMSMsgBox(myMain, "Fehler in Modul " & strModulClassName & " (" & strSubFuncName & "):" & vbNewLine & excError.Message & vbNewLine & vbNewLine & excError.StackTrace.ToString, "cms_Error: " & strSubFuncName, CMSMsgBoxButton.OKCancel, CMSMsgBoxIcon.Error_, CMSMsgBoxDefaultButton.Button1, CMSMsgBoxOptions.DefaultDesktopOnly, False) = CMSMsgBoxDialogResult.Cancel AndAlso MessageboxOnError = MessageBoxOnError_Type.SELECT_ABORT Then
            '    '        Environment.Exit(0)
            '    '    End If

            'End If

            '+++ Versuche Error Sub des Programmes aus zu führen:
            If Not PublicErrorSub(1).Equals("") Then CallByName(PublicErrorSub(0), PublicErrorSub(1).ToString, CallType.Get)

            '+++ Versuche das Programm zu beenden:
            If _QuitOnError Then Environment.Exit(0)

        End If
    End Sub

    '<Obsolete("This sub is depricated. Please use new sub instead!", False)>
    Public Sub getException(excError As Exception, objMe As Object, Optional siLevel As Level? = Nothing)

        If excError.GetType() = GetType(OperationCanceledException) Then
            myMain.Log.OperationCanceledException(Me)
            Exit Sub
        Else
            '+++ Setze Fehlerbehandlung (intern) für den Service Modi:
            If _ServiceMode Then On Error Resume Next

            Dim strModulClassName As String = String.Empty, strSubFuncName As String = String.Empty

            If TypeOf objMe Is String AndAlso objMe.ToString.Contains(":") Then
                Dim arrMe As Array = objMe.ToString.Split(":")
                'getException(excError, arrMe(0), arrMe(1), "", ErrorCode, False)
                strModulClassName = arrMe(0)
                strSubFuncName = arrMe(1)
            Else
                '2DO:
                'Wie wird das Module und funktion ausgelesen?
                'getException(excError, objMe.ToString, objMe.ToString, "", ErrorCode, False)
                strModulClassName = objMe.ToString
                strSubFuncName = objMe.ToString
            End If

            '#1 SmartInspectLogging (SIL)
            If siLevel Is Nothing Then

                myMain.Log.SIS.LogException(excError)

            Else

                Select Case siLevel
                    'Case cms_Error.ErrorCode.Information
                    '    myMain.Log.SIS.LogException(excError)
                    Case Level.Warning
                        myMain.Log.SIS.LogWarning(excError.Message)

                    Case Level.Fatal
                        myMain.Log.SIS.LogException(excError)
                        myMain.Log.SIS.LogFatal(excError.Message)

                        'Case cms_Error.ErrorCode.AdviceSupport
                        '    myMain.Log.SIS.LogException(excError)
                        '    myMain.Log.SIS.LogFatal(excError.Message)
                        '2DO Supportaction
                        'Case cms_Error.ErrorCode.AdviceDevelopment
                        '    myMain.Log.SIS.LogException(excError)
                        '    myMain.Log.SIS.LogFatal(excError.Message)
                        '2DO Developmentaction

                    Case Else
                        myMain.Log.SIS.LogException(excError)
                        'write in central log
                End Select

            End If


            'If excError.Message = "Ungültiger Leseversuch, wenn keine Daten vorhanden sind." Then
            '    'vergessen ein sdrxxxx.read zu machen!!!
            '    If CMSMsgBox(myMain, "Fehler in Modul " & strModulClassName & " (" & strSubFuncName & "):" & vbNewLine & excError.Message & vbNewLine & vbNewLine & excError.TargetSite.ToString & vbNewLine & vbNewLine & "TIPP: Sie haben wahrscheinlich nur vergessen ein sdr_XXXX.Read() zu machen!", "cms_Error: " & strSubFuncName, CMSMsgBoxButton.OKCancel, CMSMsgBoxIcon.Error_, CMSMsgBoxDefaultButton.Button1, CMSMsgBoxOptions.DefaultDesktopOnly, False) = CMSMsgBoxDialogResult.Cancel AndAlso MessageboxOnError = MessageBoxOnError_Type.SELECT_ABORT Then
            '        Environment.Exit(0)
            '    End If

            'ElseIf excError.Message = "Zeichenfolgen- oder Binärdaten würden abgeschnitten." & vbNewLine & "Die Anweisung wurde beendet." Then
            '    'Daten sind zum anfügen an die Tabelle zu lang!
            '    If CMSMsgBox(myMain, "Fehler in Modul " & strModulClassName & " (" & strSubFuncName & "):" & vbNewLine & excError.Message & vbNewLine & vbNewLine & excError.TargetSite.ToString & vbNewLine & vbNewLine & "TIPP: SQL-ERROR=Ein Wert ist zu Lang, und das Feld zu kurz! SQL-Statement fehlgeschlagen!", "cms_Error: " & strSubFuncName, CMSMsgBoxButton.OKCancel, CMSMsgBoxIcon.Error_, CMSMsgBoxDefaultButton.Button1, CMSMsgBoxOptions.DefaultDesktopOnly, False) = CMSMsgBoxDialogResult.Cancel AndAlso MessageboxOnError = MessageBoxOnError_Type.SELECT_ABORT Then
            '        Environment.Exit(0)
            '    End If
            'ElseIf Not _ServiceMode Then
            '    If CMSMsgBox(myMain, "Fehler in Modul " & strModulClassName & " (" & strSubFuncName & "):" & vbNewLine & excError.Message & vbNewLine & vbNewLine & excError.StackTrace.ToString, "cms_Error: " & strSubFuncName, CMSMsgBoxButton.OKCancel, CMSMsgBoxIcon.Error_, CMSMsgBoxDefaultButton.Button1, CMSMsgBoxOptions.DefaultDesktopOnly, False) = CMSMsgBoxDialogResult.Cancel AndAlso MessageboxOnError = MessageBoxOnError_Type.SELECT_ABORT Then
            '        Environment.Exit(0)
            '    End If

            '    'ElseIf myMain.Error.MessageboxOnError Then
            '    '    If CMSMsgBox(myMain, "Fehler in Modul " & strModulClassName & " (" & strSubFuncName & "):" & vbNewLine & excError.Message & vbNewLine & vbNewLine & excError.StackTrace.ToString, "cms_Error: " & strSubFuncName, CMSMsgBoxButton.OKCancel, CMSMsgBoxIcon.Error_, CMSMsgBoxDefaultButton.Button1, CMSMsgBoxOptions.DefaultDesktopOnly, False) = CMSMsgBoxDialogResult.Cancel AndAlso MessageboxOnError = MessageBoxOnError_Type.SELECT_ABORT Then
            '    '        Environment.Exit(0)
            '    '    End If

            'End If

            '+++ Versuche Error Sub des Programmes aus zu führen:
            If PublicErrorSub(1) IsNot Nothing AndAlso Not PublicErrorSub(1).Equals("") Then CallByName(PublicErrorSub(0), PublicErrorSub(1).ToString, CallType.Get)

            '+++ Versuche das Programm zu beenden:
            If _QuitOnError Then Environment.Exit(0)

        End If

    End Sub



    Public Sub ThrowArgumentTypeUnknownException(objArgumentType As Object)
        Throw New ArgumentOutOfRangeException(objArgumentType.GetType.ToString, objArgumentType.ToString, "This argument is INVALID!")
    End Sub

    Public Sub ThrowDeveleopmentException()
        Throw New MulticastNotSupportedException("Hey Developer. Why are we here?")
    End Sub

    Public Sub ThrowSupportException()
        Throw New NotSupportedException("Hey Supporter. This is NOT good. !You know what you have to do! Please act now. Thanks.")
    End Sub

#End Region

End Class
