
'### CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION #'
'#
'# Diese Funktion STÜRZT ab wenn folgende Befehle enthalten sind:
'# auf DB01+TSMGMT: --> Dim mimeMessage As MimeKit.MimeMessage = ameToImport.getMimeMessage()
'# auf DB01:        --> New MailAction(myMain).addJob(MailAction.MailActionType.CheckMailFolder4Zammad, msxMajaServer.getItem("AAMkAGE1MmE4MTMwLTY4MjAtNDk5MC1hYTcwLTlhNWVjZmJhNjU4YgBGAAAAAABna/tzV4tlSbcigMghaNxkBwB84duJ4PSATZ/45zJg2fBxAAAA/BdnAAB84duJ4PSATZ/45zJg2fBxAAAfoOMJAAA="))
'#
'### CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION #'



<Serializable()> <DataContract()>
Public Class APPStart
    Inherits cms_LastError

#Region "VARs"

    'Public VARs for direct Access (Fields):

    'Private VARs (Access only via Properties):

#End Region

#Region "Constructors"

    Public Sub New(ByRef objMain As Main)
        'SiAuto.Main.LogSeparator()
    End Sub

#End Region

#Region "Functions"

    Public Sub GO()
        clearLastError()
        Dim timTiming As New Timing
        Try
            'Early Start-Ups:
            'Startparameter auswerten / ggf. nur gewisse Aktionen machen:
            For Each strStartArgument As String In My.Application.CommandLineArgs

                ' Add code here to use the argument.
                Select Case strStartArgument

                    Case "test"
                        'myMain.Stage = Main.StageType.TEST

                    Case "PasswordProtected".ToLower
                        Dim myTest As New Password_Tests()
                        'myTest.checkWord_Test()

                    Case "satest"
                        'Dim myMA As New MailAction(myMain)
                        '!!! Absturz !!!
                        'myMA.addJob(MailAction.MailActionType.CheckMailFolder4Zammad, msxMajaServer.getItem("AAMkAGE1MmE4MTMwLTY4MjAtNDk5MC1hYTcwLTlhNWVjZmJhNjU4YgBGAAAAAABna/tzV4tlSbcigMghaNxkBwB84duJ4PSATZ/45zJg2fBxAAAA/BdnAAB84duJ4PSATZ/45zJg2fBxAAAfoOMJAAA="))
                        'myMA.addJob(MailAction.MailActionType.ImportMail2Zammad, New cms_Exchange(myMain).getItem("AAMkAGE1MmE4MTMwLTY4MjAtNDk5MC1hYTcwLTlhNWVjZmJhNjU4YgBGAAAAAABna/tzV4tlSbcigMghaNxkBwB84duJ4PSATZ/45zJg2fBxAAAA/BdnAAB84duJ4PSATZ/45zJg2fBxAAAfoOVIAAA="), 68194)

                        GoTo wait4Close


                End Select


            Next


            '### PART 2 ###'


            'Select Case myMain.AppType

            '    Case Main.ApplicationType.Console
            '        myMain.Log.add(My.Application.Info.ProductName & " " & My.Application.Info.Version.ToString & " is running as CONSOLE !!!", Me, Level.Message, Color.GreenYellow)

            '        'System.Windows.Forms.Application.Run(New frmMain(myMain))
            '        System.Windows.Forms.Application.Run(New frmALECSJobs(myMain))


            '    Case Main.ApplicationType.Service
            '        myMain.Log.add(myMain.ProductName() & " " & myMain.ProductVersion.ToString & " is running as SERVICE MODE !!!", Me, Level.Message, Color.GreenYellow)

            '        System.Windows.Forms.Application.Run(New frmALECSJobs(myMain))

            'End Select






            '#########################################################################
            '#########################################################################
            '#########################################################################





wait4Close:


            'Startparameter auswerten:
            For Each startArg As String In My.Application.CommandLineArgs
                ' Add code here to use the argument.

            Next

            'myMain.Log.add("Runtime Service: " & timTiming.Duration(Timing.TimingFormat.Runtime), Me, Level.Debug, Color.OliveDrab)


        Catch ex As Exception
            'myMain.Error.getException(ex, Me, Level.Fatal)
            'setLastError(ex)
        End Try
    End Sub

#End Region

End Class
