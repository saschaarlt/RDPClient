
'### CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION #'
'#
'# Diese Funktion STÜRZT ab wenn folgende Befehle enthalten sind:
'# auf DB01+TSMGMT: --> Dim mimeMessage As MimeKit.MimeMessage = ameToImport.getMimeMessage()
'# auf DB01:        --> New MailAction(myMain).addJob(MailAction.MailActionType.CheckMailFolder4Zammad, msxMajaServer.getItem("AAMkAGE1MmE4MTMwLTY4MjAtNDk5MC1hYTcwLTlhNWVjZmJhNjU4YgBGAAAAAABna/tzV4tlSbcigMghaNxkBwB84duJ4PSATZ/45zJg2fBxAAAA/BdnAAB84duJ4PSATZ/45zJg2fBxAAAfoOMJAAA="))
'#
'### CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION CAUTION #'

Imports System.Runtime.Serialization



<Serializable()> <DataContract()>
Public Class APPStart
    Inherits cms_LastError
    'Inherits ALECS_App_Start

#Region "VARs"

    Dim myMain As Main

    'Public VARs for direct Access (Fields):

    'Private VARs (Access only via Properties):

#End Region

#Region "Constructors"

    Public Sub New(ByRef objMain As Main)
        myMain = objMain
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

                    Case "rdpconnect1".ToLower
                        'Dim myTest As New Password_Tests()
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

            Dim myRDPClient1 As New ALECS_RDP(mymain)

            '# also see Unit Tests for login data

            Dim netCred As New Net.NetworkCredential()
            With netCred
                .UserName = "juergen.flock@acquisio.de"
                .Password = "o+I<Wr[n0{zheO3"
            End With

            Dim strConFile As String = myMain.AppPath.ToString.Replace("bin\Debug\", "connections\") & "cpub-TSFARMDEV-CmsRdsh.rdp"

            myRDPClient1.LogIn(netCred, strConFile)






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
