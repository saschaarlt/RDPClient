
Imports Microsoft.VisualStudio.TestTools.UnitTesting


<TestClass()>
Public Class RDPLogin_Tests
    Inherits cms_Test

#Region "Test Cases / Test Parameter (as Properties/Functions)"

    Private ardp2Test As ALECS_RDP = Nothing

#Region "Constructors"

    Public Sub New()
        MyBase.New()

        ardp2Test = New ALECS_RDP(myMain)
        checkNEW(ardp2Test)

    End Sub

#End Region

    'Public Function getTestItem() As ALECS_WinService_Entry
    '    getTestItem = New ALECS_WUS(myMain, New Guid("FB1D1753-7E37-48E5-93F0-2C9FEED8079B"))          '# ACMPClient
    'End Function

    'Public Function buildNEWItem() As ALECS_WinService_Entry
    '    buildNEWItem = New ALECS_WinService_Entry(myMain) With {
    '        .UID = New Guid("19ae76ba-b7c6-4e08-a39e-6a124b819aec"),
    '        .Active = False,
    '    }
    'End Function

#End Region

    <TestMethod()>
    Public Sub LogIn_User1_Test()

        '# Prepare:
        Dim netCred As New Net.NetworkCredential()
        With netCred
            .UserName = "juergen.flock@acquisio.de"
            .Password = "o+I<Wr[n0{zheO3"
        End With
        Dim strConFile As String = myMain.AppPath.ToString.Replace("bin\Debug\", "connections\") & "cpub-TSFARMDEV-CmsRdsh.rdp"



        '# Login with user 1
        Dim rlrLogin = ardp2Test.LogIn(netCred, strConFile)
        checkSuccessful(rlrLogin)


        logSuccessfulCheckEnd("ALECS_RDP Successful Login with User 1: " & netCred.UserName, "runs and ended successful!")

    End Sub

    <TestMethod()>
    Public Sub LogIn_User2_Test()

        '# Prepare:
        Dim netCred As New Net.NetworkCredential()
        With netCred
            .UserName = "juergen.flock@acquisio.de"
            .Password = "o+I<Wr[n0{zheO3"
        End With
        Dim strConFile As String = myMain.AppPath.ToString.Replace("bin\Debug\", "connections\") & "cpub-TSFARMDEV-unsigned.rdp"



        '# Login with user 1
        Dim rlrLogin = ardp2Test.LogIn(netCred, strConFile)
        checkSuccessful(rlrLogin)


        logSuccessfulCheckEnd("ALECS_RDP Successful Login with User 2: " & netCred.UserName, "runs and ended successful!")

    End Sub

    <TestMethod()>
    Public Sub LogIn_WrongPassword_Test()

        '# Prepare:
        Dim netCred As New Net.NetworkCredential()
        With netCred
            .UserName = "juergen.flock@acquisio.de"
            .Password = "wrong password"
        End With
        Dim strConFile As String = myMain.AppPath.ToString.Replace("bin\Debug\", "connections\") & "cpub-TSFARMDEV-CmsRdsh.rdp"



        '# Login with user 1
        Dim rlrLogin = ardp2Test.LogIn(netCred, strConFile)
        checkFailed(rlrLogin)


        logSuccessfulCheckEnd("ALECS_RDP Failure Login with User: " & netCred.UserName, "fails successful!")

    End Sub

    <TestMethod()>
    Public Sub LogIn_WongUser_Test()

        '# Prepare:
        Dim netCred As New Net.NetworkCredential()
        With netCred
            .UserName = "nobody@isfalse.com"
            .Password = "Also a strong and Secure PassPhrase with special Chars"
        End With
        Dim strConFile As String = myMain.AppPath.ToString.Replace("bin\Debug\", "connections\") & "cpub-TSFARMDEV-CmsRdsh.rdp"



        '# Login with user 1
        Dim rlrLogin = ardp2Test.LogIn(netCred, strConFile)
        checkFailed(rlrLogin)


        logSuccessfulCheckEnd("ALECS_RDP fails Login with User: " & netCred.UserName, "fails successful in case of wrong user!")

    End Sub

    <TestMethod()>
    Public Sub LogIn_WrongFile_Test()

        '# Prepare:
        Dim netCred As New Net.NetworkCredential()
        With netCred
            .UserName = "nobody@isfalse.com"
            .Password = "Also a strong and Secure PassPhrase with special Chars"
        End With
        Dim strConFile As String = myMain.AppPath.ToString.Replace("bin\Debug\", "connections\") & "cpub-TSFARMDEV-defect.rdp"



        '# Login with user 1
        Dim rlrLogin = ardp2Test.LogIn(netCred, strConFile)
        checkFailed(rlrLogin)


        logSuccessfulCheckEnd("ALECS_RDP fails Login with User: " & netCred.UserName, "fails successful in case of defect file!")

    End Sub

    Private Sub checkSuccessful(rlrResult As ALECS_RDP.RDPLoginResult)
        Assert.IsTrue(rlrResult.Successful)
        Assert.IsNull(rlrResult.Error)
        Assert.IsNotNull(rlrResult.Start)
        Assert.IsNotNull(rlrResult.TimeToLogin)
    End Sub

    Private Sub checkFailed(rlrResult As ALECS_RDP.RDPLoginResult)
        Assert.IsFalse(rlrResult.Successful)
        Assert.IsNotNull(rlrResult.Error)
        myMain.Log.SIS.LogString(Level.Debug, "Error Message", rlrResult.Error.Message)
        Assert.IsNotNull(rlrResult.Start)
        Assert.IsNotNull(rlrResult.TimeToLogin)
    End Sub

End Class
