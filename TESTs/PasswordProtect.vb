
Imports Microsoft.VisualStudio.TestTools.UnitTesting



<TestClass()>
Public Class Password_Tests
    Inherits cms_Test

#Region "Test Cases / Test Parameter (as Properties/Functions)"

#Region "Constructors"

    Public Sub New()
        MyBase.New

        '# INIT:
        myMain.Log.Clear()

    End Sub

#End Region

#End Region

    <TestMethod()>
    Public Sub checkPassword_Test()

        'PREPARE:
        myMain.Log.SIS.LogObject("AppPath: " & myMain.AppPath.FullName, myMain.AppPath)
        Dim dirFiles2Check As New IO.DirectoryInfo(myMain.AppPath.ToString.Replace("FileChecker\bin\Debug\", String.Empty) & "password-protected")
        myMain.Log.SIS.LogObject("dirFiles2Check: " & dirFiles2Check.FullName, dirFiles2Check)


        '# How to check?

        'for each file in dirFiles2Check
        For Each filItem As IO.FileInfo In dirFiles2Check.GetFiles()

            'myMain.Log.SIS.LogObject("found filItem: " & filItem.Name, filItem)

            ''check if it's a password protected file or not
            'Dim ofiItem As New Form1(myMain)

            'Dim bolStatus As Boolean? = ofiItem.checkFile(filItem)

            'If bolStatus.HasValue Then

            '    If filItem.Name.ToLower.Contains("safe") Then

            '        Assert.IsTrue(bolStatus)

            '    Else

            '        Assert.IsFalse(bolStatus)

            '    End If

            'Else

            '    '# Why is bol Status unknown?
            '    myMain.Log.SIS.LogString("filItem.Extension.ToLower", filItem.Extension.ToLower)
            '    Select Case filItem.Extension.ToLower
            '        Case ".png"
            '            'OK

            '        Case Else
            '            Assert.Fail(filItem.Name.ToLower & " has raised an error!")

            '    End Select


            'End If

        Next

        myMain.Log.SIS.LogColored(Color.LightGreen, "The Unit Test 'checkPassword' ends")

    End Sub

End Class
