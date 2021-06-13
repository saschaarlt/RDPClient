
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Gurock.SmartInspect
Imports System
Imports System.Collections.Specialized
Imports System.IO
Imports System.Net
Imports System.Text



Public Class cms_Test
    Inherits cms_LastError

#Region "VARs"

    Public Shared myMain As Main

    'Public VARs for direct Access (Fields):


    'Private VARs (Access only via Properties):
    Private _ThrowFailedMessage As New Exception("The UnitTest has to be failed, but did NOT failed :-(. So we have a problem! Please check why the test is SUCCESSFUL!")

#End Region

#Region "Constructors"

    '# INITed by ALECS_JOB Queue or DeserializeObject
    Public Sub New()

        '# INITs a new Mai for Unit Tests with boldClear=True
        myMain = New Main()

        myMain.UnitTest = True

    End Sub

    Public Sub New(ByRef objMain As Main)
        myMain = objMain
        myMain.UnitTest = True
    End Sub

#End Region

#Region "Functions"

#Region " NEW's"

    Public Overloads Sub checkNEW(objME As Object)
        Assert.IsNotNull(objME, "Unable to create object!")

        '# check for Last error exception
        'checkLastError(objME)

        Assert.IsNotNull(myMain, "myMain could NOT initialised :-(")

        'muss vorher abgefragt werden ob es das überhaupt gibt...
        'Assert.IsFalse(objME.isLoaded(), "Please NOT Load objects on INIT. This slows down the performance :-(")

        myMain.Log.SIS.LogObject("checkNew (successful), object is " & objME.GetType.ToString, objME)
    End Sub

    Public Sub CheckUID(uidGUID As Guid?, Optional bolDefaultEmpty As Boolean = True)
        If bolDefaultEmpty Then
            Assert.IsTrue(myMain.HELPER.AnEmptyValue(uidGUID))
        Else
            Assert.IsTrue(myMain.HELPER.NotAnEmptyValue(uidGUID))
        End If
    End Sub

    '<Obsolete("Check Calling, CheckNew_WholeCMSDefaults4AXXObjects or CheckNew.")>
    Public Sub CheckActive(bolActive As Boolean?, Optional bolDefaultEmpty As Boolean = True)
        If bolDefaultEmpty Then
            Assert.IsFalse(bolActive.HasValue)
            Assert.IsNull(bolActive)
        Else
            Assert.IsTrue(bolActive.HasValue)
        End If
    End Sub

#End Region

    Public Sub checkReturnValue(retValue As returnValue, Optional objType As Type = Nothing)
        Assert.IsTrue(retValue.Success)
        Assert.IsNull(retValue.Exception, "Exception is NOT null.")
        Assert.IsNotNull(retValue.Value)

        If objType IsNot Nothing Then
            Assert.IsInstanceOfType(retValue.Value, objType)
        End If
    End Sub

#Region "Load / Save"

    Public Sub CheckLoad(objALECS As Object)

        Assert.IsTrue(objALECS.Load())

        If TypeOf objALECS Is IALECS_Entry.ILoadable Then

            Assert.IsTrue(objALECS.isLoaded())

        End If

        'Check defaults like CheckNew_WholeCMSDefaults4AXXObjects(objALECS, False, False)
        checkNEW(objALECS)
        CheckUID(objALECS.UID, False)
        'Assert.IsNull(objALECS.ID)
        CheckActive(objALECS.Active, False)

    End Sub

    Public Function CheckLoad_ALECSObject(objALECS As Object) As Boolean
        CheckLoad_ALECSObject = False

        If TypeOf objALECS Is IALECS_Entry.ILoadable Then
            'myMain.Log.add(objALECS.GetType.Name & " has a LoadAble interface :-)", Me)
            Assert.IsFalse(objALECS.isLoaded())
        End If

        Assert.IsTrue(objALECS.Load())

        If TypeOf objALECS Is IALECS_Entry.ILoadable Then

            Assert.IsTrue(objALECS.isLoaded())

        End If

        'Check defaults like CheckNew_WholeCMSDefaults4AXXObjects(objALECS, False, False)
        checkNEW(objALECS)
        CheckUID(objALECS.UID, False)
        'Assert.IsNull(objALECS.ID)
        CheckActive(objALECS.Active, False)

        CheckLoad_ALECSObject = True

    End Function

    Public Function CheckSave_ALECSObject(objALECS As Object) As Boolean
        CheckSave_ALECSObject = False

        Assert.IsTrue(objALECS.Save())

        'Check defaults like CheckNew_WholeCMSDefaults4AXXObjects(objALECS, False, False)
        checkNEW(objALECS)
        CheckUID(objALECS.UID, False)
        'Assert.IsNull(objALECS.ID)
        CheckActive(objALECS.Active, False)

        CheckSave_ALECSObject = True

    End Function

#End Region


#Region "Stateable: Activate / DEactivate"

    Public Sub CheckStateable_ALECSObject(objALECS As Object)

        '1 Load the object if it's possible:
        If TypeOf objALECS Is IALECS_Entry.ILoadable OrElse TypeOf objALECS Is IALECS_Entry.ILoadable Then
            Assert.IsTrue(objALECS.Load())
        End If

        '#2 Stateable object?
        If Not TypeOf objALECS Is IALECS_Entry.IStateable AndAlso Not TypeOf objALECS Is IALECS_Entry.IStateable Then
            Throw New EvaluateException("is not a IALECS_Entry.IStateable Base!")
        Else

            myMain.Log.add(objALECS.GetType.Name & " has a StateAble interface :-)", Me)
            Dim bolFirstRun As Boolean = True

ReRun:
            Select Case objALECS.Active
                Case True
                    ' is active try deactivate
                    Assert.IsTrue(objALECS.Deactivate)
                    Assert.IsFalse(objALECS.Active)
                    If TypeOf objALECS Is IALECS_Entry.ILoadable Then
                        Assert.IsTrue(objALECS.Load)    '# reload
                    End If
                    Assert.IsFalse(objALECS.Active)

                Case False
                    ' is NOT active try activate
                    Assert.IsTrue(objALECS.Activate)
                    Assert.IsTrue(objALECS.Active)
                    If TypeOf objALECS Is IALECS_Entry.ILoadable Then
                        Assert.IsTrue(objALECS.Load)    '# reload
                    End If
                    Assert.IsTrue(objALECS.Active)

                Case Else
                    'Nothing
                    Throw New NotImplementedException("")

            End Select

            ' and (back) reverse
            If bolFirstRun Then
                bolFirstRun = False
                GoTo ReRun
            End If

        End If
    End Sub

#End Region

#Region "Serialization / DEserialization"

    Public Sub CheckSerializeObject(objToTest As Object)
        Dim strSerialized As String = Newtonsoft.Json.JsonConvert.SerializeObject(objToTest)
        Assert.AreNotEqual("{}", strSerialized, "The serialized object is empty!")
        Assert.AreEqual("{}", Newtonsoft.Json.JsonConvert.SerializeObject(New Object()), "A serialized object doesn't equal '{}'!")
    End Sub

    Private Sub CheckSerializeObject4HangfireDashboard(obj2Test As Object)
        '# 1. serialize the object:
        Dim strSerialized As String = Newtonsoft.Json.JsonConvert.SerializeObject(obj2Test)
        Assert.AreNotEqual("{}", strSerialized, "The serialized object is empty!")
        ' Finish: this is important that Dashboard can serialized jobs to enqueue it

        '# 2. DEserialize it back:
        Dim objDeserialized As Object = Newtonsoft.Json.JsonConvert.DeserializeObject(strSerialized, obj2Test.GetType())

        'myMain.Log.SIS.LogObject("1 " & obj2Test.GetType.ToString, obj2Test)
        'myMain.Log.SIS.LogObject("2 " & objDeserialized.GetType.ToString, objDeserialized)
        If obj2Test.Equals(objDeserialized) Then
            'myMain.Log.add("HORRAY! obj2Test and objDeserialized (above) are both equal.", Me, Level.Debug, Color.Lime)
            'myMain.Log.logAssert("BOTH objects are EQUAL!")
        Else
            myMain.Log.logJson("This is the (serialized) JSON String from (1) original object. Is anaything in there?", strSerialized)
            myMain.Log.SIS.LogObject("1 " & obj2Test.GetType.ToString, obj2Test)
            myMain.Log.SIS.LogObject("2 " & objDeserialized.GetType.ToString, objDeserialized)
            myMain.Log.add("OH SHIT! obj2Test and objDeserialized (above) are differents each other.", Me, Level.Debug, Color.Red)
            myMain.Log.logAssert("BOTH objects are DIFFERENT!")
        End If
        Assert.AreEqual(obj2Test, objDeserialized, "The deserialized object doesn't equal the original object!")

        Assert.AreEqual("{}", Newtonsoft.Json.JsonConvert.SerializeObject(New Object()), "A serialized object doesn't equal '{}'!")

    End Sub

    Public Sub CheckSerializationComplete(obj2Test As Object)

        '4 Dashboard?
        CheckSerializeObject4HangfireDashboard(obj2Test)

        'REPEAT:
        'Dim strSerialized As String = Newtonsoft.Json.JsonConvert.SerializeObject(obj2Test)
        'Assert.AreNotEqual("{}", strSerialized, "The serialized object is empty!")

        'Assert.AreEqual("{}", Newtonsoft.Json.JsonConvert.SerializeObject(New Object()), "A serialized object doesn't equal '{}'!")

    End Sub

#End Region

#Region "Object comparables"

    Public Function compareObjects(objSource As Object, objDestination As Object) As Boolean
        compareObjects = False
        clearLastError()
        Try
            '# Source Fields:
            For Each fldSource As System.Reflection.FieldInfo In objSource.GetType.GetFields()
                Dim typSource As Type = objSource.[GetType], typDestination As Type = objDestination.[GetType]
                '# Try Property
                Dim prpDestination As System.Reflection.PropertyInfo = typDestination.GetProperty(fldSource.Name)
                If Not prpDestination Is Nothing Then
                    myMain.Log.SIS.LogObject("!!!!!!!!! ...compare Property: " & prpDestination.Name, prpDestination)
                    Assert.AreEqual(typSource.GetProperty(fldSource.Name), prpDestination)
                    Throw New NotImplementedException("")
                Else
                    '# Try Field:
                    Dim fldDestination As System.Reflection.FieldInfo = typDestination.GetField(fldSource.Name)
                    If Not fldDestination Is Nothing Then
                        myMain.Log.SIS.LogObject("...compare Field: " & fldDestination.Name, fldDestination)
                        Assert.AreEqual(typSource.GetField(fldSource.Name), fldDestination)
                        Assert.AreEqual(fldSource.GetValue(objSource), fldDestination.GetValue(objDestination))
                    Else
                        'myMain.Log.add("Destination has NO Property, NO Field: " & fldSource.Name, , Me, Color.Orange)
                        Throw New Exception("Destination has NO Property, NO Field: " & fldSource.Name)
                    End If
                End If
            Next

            '# Source Properties:
            For Each prpSource As System.Reflection.PropertyInfo In objSource.GetType().GetProperties()
                If prpSource.CanRead Then
                    Dim typSource As Type = objSource.[GetType], typDestination As Type = objDestination.[GetType]
                    '# Try Property
                    Dim prpDestination As System.Reflection.PropertyInfo = typDestination.GetProperty(prpSource.Name)
                    If Not prpDestination Is Nothing Then
                        myMain.Log.SIS.LogObject("...compare Property: " & prpDestination.Name, prpDestination)
                        Assert.AreEqual(typSource.GetProperty(prpSource.Name), prpDestination)
                        Assert.AreEqual(prpSource.GetValue(objSource, Nothing), prpDestination.GetValue(objDestination, Nothing))
                    Else
                        '# Try Field:
                        Dim fldDestination As System.Reflection.FieldInfo = typDestination.GetField(prpSource.Name)
                        If Not fldDestination Is Nothing Then
                            myMain.Log.SIS.LogObject("!!!!!!! ...compare Field: " & fldDestination.Name, fldDestination)
                            Assert.AreEqual(typSource.GetField(prpSource.Name), fldDestination)
                            Throw New NotImplementedException("")
                        Else
                            'myMain.Log.add("Destination has NO Property, NO Field: " & prpSource.Name, , Me, Color.Orange)
                            Throw New Exception("Destination has NO Property, NO Field: " & prpSource.Name)
                        End If
                    End If
                End If
            Next
            compareObjects = True


        Catch exAFE As AssertFailedException
            myMain.Error.getException(exAFE, Me)
            setLastError(exAFE)
            Return False
        Catch ex As Exception
            myMain.Error.getException(ex, Me)
            setLastError(ex)
            Return False
        End Try
    End Function

#End Region

#Region "Runtimes"

    Public Sub checkRuntimeCache(timWithoutCache As Timing, timWithCache As Timing, Optional tspAccelerated As TimeSpan = Nothing)

        myMain.Log.logRuntimeSI(timWithoutCache, "Runtime withOUT Cache")
        myMain.Log.logRuntimeSI(timWithCache, "Runtime WITH Cache")

        '# Run without Cache is longer that second one (with Cache)
        Assert.IsTrue(timWithoutCache.Result.ElapsedTicks > timWithCache.Result.ElapsedTicks, "The runtime for running with cache is faster than without :-o")

        '# check if runtime1 is greater then 
        If IsNothing(tspAccelerated) Then tspAccelerated = TimeSpan.FromMilliseconds(1)
        Assert.IsTrue((timWithoutCache.Result.ElapsedMilliseconds - timWithCache.Result.ElapsedMilliseconds) >= tspAccelerated.TotalMilliseconds, "The runtime for running with cache is NOT so Accelerated then we definine :-(")

    End Sub

#End Region



    Public Sub ThrowFailed()
        Throw _ThrowFailedMessage
    End Sub

    Public Sub logSuccessfulCheckEnd(strCheckName As String, Optional strDetails As String = "erfolgreich beendet")
        With myMain.Log.SIS
            .Color = Color.LightGreen
            .AddCheckpoint(Level.Message, "UnitTest", strCheckName & If(strDetails IsNot Nothing, ": " & strDetails, ""))
            .ResetColor()
            .LogSeparator(Level.Message)
        End With
    End Sub

#End Region

End Class
