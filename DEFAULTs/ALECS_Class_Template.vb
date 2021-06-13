
Imports System.Runtime.Serialization



Public Interface IALECS_Entry



    '### VARS # VARS # VARS # VARS # VARS # VARS # VARS # VARS # VARS # VARS # VARS # VARS # VARS ###'

    '<NonSerialized>
    'Property myMain As Main

    <DataMember(Order:=1)>
    Property UID As Guid?

    <DataMember(EmitDefaultValue:=False, Order:=2)>
    Property ID As String

    <DataMember(EmitDefaultValue:=False, Order:=3)>
    ReadOnly Property Active As Boolean?



    <DataMember(EmitDefaultValue:=False, Order:=10)>
    Property Name As String



    ReadOnly Property Created As DateTime?
    Sub setCreated(dteDateTime As DateTime)

    ReadOnly Property changed As DateTime?


    '### Debug & Display # Debug & Display # Debug & Display # Debug & Display # Debug & Display ###'

    Function getDisplayName() As String

    Sub DebugMe(Optional strPREfix As String = "")  ', Optional bolForce As Boolean = False



    '### Main-Stuff # Main-Stuff # Main-Stuff # Main-Stuff # Main-Stuff # Main-Stuff # Main-Stuff ###'

    Function Clone() As Object
    'PRIMED? Function Clone(Of T)() As T

    'Function Equals(obj As Object) As Boolean

    'not yet:  Function GetHashCode() As Integer

    Function getLastError() As Exception



    '### LOAD # LOAD # LOAD # LOAD # LOAD # LOAD # LOAD # LOAD # LOAD # LOAD # LOAD # LOAD # LOAD ###'

    Public Interface ILoadable
        Inherits IALECS_Entry

        Function BaseLoad() As DataRow

        Function Load() As Boolean

        Function getSelection2(Of T)(Optional bolLoadObject As Boolean = False, Optional lstCountdown As List(Of String) = Nothing) As Dictionary(Of Guid, T)

        'Function getSelection() As Dictionary(Of Guid, Object)

        Function isLoaded() As Boolean

        Sub setLoaded(Optional bolLoaded As Boolean = True)

        Function CheckIfExist(Optional bolUseCache As Boolean = True) As Boolean

    End Interface



    '### SAVE # SAVE # SAVE # SAVE # SAVE # SAVE # SAVE # SAVE # SAVE # SAVE # SAVE # SAVE # SAVE ###'

    Public Interface ISaveable

        Function BaseSave() As Boolean

        Function Save() As Boolean

    End Interface



    '### STATEs # STATEs # STATEs # STATEs # STATEs # STATEs # STATEs # STATEs # STATEs # STATEs # STATEs ###'

    Public Interface IStateable

        '<DataMember(EmitDefaultValue:=False, Order:=3)>
        'IS already in IALECS implemented ReadOnly Property Active As Boolean?

        Function Activate() As Boolean

        Function Deactivate() As Boolean

    End Interface



    '### CREATE # DELETE # CREATE # DELETE # CREATE # DELETE # CREATE # DELETE # CREATE # DELETE # CREATE # DELETE ###'

    Public Interface ICreateable

        Function Create() As Boolean

    End Interface

    Public Interface IDeletable

        Function Delete() As Boolean

    End Interface





    '### CACHE # CACHE # CACHE # CACHE # CACHE # CACHE # CACHE # CACHE # CACHE # CACHE # CACHE # CACHE # CACHE ###'

    Public Interface ICache
        Inherits IALECS_Entry

        Function LoadFromCache() As Boolean

        'Function isLoaded() As Boolean

        'Sub setLoaded(Optional bolLoaded As Boolean = True)

        'Function CheckIfExist(Optional bolUseCache As Boolean = True) As Boolean

    End Interface

End Interface



<Serializable()> <DataContract()>
Public MustInherit Class ALECS_Template_Entry_FAKE
    Inherits cms_LastError
    Implements IALECS_Entry, IALECS_Entry.ILoadable, IALECS_Entry.ISaveable, IALECS_Entry.IStateable, IALECS_Entry.ICreateable, IALECS_Entry.IDeletable

#Region "VARs"

    'Public VARs for direct Access (Fields):
    <NonSerialized>             ' for Caches... please forgot when store!
    Friend myMain As Main

    <DataMember(Order:=1)>
    Public Property UID As Guid? = Nothing Implements IALECS_Entry.UID
    <DataMember(EmitDefaultValue:=False, Order:=2)>
    Public Property ID As String Implements IALECS_Entry.ID
    '<DataMember(EmitDefaultValue:=False, Order:=3)>
    'Public ReadOnly Property Active As Boolean? = Nothing Implements IALECS_Entry.IStateable.Active

    Public Property Active As Boolean? = Nothing Implements IALECS_Entry.Active

    <DataMember(Order:=11)>
    Public Property Name As String = Nothing Implements IALECS_Entry.Name

    Public Value As String = Nothing

    Public ReadOnly Property Created As DateTime? Implements IALECS_Entry.Created
    Public ReadOnly Property changed As DateTime? Implements IALECS_Entry.changed

    Public NotImplementedReturn As String = "This is the return of a FAKE class (ALECS_Template_Entry_FAKE). It will return, because this FAKE class was implemeted to reduce the code / complexness or for SECURITY reasons."
    ''Public Tenant As ALECS_Tenant_Entry = Nothing

    '<DataMember(EmitDefaultValue:=False, Order:=89)>
    'Public Comment As String = Nothing


    'Private VARs (Access only via Properties):
    Private _ChildObjectName As String = Nothing ', _ObjectChild As Object = Nothing

#End Region

#Region "Properties"

    Public ReadOnly Property ChildObjectName As String
        Get
            Return _ChildObjectName
        End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
        Throw New NotImplementedException(NotImplementedReturn)
    End Sub

    Public Sub New(ByRef objMain As Main)
        myMain = objMain
        Throw New NotImplementedException(NotImplementedReturn)
    End Sub

    '# INITed by ALECS_JOB Queue or DeserializeObject (of Childs)
    Public Sub New(objType As Type)
        myMain = New Main()

        _ChildObjectName = objType.Name.ToLower
        '_ChildType = objType
    End Sub

    Public Sub New(ByRef objMain As Main, objType As Type, Optional objUID As Guid? = Nothing)
        myMain = objMain
        _ChildObjectName = objType.Name.ToLower
        '_ChildType = objType
        UID = objUID
    End Sub

    Public Sub New(ByRef objMain As Main, bolNoError As Boolean)
        myMain = objMain
        If Not bolNoError Then Throw New NotImplementedException(NotImplementedReturn)
    End Sub

#End Region

#Region "IALECS_Entry Functions"

    ''' <summary>
    '''     THIS is the BASE.Load()
    ''' </summary>
    Public Function BaseLoad() As DataRow Implements IALECS_Entry.ILoadable.BaseLoad
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Overridable Function Load() As Boolean Implements IALECS_Entry.ILoadable.Load
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Function isLoaded() As Boolean Implements IALECS_Entry.ILoadable.isLoaded
        Return True
    End Function

    Public Sub setLoaded(Optional bolLoaded As Boolean = True) Implements IALECS_Entry.ILoadable.setLoaded
        Throw New NotImplementedException(NotImplementedReturn)
    End Sub

    ''' <summary>
    '''     THIS is the BASE.Save()
    ''' </summary>
    Public Function BaseSave() As Boolean Implements IALECS_Entry.ISaveable.BaseSave
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Overridable Function Save() As Boolean Implements IALECS_Entry.ISaveable.Save
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Function Activate() As Boolean Implements IALECS_Entry.IStateable.Activate
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Function Deactivate() As Boolean Implements IALECS_Entry.IStateable.Deactivate
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Overridable Function Create() As Boolean Implements IALECS_Entry.ICreateable.Create
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    ''' <summary>
    ''' Delete's the record from database and initialize this object new (UID is also thrown away!)
    ''' </summary>
    Public Overridable Function Delete() As Boolean Implements IALECS_Entry.IDeletable.Delete
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    ''' <summary>
    ''' this is the BASE Selection for ALECS object's which based on DB's (also needed one for Cache, etc.)
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    Public Function getSelection2(Of T)(Optional bolLoadObject As Boolean = False, Optional lstCountdown As List(Of String) = Nothing) As Dictionary(Of Guid, T) Implements IALECS_Entry.ILoadable.getSelection2
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Function getSelection() As Dictionary(Of Guid, Object) 'Implements IALECS_Entry.ILoadable.getSelection
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    'Public Function Clone() Implements IALECS_Entry.Clone
    '    Throw New NotImplementedException(NotImplementedReturn)
    'End Function

    Public Overridable Function Clone() As Object Implements IALECS_Entry.Clone
        Return Me.MemberwiseClone
    End Function

    Public Overridable Function getDisplayName() As String Implements IALECS_Entry.getDisplayName
        Return String.Empty
    End Function

    Public Overridable Sub DebugMe(Optional strPREfix As String = "") Implements IALECS_Entry.DebugMe
        Throw New NotImplementedException(NotImplementedReturn)
    End Sub

    ''' <summary>
    ''' check in database if record with UID already exists (BASE functionality for all childs)
    ''' </summary>
    Public Overridable Function CheckIfExist(Optional bolUseCache As Boolean = True) As Boolean Implements IALECS_Entry.ILoadable.CheckIfExist
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Overridable Function getLastError() As Exception Implements IALECS_Entry.getLastError
        Throw New NotImplementedException(NotImplementedReturn)
    End Function

    Public Sub setCreated(dteDateTime As DateTime) Implements IALECS_Entry.setCreated
        Throw New NotImplementedException(NotImplementedReturn)
    End Sub

#End Region

End Class
