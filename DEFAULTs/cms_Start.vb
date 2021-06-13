


Public Module cms_StartModule

    <STAThread>
    Public Sub Main()
        Try
            'comment by salman and testing firefox work.
            'Dim ffd As FirefoxDownload = New FirefoxDownload(My)
            'Dim path As String = "D:\HandsomeSolution\Fiverr Client\saschaarlt\Firefox\Setup.msi"
            'ffd.DownloadFirefox(path)

            Dim myStart As New cms_Start()
            myStart.Start()


        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            'Log2File(AppPath(True) & myMain.ProductName() & ".err", "UPS", "Error on Main()", ex.Message.ToString, ex.InnerException.StackTrace.ToString, "cms_Start.vb:Main()")
        End Try
    End Sub

End Module





Public Class cms_Start

#Region "VARs"

    'MAIN (with INIT):
    Private myMain As Main

    'Global Thread Vars (http://stackoverflow.com/questions/17841377/stop-all-threads-if-an-error-is-detected-in-one-of-them):
    Private _ThreadStop As Long = 0L

    'Threads for Standard Usage:
    Private WithEvents appStart As APPStart, thrAPPStart As Threading.Thread
    'Private WithEvents logWindows As Form, thrLog As Threading.Thread
    Private thrLog As Threading.Thread

#End Region

#Region "Constructors"

    Public Sub New()
        Try
            '### for manually INIT cms_Main ###'
            myMain = New Main()


        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try
    End Sub

#End Region

#Region "Functions"

    Public Sub Start()
        ' Code zum Starten des Dienstes hier einfügen. Diese Methode sollte Vorgänge
        ' ausführen, damit der Dienst gestartet werden kann.
        Try

            ' ### START ### START ### START ### START ### START ### START ### START ###
            'Application.DoEvents()
            appStart = New APPStart(myMain)

            thrAPPStart = New Threading.Thread(AddressOf appStart.GO)
            thrAPPStart.Name = "APPStart"
            '# ERROR: Für den aktuellen Thread muss der STA-Modus (Single Thread Apartment) festgelegt werden, bevor OLE-Aufrufe ausgeführt werden können.
            'Wurde damit behoben:
            thrAPPStart.TrySetApartmentState(Threading.ApartmentState.STA)
            Dim retStart As returnValue = ThreadStart(thrAPPStart, False)
            If Not retStart.Success Then Throw retStart.Exception

            ' ### END ### END ### END ### END ### END ### END ### END ###


        Catch ex As Exception
            'myMain.Error.getException(ex, Me)
        End Try
    End Sub

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

#End Region

End Class