''' <summary>
''' システムの標準日時クラス
''' </summary>
''' <remarks></remarks>
Public Class SystemDate
    Private _currentDateTime As Nullable(Of DateTime)

    Public Shared DefaultBehavior As IBehavior
    Private ReadOnly behavior As IBehavior

    Public Interface IBehavior
        Function GetDateTimeNow() As DateTime?
    End Interface

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Me.New(DefaultBehavior)
    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="behavior"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal behavior As IBehavior)
        Me.behavior = behavior
        If behavior Is Nothing Then
            Throw New SystemException("FwSetup.Initialize()が必要")
        End If
    End Sub

    ''' <summary>
    ''' 現在日時を返す
    ''' </summary>
    ''' <returns>現在日時</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentDateTime() As DateTime
        Get
            If Not _currentDateTime.HasValue Then
                _currentDateTime = behavior.GetDateTimeNow
            End If
            Return _currentDateTime.Value
        End Get
    End Property

    ''' <summary>
    ''' 現在日付をDB書式にして返す
    ''' </summary>
    ''' <returns>現在日付(DB書式)</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentDateDbFormat() As String
        Get
            Return DateUtil.ConvDateToHyphenYyyymmdd(CurrentDateTime)
        End Get
    End Property

    ''' <summary>
    ''' 現在日付をYYYYMMDD書式にして返す
    ''' </summary>
    ''' <returns>現在日付(YYYYMMDD書式)</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentDateAsInteger() As Integer
        Get
            Return DateUtil.ConvDateToInteger(CurrentDateTime).Value
        End Get
    End Property

    ''' <summary>
    ''' 現在時刻をDB書式にして返す
    ''' </summary>
    ''' <returns>現在時刻(DB書式)</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentTimeDbFormat() As String
        Get
            Return DateUtil.ConvTimeToColonHhmmss(CurrentDateTime)
        End Get
    End Property

    ''' <summary>
    ''' 現在時刻をHHMMSS書式にして返す
    ''' </summary>
    ''' <returns>現在時刻(HHMMSS書式)</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentTimeAsInteger() As Integer
        Get
            Return DateUtil.ConvTimeToInteger(CurrentDateTime).Value
        End Get
    End Property

    ''' <summary>
    ''' 標準日時をクリアし、再度取得することを促す
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()
        _currentDateTime = Nothing
    End Sub
End Class
