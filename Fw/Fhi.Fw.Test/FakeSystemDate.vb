''' <summary>
''' 偽装システム日時
''' </summary>
''' <remarks></remarks>
Public Class FakeSystemDate : Inherits SystemDate

    Private ReadOnly aBehavior As Behavior

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="dateTimeString">偽装日時を表す文字列</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dateTimeString As String)
        Me.New(New Behavior(dateTimeString))
    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="aDateTime">偽装日時</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal aDateTime As DateTime)
        Me.New(New Behavior(aDateTime))
    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="behavior">偽装振る舞い実装</param>
    ''' <remarks></remarks>
    Private Sub New(ByVal behavior As Behavior)
        MyBase.New(behavior)
        Me.aBehavior = behavior
    End Sub

    ''' <summary>
    ''' 偽装日時を設定する
    ''' </summary>
    ''' <param name="dateTimeString">偽装日時を表す文字列</param>
    ''' <remarks></remarks>
    Public Sub SetFakeDateTime(ByVal dateTimeString As String)
        SetFakeDateTime(CDate(dateTimeString))
    End Sub

    ''' <summary>
    ''' 偽装日時を設定する
    ''' </summary>
    ''' <param name="aDateTime">偽装日時</param>
    ''' <remarks></remarks>
    Public Sub SetFakeDateTime(ByVal aDateTime As DateTime)
        aBehavior.FakeDateTime = aDateTime
    End Sub

    Public Class Behavior : Implements SystemDate.IBehavior
        Public FakeDateTime As DateTime

        Public Sub New(ByVal dateTimeString As String)
            Me.New(CDate(dateTimeString))
        End Sub

        Public Sub New(ByVal aDateTime As DateTime)
            FakeDateTime = aDateTime
        End Sub

        Public Function GetDateTimeNow() As Date? Implements IBehavior.GetDateTimeNow
            Return FakeDateTime
        End Function
    End Class
End Class
