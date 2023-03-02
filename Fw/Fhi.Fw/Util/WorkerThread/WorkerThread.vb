Imports Fhi.Fw.Lang.Threading

Namespace Util.WorkerThread
    ''' <summary>
    ''' ワーカースレッドクラス
    ''' </summary>
    ''' <remarks>VBのThreadは継承できないのでFwのAThreadを継承</remarks>
    Friend Class WorkerThread : Inherits AThread
        Private ReadOnly aChannel As Channel
        Private isRunnable As Boolean = True

        Public Sub New(ByVal name As String, ByVal aChannel As Channel)
            MyBase.New(name)
            Me.aChannel = aChannel
        End Sub

        ''' <summary>
        ''' スレッドワーカーを開始する
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub Run()
            While isRunnable
                Dim task As ITask = aChannel.TakeTask()
                task.Execute()
            End While
        End Sub

        ''' <summary>
        ''' スレッドワーカーを停止する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub StopWorker()
            isRunnable = False
        End Sub
    End Class
End Namespace