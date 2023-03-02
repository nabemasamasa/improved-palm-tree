Namespace Util.WorkerThread
    ''' <summary>
    ''' Executeを保証するタスクインターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface ITask
        ''' <summary>
        ''' タスクを実行する
        ''' </summary>
        ''' <remarks></remarks>
        Sub Execute()
    End Interface
End Namespace