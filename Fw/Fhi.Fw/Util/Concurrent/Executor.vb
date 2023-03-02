Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent

    ''' <summary>
    ''' スレッドの実行を抽象化したインターフェース
    ''' </summary>
    ''' <remarks>javaのExecutorインターフェース相当</remarks>
    Public Interface Executor

        ''' <summary>
        ''' 実行する
        ''' </summary>
        ''' <param name="command">実行する内容</param>
        ''' <remarks></remarks>
        Sub Execute(ByVal command As Runnable)
    End Interface
End Namespace