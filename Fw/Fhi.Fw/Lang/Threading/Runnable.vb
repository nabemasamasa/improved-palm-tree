Namespace Lang.Threading
    ''' <summary>
    ''' スレッドが実行する「仕事」を表すインターフェース
    ''' </summary>
    ''' <remarks>JavaのRunnableインターフェース相当</remarks>
    Public Interface Runnable
        Sub Run()
    End Interface
End Namespace