Imports System.Threading
Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent
    ''' <summary>
    ''' スレッドの生成を抽象化したインターフェース
    ''' </summary>
    ''' <remarks>JavaのThreadFactoryインターフェース相当</remarks>
    Public Interface ThreadFactory
        Function NewThread(ByVal r As Runnable) As Thread
        Function NewThread(ByVal dlgtThreadStart As ThreadStart) As Thread
    End Interface
End Namespace