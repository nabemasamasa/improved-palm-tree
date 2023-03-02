Imports Fhi.Fw.Util

Namespace Util
    Public Class SynchronizedIndexedListTest : Inherits IndexedListTest

        Protected Overrides Function NewList(Of T)() As IIndexedList(Of T)
            Return New SynchronizedIndexedList(Of T)
        End Function

        Protected Overrides Function NewList(Of T)(ByVal isCreateGeneric As Boolean) As IIndexedList(Of T)
            Return New SynchronizedIndexedList(Of T)(isCreateGeneric)
        End Function

    End Class
End Namespace