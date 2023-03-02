Imports NUnit.Framework
Imports Fhi.Fw.Util
Imports System.Text

Namespace Util
    ''' <summary>
    ''' 当テストクラスに限っては、ここに IIndexedListインターフェースのテストメソッド記述し、
    ''' 実装クラスのテストクラスで継承して使用する
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class IndexedListTest

        Protected MustOverride Function NewList(Of T)() As IIndexedList(Of T)

        Protected MustOverride Function NewList(Of T)(ByVal isCreateGeneric As Boolean) As IIndexedList(Of T)

#Region "本体のtest"
        Public Class MainTest : Inherits IndexedListTest

            Protected Overrides Function NewList(Of T)() As IIndexedList(Of t)
                Return New IndexedList(Of T)
            End Function

            Protected Overrides Function NewList(Of T)(ByVal isCreateGeneric As Boolean) As IIndexedList(Of t)
                Return New IndexedList(Of T)(isCreateGeneric)
            End Function

            Private Class Hoge
                Public Fuga As String
            End Class

            Private Interface MyInterface
                Sub hoge()
            End Interface

            Private Class MyInterfaceImpl : Implements MyInterface
                Public Sub hoge() Implements MyInterface.hoge
                End Sub
            End Class

            <Test()> Public Sub IsCreateGeneric_false_でもInteger型は初期値ゼロ()
                Dim list As New IndexedList(Of Integer)
                Assert.AreEqual(0, list(2))
            End Sub

            <Test()> Public Sub IsCreateGeneric_false_でNullableInteger型は初期値null()
                Dim list As New IndexedList(Of Integer?)
                Assert.IsNull(list(2))
            End Sub

            <Test()> Public Sub IsCreateGeneric_false_で通常のクラス型は初期値null()
                Dim list As New IndexedList(Of Hoge)
                Assert.IsNull(list(2))
            End Sub

            <Test()> Public Sub 通常のクラスは引数なしコンストラクタで作成したインスタンス()
                Dim list As New IndexedList(Of Hoge)(True)
                Assert.IsNotNull(list(1))
                Assert.IsNull(list(1).Fuga)
            End Sub

            <Test()> Public Sub Interface型の救済にインスタンス化するタイプを指定させる()
                Dim list As New IndexedList(Of MyInterface)(True, GetType(MyInterfaceImpl))
                Assert.IsTrue(TypeOf list(1) Is MyInterfaceImpl)
            End Sub

            <Test()> Public Sub Constructor_別コレクションから生成()
                Dim params As New List(Of Integer)(New Integer() {3, 1, 2})
                Dim list As New IndexedList(Of Integer)(params)

                Assert.AreEqual(3, list.Indexes.Count)
                Assert.AreEqual(3, list(0))
                Assert.AreEqual(1, list(1))
                Assert.AreEqual(2, list(2))
            End Sub

            <Test(), ExpectedException(GetType(NotSupportedException))> Public Sub インスタンス生成でString型は未対応()
                Dim list As New IndexedList(Of String)(True)
                Assert.IsNull(list(1))
            End Sub

        End Class
#End Region

        <Test()> Public Sub HasValue_値のあるなし()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"

            Assert.IsFalse(aList.HasValue(0))
            Assert.IsFalse(aList.HasValue(1))
            Assert.IsFalse(aList.HasValue(2))
            Assert.IsTrue(aList.HasValue(3))
        End Sub

        <Test()> Public Sub Item_値をセットして取得して()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(4) = "four"

            Assert.AreEqual("three", aList.Item(3))
            Assert.AreEqual("four", aList.Item(4))
        End Sub

        <Test()> Public Sub Items_順不同で値一覧を返す()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(4) = "four"

            Dim actuals As ICollection(Of String) = aList.Items
            Assert.AreEqual(2, actuals.Count)
            For Each actual As String In actuals
                If "three".Equals(actual) Then
                    Assert.IsTrue(True)
                ElseIf "four".Equals(actual) Then
                    Assert.IsTrue(True)
                Else
                    Assert.Fail()
                End If
            Next
        End Sub

        <Test()> Public Sub Indexes_順不同でindex一覧を返す()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(4) = "four"

            Dim actuals As ICollection(Of Integer) = aList.Indexes
            Assert.AreEqual(2, actuals.Count)
            For Each actual As Integer In actuals
                If actual = 3 Then
                    Assert.IsTrue(True)
                ElseIf actual = 4 Then
                    Assert.IsTrue(True)
                Else
                    Assert.Fail()
                End If
            Next
        End Sub

        <Test()> Public Sub GetMaxIndex_インスタンス生成直後はnull()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            Assert.IsNull(aList.GetMaxIndex)
        End Sub

        <Test()> Public Sub GetMaxIndex_最大indexを返す()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(4) = "four"

            Assert.IsNotNull(aList.GetMaxIndex)
            Assert.AreEqual(4, aList.GetMaxIndex)
        End Sub

        <Test()> Public Sub Add_ゼロindexから値を追加()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Add("one")
            aList.Add("two")

            Assert.AreEqual("one", aList.Item(0))
            Assert.AreEqual("two", aList.Item(1))
        End Sub

        <Test()> Public Sub AddRange_コレクションを追加()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.AddRange(New String() {"1", "2", "3"})

            Assert.AreEqual("1", aList.Item(0))
            Assert.AreEqual("2", aList.Item(1))
            Assert.AreEqual("3", aList.Item(2))
        End Sub

        <Test()> Public Sub Clear_要素をクリアする()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.AddRange(New String() {"1", "2", "3"})

            aList.Clear()

            Assert.IsNull(aList.GetMaxIndex)
            Assert.AreEqual(0, aList.Items.Count)
            Assert.AreEqual(0, aList.Indexes.Count)
        End Sub

        <Test()> Public Sub InsertAt_空情報を挿入()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(4) = "four"

            aList.InsertAt(4)

            Assert.AreEqual("three", aList.Item(3))
            Assert.AreEqual("four", aList.Item(5))
        End Sub

        <Test()> Public Sub InsertAt_空情報を挿入するから挿入したindexは値無し()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(4) = "four"

            aList.InsertAt(4)

            Assert.IsFalse(aList.HasValue(4))
        End Sub

        <Test()> Public Sub InsertAt_複数挿入もできる()
            Dim aList As IIndexedList(Of Integer) = NewList(Of Integer)()

            aList.Item(0) = 10
            aList.Item(1) = 20

            aList.InsertAt(1, 4)

            Assert.AreEqual(10, aList(0))
            Assert.IsFalse(aList.HasValue(1))
            Assert.IsFalse(aList.HasValue(2))
            Assert.IsFalse(aList.HasValue(3))
            Assert.IsFalse(aList.HasValue(4))
            Assert.AreEqual(20, aList(5))
        End Sub

        <Test()> Public Sub RemoveAt_値を除去してそれ以降の要素が前詰めになる()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(4) = "four"
            aList.Item(5) = "five"

            aList.RemoveAt(4)

            Assert.AreEqual("three", aList.Item(3))
            Assert.AreEqual("five", aList.Item(4))
        End Sub

        <Test()> Public Sub RemoveAt_複数要素の除去もできる()
            Dim aList As IIndexedList(Of Integer) = NewList(Of Integer)()

            aList.Item(0) = 10
            aList.Item(1) = 20
            aList.Item(2) = 30
            aList.Item(3) = 40
            aList.Item(4) = 50
            aList.Item(5) = 60

            aList.RemoveAt(1, 4)

            Assert.AreEqual(10, aList(0))
            Assert.AreEqual(60, aList(1))
        End Sub

        <Test()> Public Sub ClearAt_値を除去するけど_後続の情報は前詰めしない()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(0) = "a"
            aList(1) = "b"
            aList(2) = "c"

            aList.ClearAt(1)

            Assert.AreEqual("a", aList(0))
            Assert.AreEqual("c", aList(2), "クリアされた後続の情報は、そのまま")
            Assert.AreEqual(2, aList.GetMaxIndex)
            Assert.AreEqual(2, aList.Indexes.Count, "index={0,2}のみだから")
        End Sub

        <Test()> Public Sub Extract_値を抜き出す()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(5) = "five"

            Dim actuals As String() = aList.Extract(2, 4)

            Assert.AreEqual(4, actuals.Length)
            Assert.AreEqual("three", actuals(1))
            Assert.AreEqual("five", actuals(3))
        End Sub

        <Test()> Public Sub Extract_値を抜き出す_値が無ければnullが抜き出される()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(5) = "five"

            Dim actuals As String() = aList.Extract(2, 5)

            Assert.AreEqual(5, actuals.Length)
            Assert.IsNull(actuals(0))
            Assert.IsNotNull(actuals(1))
            Assert.IsNull(actuals(2))
            Assert.IsNotNull(actuals(3))
            Assert.IsNull(actuals(4))
        End Sub

        <Test()> Public Sub Supersede_値を差し替える_但しnull値では差し替えしない()

            Dim aList As IIndexedList(Of StringBuilder) = NewList(Of StringBuilder)(True)
            aList.Item(3).Append("three")
            aList.Item(5).Append("five")

            aList.Supersede(2, New StringBuilder() {New StringBuilder("one"), _
                                                    Nothing, _
                                                    New StringBuilder("three"), _
                                                    New StringBuilder("four")})

            Assert.AreEqual(5, aList.GetMaxIndex)
            Assert.AreEqual("one", aList.Item(2).ToString)
            Assert.AreEqual("three", aList.Item(3).ToString)
            Assert.AreEqual("three", aList.Item(4).ToString)
            Assert.AreEqual("four", aList.Item(5).ToString)
        End Sub

        <Test()> Public Sub Extract_Supersede_空行は空行のままSupersede()
            Dim testee As IIndexedList(Of StringBuilder) = NewList(Of StringBuilder)(True)
            testee(1).Append("a")
            testee(2).Append("b")
            testee(4).Append("c")
            testee(6).Append("d")

            Dim extract As StringBuilder() = testee.Extract(2, 4)

            testee.RemoveAt(2, 4)
            testee.InsertAt(2, 4)

            testee.Supersede(2, extract)

            Assert.IsTrue(testee.Indexes.Contains(1))
            Assert.IsTrue(testee.Indexes.Contains(2))
            Assert.IsTrue(testee.Indexes.Contains(4))
            Assert.IsTrue(testee.Indexes.Contains(6))
            Assert.IsFalse(testee.Indexes.Contains(0))
            Assert.IsFalse(testee.Indexes.Contains(3))
            Assert.IsFalse(testee.Indexes.Contains(5))

            Assert.AreEqual("a", testee(1).ToString)
            Assert.AreEqual("b", testee(2).ToString)
            Assert.AreEqual("c", testee(4).ToString)
            Assert.AreEqual("d", testee(6).ToString)
        End Sub

        <Test()> Public Sub IndexOf_一致する値のindexを返す()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(5) = "five"

            Assert.AreEqual(3, aList.IndexOf("three"))
            Assert.AreEqual(5, aList.IndexOf("five"))
        End Sub

        <Test()> Public Sub IndexOf_値が一致しなければマイナス1()

            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(5) = "five"

            Assert.AreEqual(-1, aList.IndexOf("four"))
        End Sub

        <Test()> Public Sub GetEnumrator_値をforeach()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList.Item(3) = "three"
            aList.Item(9) = "nine"

            Dim counter As Integer = 0
            For Each actual As String In aList
                If "three".Equals(actual) Then
                    counter += 1
                    Assert.IsTrue(True)
                ElseIf "nine".Equals(actual) Then
                    counter += 1
                    Assert.IsTrue(True)
                Else
                    Assert.Fail()
                End If
            Next
            Assert.AreEqual(2, counter)
        End Sub

        <Test()> Public Sub Constructor_CreateGenericなので_未設定の位置indexでもインスタンスを返す()
            Dim aList As IIndexedList(Of StringBuilder) = NewList(Of StringBuilder)(True)
            Assert.IsNotNull(aList.Item(1))
        End Sub

        <Test()> Public Sub Constructor_CreateGenericなので_未設定の位置indexでもインスタンスを返す_が負数はエラー()
            Dim aList As IIndexedList(Of StringBuilder) = NewList(Of StringBuilder)(True)
            Try
                Dim hoge = aList.Item(-1)

            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub SuppressGap_歯抜けのindexを詰める()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(0) = "a"
            aList(2) = "c"

            aList.SuppressGap()

            Assert.AreEqual("a", aList(0))
            Assert.AreEqual("c", aList(1), "index=1に詰まる")
            Assert.AreEqual(1, aList.GetMaxIndex)
        End Sub

        <Test()> Public Sub SuppressGap_Nothing値も詰める()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(0) = "a"
            aList(1) = Nothing
            aList(2) = "c"

            aList.SuppressGap()

            Assert.AreEqual("a", aList(0))
            Assert.AreEqual("c", aList(1), "index=1に詰まる")
            Assert.AreEqual(1, aList.GetMaxIndex)
        End Sub

        <Test()> Public Sub SuppressGap_プリミティブ型のNothing値は詰まらない()
            Dim aList As IIndexedList(Of Integer) = NewList(Of Integer)()
            aList(0) = 3
            aList(1) = Nothing
            aList(2) = 4

            aList.SuppressGap()

            Assert.AreEqual(3, aList(0))
            Assert.AreEqual(0, aList(1), "プリミティブ型だから、Nothing=0だから、詰まらない")
            Assert.AreEqual(4, aList(2))
            Assert.AreEqual(2, aList.GetMaxIndex)
        End Sub

        <Test()> Public Sub SuppressGap_末尾のNothing値も詰める()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(0) = "a"
            aList(1) = Nothing
            aList(2) = "c"
            aList(3) = Nothing

            aList.SuppressGap()

            Assert.AreEqual("a", aList(0))
            Assert.AreEqual("c", aList(1), "index=1に詰まる")
            Assert.AreEqual(1, aList.GetMaxIndex, "末尾のindexも詰まる")
        End Sub

        <Test()> Public Sub TrimEnd_末尾のNothing値を詰める()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(0) = "a"
            aList(1) = Nothing

            aList.TrimEnd()

            Assert.AreEqual("a", aList(0))
            Assert.AreEqual(0, aList.GetMaxIndex, "末尾のNothing値が詰まる")
            Assert.AreEqual(1, aList.Indexes.Count)
        End Sub

        <Test()> Public Sub TrimEnd_Nothingのみだと()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(0) = Nothing
            aList(1) = Nothing

            aList.TrimEnd()

            Assert.IsFalse(aList.GetMaxIndex.HasValue, "Nothingのみだから値は無くなった")
            Assert.AreEqual(0, aList.Indexes.Count)
        End Sub

        <Test()> Public Sub TrimEnd_末尾のNothing値を詰める_飛び飛びでも末尾Nothingから除去()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(1) = "a"
            aList(3) = Nothing
            aList(5) = Nothing

            aList.TrimEnd()

            Assert.AreEqual("a", aList(1))
            Assert.AreEqual(1, aList.GetMaxIndex, "添え字が飛び飛びでも、末尾Nothingから除去")
            Assert.AreEqual(1, aList.Indexes.Count)
        End Sub

        <Test()> Public Sub TrimEnd_途中のNothing値は詰めない()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(0) = "a"
            aList(1) = Nothing
            aList(2) = "c"

            aList.TrimEnd()

            Assert.AreEqual("a", aList(0))
            Assert.IsNull(aList(1), "末尾じゃないNothingはそのまま")
            Assert.AreEqual("c", aList(2))
            Assert.AreEqual(2, aList.GetMaxIndex)
            Assert.AreEqual(3, aList.Indexes.Count, "Nothingを含めた設定値の数")
        End Sub

        <Test()> Public Sub TrimEnd_先頭のNothing値は詰めない()
            Dim aList As IIndexedList(Of String) = NewList(Of String)()
            aList(0) = Nothing
            aList(1) = "b"
            aList(2) = "c"

            aList.TrimEnd()

            Assert.IsNull(aList(0), "末尾じゃないNothingはそのまま")
            Assert.AreEqual("b", aList(1))
            Assert.AreEqual("c", aList(2))
            Assert.AreEqual(2, aList.GetMaxIndex)
            Assert.AreEqual(3, aList.Indexes.Count, "Nothingを含めた設定値の数")
        End Sub

        <Test()> Public Sub Enumerableが添え字ソート順で取得できる()
            Dim sut As IIndexedList(Of String) = NewList(Of String)()
            sut(2) = "2"
            sut(3) = "3"
            sut(0) = "0"
            sut(1) = "1"
            Assert.That(Join(sut.ToArray), [Is].EqualTo("0 1 2 3"))
        End Sub

        <Test()> Public Sub Enumerableが添え字ソート順で取得できる_歯抜けはそのまま()
            Dim sut As IIndexedList(Of String) = NewList(Of String)()
            sut(2) = "2"
            sut(3) = "3"
            sut(0) = "0"
            Assert.That(Join(sut.ToArray), [Is].EqualTo("0 2 3"))
        End Sub

    End Class
End Namespace
