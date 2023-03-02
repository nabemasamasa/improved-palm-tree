Imports NUnit.Framework
Imports Fhi.Fw.Domain

Public MustInherit Class CollectionUtilTest

#Region "Nested classes..."

    Private Class TestVo
        Public Id As Integer
        Public SubId As Integer
        Public name As String
    End Class

    Private Class StrPvo : Inherits PrimitiveValueObject(Of String)
        Public Sub New(value As String)
            MyBase.New(value)
        End Sub
    End Class

#End Region

    Public Class [Default] : Inherits CollectionUtilTest

        <Test()> Public Sub SubtractBy_String型の差を取得()
            Dim actuals As String() = CollectionUtil.SubtractBy(New String() {"a", "b", "c"}, New String() {"b"})
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("a", actuals(0))
            Assert.AreEqual("c", actuals(1))
        End Sub

        <Test()> Public Sub SubtractBy_Vo型の差を取得_一意キー取得のラムダ式を使用()
            Dim actuals As TestVo() = CollectionUtil.SubtractBy(New TestVo() { _
                                                            New TestVo() With {.Id = 1, .SubId = 1, .name = "11"}, _
                                                            New TestVo() With {.Id = 1, .SubId = 2, .name = "121"}, _
                                                            New TestVo() With {.Id = 2, .SubId = 1, .name = "21"} _
                                                        }, _
                                                        New TestVo() { _
                                                            New TestVo() With {.Id = 1, .SubId = 2, .name = "122"} _
                                                        }, _
                                                        Function(vo As TestVo) EzUtil.MakeKey(vo.Id, vo.SubId))
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("11", actuals(0).name)
            Assert.AreEqual("21", actuals(1).name)
        End Sub

        <Test()> Public Sub MergeBy_Int_通常()
            Dim actuals As ICollection(Of Integer) = CollectionUtil.MergeBy(New Integer() {1, 2, 3}, New Integer() {2, 3, 4})
            Assert.AreEqual(4, actuals.Count)
            Assert.IsTrue(actuals.Contains(1))
            Assert.IsTrue(actuals.Contains(2))
            Assert.IsTrue(actuals.Contains(3))
            Assert.IsTrue(actuals.Contains(4))
        End Sub

        <Test()> Public Sub MergeBy_Int_一つも重複していない場合()
            Dim actuals As ICollection(Of Integer) = CollectionUtil.MergeBy(New Integer() {1, 2, 3}, New Integer() {7, 8, 9})
            Assert.AreEqual(6, actuals.Count)
            Assert.IsTrue(actuals.Contains(1))
            Assert.IsTrue(actuals.Contains(2))
            Assert.IsTrue(actuals.Contains(3))
            Assert.IsTrue(actuals.Contains(7))
            Assert.IsTrue(actuals.Contains(8))
            Assert.IsTrue(actuals.Contains(9))
        End Sub

        <Test()> Public Sub MergeBy_Int_ソートはしない()
            Dim actuals As Integer() = CollectionUtil.MergeBy(New Integer() {3, 1, 2}, New Integer() {})
            Assert.AreEqual(3, actuals.Length)
            Assert.AreEqual(3, actuals(0))
            Assert.AreEqual(1, actuals(1))
            Assert.AreEqual(2, actuals(2))
        End Sub

        <Test()> Public Sub MergeBy_String_通常()
            Dim actuals As String() = CollectionUtil.MergeBy(New String() {"a", "b", "c"}, New String() {"b", "c", "d"})
            Assert.AreEqual(4, actuals.Length)
            Assert.AreEqual("a", actuals(0))
            Assert.AreEqual("b", actuals(1))
            Assert.AreEqual("c", actuals(2))
            Assert.AreEqual("d", actuals(3))
        End Sub

        <Test()> Public Sub MergeBy_String_一つも重複していない場合()
            Dim actuals As String() = CollectionUtil.MergeBy(New String() {"a", "b", "c"}, New String() {"x", "y", "z"})
            Assert.AreEqual(6, actuals.Length)
        End Sub

        <Test()> Public Sub MergeBy_String_ソートはしない()
            Dim actuals As String() = CollectionUtil.MergeBy(New String() {"c", "a", "b"}, New String() {})
            Assert.AreEqual(3, actuals.Length)
            Assert.AreEqual("c", actuals(0))
            Assert.AreEqual("a", actuals(1))
            Assert.AreEqual("b", actuals(2))
        End Sub

        <Test()> Public Sub MergeBy_Vo型をマージ_一意キー取得のラムダ式を使用()
            Dim actuals As TestVo() = CollectionUtil.MergeBy(New TestVo() { _
                                                            New TestVo() With {.Id = 1, .SubId = 1, .name = "a1"}, _
                                                            New TestVo() With {.Id = 1, .SubId = 2, .name = "a2"}, _
                                                            New TestVo() With {.Id = 2, .SubId = 1, .name = "a3"} _
                                                        }, _
                                                        New TestVo() { _
                                                            New TestVo() With {.Id = 1, .SubId = 2, .name = "b1"}, _
                                                            New TestVo() With {.Id = 2, .SubId = 2, .name = "b2"} _
                                                        }, _
                                                        Function(vo As TestVo) EzUtil.MakeKey(vo.Id, vo.SubId))
            Assert.AreEqual(4, actuals.Length)
            Assert.AreEqual("a1", actuals(0).name)
            Assert.AreEqual("a2", actuals(1).name)
            Assert.AreEqual("a3", actuals(2).name)
            Assert.AreEqual("b2", actuals(3).name)
        End Sub

        <Test()> Public Sub Split_商はゼロで余りがある()
            Dim aa As New List(Of Integer)(New Integer() {1, 2, 3})
            Dim actuals As Integer()() = CollectionUtil.Split(aa, 4)
            Assert.AreEqual(1, actuals.Length)
            Assert.AreEqual(3, actuals(0).Length)
        End Sub

        <Test()> Public Sub Split_商があり余りがゼロ()
            Dim aa As New List(Of Integer)(New Integer() {1, 2, 3})
            Dim actuals As Integer()() = CollectionUtil.Split(aa, 3)
            Assert.AreEqual(1, actuals.Length)
            Assert.AreEqual(3, actuals(0).Length)
        End Sub

        <Test()> Public Sub Split_商と余りがある()
            Dim aa As New List(Of Integer)(New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 0})
            Dim actuals As Integer()() = CollectionUtil.Split(aa, 4)
            Assert.AreEqual(3, actuals.Length)
            Assert.AreEqual(4, actuals(0).Length)
            Assert.AreEqual(4, actuals(1).Length)
            Assert.AreEqual(2, actuals(2).Length)
        End Sub

        <Test()> Public Sub Split_要素がゼロ()
            Dim aa As New List(Of Integer)
            Dim actuals As Integer()() = CollectionUtil.Split(aa, 4)
            Assert.AreEqual(0, actuals.Length)
        End Sub

        <Test()> Public Sub MakeVosByKey_Vo型をマージ_一意キー取得のラムダ式を使用()
            Dim actuals As Dictionary(Of Integer, List(Of TestVo)) = _
                        CollectionUtil.MakeVosByKey(New TestVo() { _
                                                        New TestVo() With {.Id = 1, .SubId = 1, .name = "a1"}, _
                                                        New TestVo() With {.Id = 1, .SubId = 2, .name = "a2"}, _
                                                        New TestVo() With {.Id = 2, .SubId = 3, .name = "a3"} _
                                                    }, _
                                                    Function(vo As TestVo) vo.Id)
            Assert.AreEqual(2, actuals.Count)
            Assert.AreEqual(True, actuals.ContainsKey(1))
            Assert.AreEqual(True, actuals.ContainsKey(2))
            Assert.AreEqual(False, actuals.ContainsKey(3))
            Assert.AreEqual(2, actuals(1).Count)
            Assert.AreEqual(1, actuals(2).Count)
            Assert.AreEqual("a1", actuals(1)(0).name)
            Assert.AreEqual("a2", actuals(1)(1).name)
        End Sub

        <Test()> Public Sub ConvEnumerableToList_DictionaryのKeysを変換できる()
            Dim hoge As New Dictionary(Of String, String)
            hoge.Add("A", "a")
            hoge.Add("B", "b")
            Dim actuals As List(Of String) = CollectionUtil.ConvEnumerableToList(Of String)(hoge.Keys)
            Assert.AreEqual("A,B", Join(actuals.ToArray, ","))
        End Sub

        <Test()> Public Sub ConvEnumerableToList_Nothing含む値も変換できる()
            Dim actuals As List(Of String) = CollectionUtil.ConvEnumerableToList(Of String)(New String() {"A", Nothing, "C"})
            Assert.AreEqual(3, actuals.Count)
            Assert.AreEqual("A", actuals(0))
            Assert.IsNull(actuals(1), "Nothingも変換できる")
            Assert.AreEqual("C", actuals(2))
        End Sub
    End Class

    Public Class ContainsTest : Inherits CollectionUtilTest

        <Test()> Public Sub ContainsEmpty_通常パターン()
            Assert.IsFalse(CollectionUtil.ContainsEmpty("a", "b"))
            Assert.IsTrue(CollectionUtil.ContainsEmpty("a", Nothing))
            Assert.IsTrue(CollectionUtil.ContainsEmpty("a", String.Empty))
            Assert.IsTrue(CollectionUtil.ContainsEmpty("a", "  "))
        End Sub

        <Test()> Public Sub ContainsNull_通常パターン()
            Assert.IsFalse(CollectionUtil.ContainsNull("a", "b", -1))
            Assert.IsFalse(CollectionUtil.ContainsNull("", " ", "　", String.Empty))
            Assert.IsTrue(CollectionUtil.ContainsNull("", Nothing))
        End Sub

        <Test()> Public Sub Contains_True()
            Assert.IsTrue(CollectionUtil.Contains(New Object() {"a", "b", "c"}, "b"))
            Assert.IsTrue(CollectionUtil.Contains(New Object() {"a", "b", Nothing}, Nothing))
        End Sub

        <Test()> Public Sub Contains_DateTimeが含まれていても動作する()
            Assert.IsTrue(CollectionUtil.Contains(New Object() {"a", CDate("2011/2/3"), "c"}, CDate("2011/2/3")))
            Assert.IsFalse(CollectionUtil.Contains(New Object() {"a", CDate("2011/2/3"), "c"}, "b"))
            Assert.IsFalse(CollectionUtil.Contains(New Object() {"a", "b", "c"}, CDate("2011/2/3")))
        End Sub

        <Test()> Public Sub Contains_False()
            Assert.IsFalse(CollectionUtil.Contains(New Object() {"a", "b", "c"}, "x"))
            Assert.IsFalse(CollectionUtil.Contains(New Object() {"a", "b", "c"}, Nothing))
            Assert.IsFalse(CollectionUtil.Contains(New Object() {"a", Nothing, "c"}, "x"))
        End Sub

        <Test> Public Sub ContainsIgnorePvo_PVO配列に対してPrimitive型で検証できる()
            Dim pvoArray As Object() = New Object() {New StrPvo("a"), New StrPvo("b"), New StrPvo("c")}
            Assert.IsTrue(CollectionUtil.ContainsIgnorePvo(pvoArray, "b"))
            Assert.IsFalse(CollectionUtil.ContainsIgnorePvo(pvoArray, "d"))
        End Sub

        <Test> Public Sub ContainsIgnorePvo_Primitive型配列に対してPVOで検証できる()
            Dim strArray As Object() = New Object() {"a", "b", "c"}
            Assert.IsTrue(CollectionUtil.ContainsIgnorePvo(strArray, New StrPvo("b")))
            Assert.IsFalse(CollectionUtil.ContainsIgnorePvo(strArray, New StrPvo("d")))
        End Sub
    End Class

    Public Class IsEmpty_IsNotEmpty_ : Inherits CollectionUtilTest

        <Test()> Public Sub NothingならIsEmptyはTrue_ICollection引数()
            Dim collection As ICollection = Nothing
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub NothingならIsEmptyはTrue_ICollectionジェネリック引数()
            Dim collection As ICollection(Of TestVo) = Nothing
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub NothingならIsEmptyはTrue_IList引数()
            Dim collection As System.Collections.IList = Nothing
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub NothingならIsEmptyはTrue_Listジェネリック引数()
            Dim collection As List(Of TestVo) = Nothing
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub 中身ゼロならIsEmptyはTrue_ICollection引数()
            Dim collection As ICollection = New List(Of String)
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub 中身ゼロならIsEmptyはTrue_ICollectionジェネリック引数()
            Dim collection As ICollection(Of TestVo) = New List(Of TestVo)
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub 中身ゼロならIsEmptyはTrue_IList引数()
            Dim collection As System.Collections.IList = New List(Of String)
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub 中身ゼロならIsEmptyはTrue_Listジェネリック引数()
            Dim collection As List(Of TestVo) = New List(Of TestVo)
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub 中身有りならIsEmptyはFalse_ICollection引数()
            Dim collection As ICollection = New List(Of String)(New String() {"aaa"})
            Assert.IsFalse(CollectionUtil.IsEmpty(collection))
        End Sub

        Private Class TestingCollectionObject : Inherits CollectionObject(Of String)
            Public Sub New()
            End Sub
            Public Sub New(ByVal src As CollectionObject(Of String))
                MyBase.New(src)
            End Sub
            Public Sub New(ByVal initialList As IEnumerable(Of String))
                MyBase.New(initialList)
            End Sub
        End Class

        <Test()> Public Sub 中身ゼロならIsEmptyはTrue_CollectionObject()
            Dim collection As TestingCollectionObject = New TestingCollectionObject
            Assert.IsTrue(CollectionUtil.IsEmpty(collection))
        End Sub

        <Test()> Public Sub 中身有りならIsNotEmptyはTrue_CollectionObject()
            Dim collection As TestingCollectionObject = New TestingCollectionObject({"a", "xyz"})
            Assert.IsTrue(CollectionUtil.IsNotEmpty(collection))
        End Sub

    End Class

    Public Class IsEqualIfNull_ : Inherits EzUtilTest

        <Test()> Public Sub 型が同じで_値有有で_一致するなら_true()
            Dim ars1 As String() = {"x", "y"}
            Dim ars2 As String() = {"x", "y"}
            Assert.IsTrue(CollectionUtil.IsEqualIfNull(ars1, ars2), "String配列型")

            Dim ari1 As Integer?() = {11, 22}
            Dim ari2 As Integer?() = {11, 22}
            Assert.IsTrue(CollectionUtil.IsEqualIfNull(ari1, ari2), "Integer?配列型")

            Dim ard1 As DateTime?() = {CDate("2012/03/04 05:06:07"), CDate("2012/12/12 12:12:12")}
            Dim ard2 As DateTime?() = {CDate("2012/03/04 05:06:07"), CDate("2012/12/12 12:12:12")}
            Assert.IsTrue(CollectionUtil.IsEqualIfNull(ard1, ard2), "DateTime?配列型")
        End Sub

        <Test()> Public Sub 型が同じで_値有有で_不一致なら_false()
            Dim ars1 As String() = {"x", "y"}
            Dim ars2 As String() = {"y", "x"}
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ars1, ars2), "String配列型")

            Dim ari1 As Integer?() = {11, 22}
            Dim ari2 As Integer?() = {22, 11}
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ari1, ari2), "Integer?配列型")

            Dim ard1 As DateTime?() = {CDate("2012/03/04 05:06:07"), CDate("2012/12/12 12:12:12")}
            Dim ard2 As DateTime?() = {CDate("2012/12/12 12:12:12"), CDate("2012/03/04 05:06:07")}
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ard1, ard2), "DateTime?配列型")
        End Sub

        <Test()> Public Sub 型が同じで_値有有で_長さ不一致なら_false()
            Dim ars1 As String() = {"x", "y"}
            Dim ars2 As String() = {"x"}
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ars1, ars2), "String配列型")

            Dim ari1 As Integer?() = {11, 22}
            Dim ari2 As Integer?() = {22}
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ari1, ari2), "Integer?配列型")

            Dim ard1 As DateTime?() = {CDate("2012/12/12 12:12:12")}
            Dim ard2 As DateTime?() = {CDate("2012/12/12 12:12:12"), CDate("2012/03/04 05:06:07")}
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ard1, ard2), "DateTime?配列型")
        End Sub

        <Test()> Public Sub 型が同じで_値有無なら_false()
            Dim ars1 As String() = {"x", "y"}
            Dim ars2 As String() = Nothing
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ars1, ars2), "String配列型")

            Dim ari1 As Integer?() = Nothing
            Dim ari2 As Integer?() = {11, 22}
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ari1, ari2), "Integer?配列型")

            Dim ard1 As DateTime?() = {CDate("2012/12/12 12:12:12"), CDate("2012/03/04 05:06:07")}
            Dim ard2 As DateTime?() = Nothing
            Assert.IsFalse(CollectionUtil.IsEqualIfNull(ard1, ard2), "DateTime?配列型")
        End Sub

        <Test()> Public Sub 型が同じで_値無無なら_false()
            Dim ars1 As String() = Nothing
            Dim ars2 As String() = Nothing
            Assert.IsTrue(CollectionUtil.IsEqualIfNull(ars1, ars2), "String配列型")

            Dim ari1 As Integer?() = Nothing
            Dim ari2 As Integer?() = Nothing
            Assert.IsTrue(CollectionUtil.IsEqualIfNull(ari1, ari2), "Integer?配列型")

            Dim ard1 As DateTime?() = Nothing
            Dim ard2 As DateTime?() = Nothing
            Assert.IsTrue(CollectionUtil.IsEqualIfNull(ard1, ard2), "DateTime?配列型")
        End Sub

    End Class

    Public Class CallIfEmptyTest : Inherits CollectionUtilTest

        <Test()> Public Sub Func_Nullだから_処理しない_実行時エラーにならない()
            Dim collectionIfEmpty As String() = Nothing
            Dim actual As Integer = EzUtil.CallIfEmpty(collectionIfEmpty, Function(c) c.Length)
            Assert.That(actual, [Is].EqualTo(0))
        End Sub

        <Test()> Public Sub Func_Emptyだから_処理しない()
            Dim collectionIfEmpty As String() = {}
            Dim actual As Integer = EzUtil.CallIfEmpty(collectionIfEmpty, Function(c) 999)
            Assert.That(actual, [Is].EqualTo(0))
        End Sub

        <Test()> Public Sub Func_有効値だから_処理する()
            Dim collectionIfEmpty As String() = {"a", "b", "c"}
            Dim actual As Integer = EzUtil.CallIfEmpty(collectionIfEmpty, Function(c) c.Length)
            Assert.That(actual, [Is].EqualTo(3))
        End Sub

        <Test()> Public Sub Sub_Nullだから_処理しない_実行時エラーにならない()
            Dim collectionIfEmpty As String() = Nothing
            Dim actual As String() = {"A"}
            EzUtil.CallIfEmpty(collectionIfEmpty, Sub(c) actual = c.Select(Function(v) v.ToUpper).ToArray)
            Assert.That(actual, [Is].EquivalentTo({"A"}))
        End Sub

        <Test()> Public Sub Sub_Emptyだから_処理しない()
            Dim collectionIfEmpty As String() = {}
            Dim actual As String() = {"A"}
            EzUtil.CallIfEmpty(collectionIfEmpty, Sub(c) actual = c.Select(Function(v) v.ToUpper).ToArray)
            Assert.That(actual, [Is].EquivalentTo({"A"}))
        End Sub

        <Test()> Public Sub Sub_有効値だから_処理する()
            Dim collectionIfEmpty As String() = {"a", "b", "c"}
            Dim actual As String() = {"A"}
            EzUtil.CallIfEmpty(collectionIfEmpty, Sub(c) actual = c.Select(Function(v) v.ToUpper).ToArray)
            Assert.That(actual, [Is].EquivalentTo({"A", "B", "C"}))
        End Sub

        <Test()> Public Sub Func_Nullだから_処理しない_実行時エラーにならない_callback引数なし()
            Dim collectionIfEmpty As String() = Nothing
            Dim actual As Integer = EzUtil.CallIfEmpty(collectionIfEmpty, Function() collectionIfEmpty.Length)
            Assert.That(actual, [Is].EqualTo(0))
        End Sub

        <Test()> Public Sub Func_Emptyだから_処理しない_callback引数なし()
            Dim collectionIfEmpty As String() = {}
            Dim actual As Integer = EzUtil.CallIfEmpty(collectionIfEmpty, Function() 999)
            Assert.That(actual, [Is].EqualTo(0))
        End Sub

        <Test()> Public Sub Func_有効値だから_処理する_callback引数なし()
            Dim collectionIfEmpty As String() = {"a", "b", "c"}
            Dim actual As Integer = EzUtil.CallIfEmpty(collectionIfEmpty, Function() collectionIfEmpty.Length)
            Assert.That(actual, [Is].EqualTo(3))
        End Sub

        <Test()> Public Sub Sub_Nullだから_処理しない_実行時エラーにならない_callback引数なし()
            Dim collectionIfEmpty As String() = Nothing
            Dim actual As String() = {"A"}
            EzUtil.CallIfEmpty(collectionIfEmpty, Sub() actual = collectionIfEmpty.Select(Function(v) v.ToUpper).ToArray)
            Assert.That(actual, [Is].EquivalentTo({"A"}))
        End Sub

        <Test()> Public Sub Sub_Emptyだから_処理しない_callback引数なし()
            Dim collectionIfEmpty As String() = {}
            Dim actual As String() = {"A"}
            EzUtil.CallIfEmpty(collectionIfEmpty, Sub() actual = collectionIfEmpty.Select(Function(v) v.ToUpper).ToArray)
            Assert.That(actual, [Is].EquivalentTo({"A"}))
        End Sub

        <Test()> Public Sub Sub_有効値だから_処理する_callback引数なし()
            Dim collectionIfEmpty As String() = {"a", "b", "c"}
            Dim actual As String() = {"A"}
            EzUtil.CallIfEmpty(collectionIfEmpty, Sub() actual = collectionIfEmpty.Select(Function(v) v.ToUpper).ToArray)
            Assert.That(actual, [Is].EquivalentTo({"A", "B", "C"}))
        End Sub

    End Class

    Public Class ConvToTwoDimensionalArrayTest : Inherits CollectionUtilTest

        <Test()> Public Sub _2次元配列を渡したら_2次元配列のまま()
            Dim data As Object(,) = {{"a", "b", "c"}, {"x", "y", "z"}}
            Dim expected As Object(,) = {{"a", "b", "c"}, {"x", "y", "z"}}
            Dim actuals As Object(,) = CollectionUtil.ConvToTwoDimensionalArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _1次元配列を渡したら_2次元配列になる()
            Dim data As Object() = {"a", "b", "c"}
            Dim expected As Object(,) = {{"a", "b", "c"}}
            Dim actuals As Object(,) = CollectionUtil.ConvToTwoDimensionalArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _List値を渡したら_2次元配列になる()
            Dim data As List(Of String) = (New String() {"a", "b", "c"}).ToList
            Dim expected As Object(,) = {{"a", "b", "c"}}
            Dim actuals As Object(,) = CollectionUtil.ConvToTwoDimensionalArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _2段階配列を渡したら_2次元配列になる_1段階目の長さ1()
            Dim data As String()() = {New String() {"a", "b", "c"}}
            Dim expected As Object(,) = {{"a", "b", "c"}}
            Dim actuals As Object(,) = CollectionUtil.ConvToTwoDimensionalArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _2段階配列を渡したら_2次元配列になる()
            Dim data As String()() = {New String() {"a", "b", "c"}, New String() {"x", "y", "z"}}
            Dim expected As Object(,) = {{"a", "b", "c"}, {"x", "y", "z"}}
            Dim actuals As Object(,) = CollectionUtil.ConvToTwoDimensionalArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub 文字列を渡しても_2次元配列になる()
            Dim expected As Object(,) = {{"abc"}}
            Dim actuals As Object(,) = CollectionUtil.ConvToTwoDimensionalArray("abc")
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub プリミティブ値を渡しても_2次元配列になる()
            Dim expected As Object(,) = {{123}}
            Dim actuals As Object(,) = CollectionUtil.ConvToTwoDimensionalArray(123)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub null値を渡したら_長さゼロの2次元配列になる()
            Dim expected As Object(,) = {{}}
            Dim actuals As Object(,) = CollectionUtil.ConvToTwoDimensionalArray(Nothing)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class ConvToTwoJaggedArrayTest : Inherits CollectionUtilTest

        <Test()> Public Sub _2段階配列を渡したら_2段階配列のまま()
            Dim data As String()() = {New String() {"a", "b", "c"}, New String() {"x", "y", "z"}}
            Dim expected As Object()() = {New Object() {"a", "b", "c"}, New Object() {"x", "y", "z"}}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _1次元配列を渡したら_2段階配列になる()
            Dim data As Object() = {"a", "b", 3}
            Dim expected As Object()() = {New Object() {"a", "b", 3}}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _1次元配列を渡したら_2段階配列になる_ConvToTwoJaggedArrayAsStringはすべて文字列型になる()
            Dim data As Object() = {"a", "b", 3}
            Dim expected As String()() = {New String() {"a", "b", "3"}}
            Dim actuals As String()() = CollectionUtil.ConvToTwoJaggedArrayAsString(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _List値を渡したら_2段階配列になる()
            Dim data As List(Of String) = (New String() {"a", "b", "c"}).ToList
            Dim expected As Object()() = {New Object() {"a", "b", "c"}}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _2次元配列を渡したら_2段階配列になる_1段階目の長さ1()
            Dim data As Object(,) = {{"a", "b", "c"}}
            Dim expected As Object()() = {New String() {"a", "b", "c"}}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub _2段階配列を渡したら_2次元配列になる()
            Dim data As Object(,) = {{"a", "b", "c"}, {"x", "y", "z"}}
            Dim expected As Object()() = {New Object() {"a", "b", "c"}, New Object() {"x", "y", "z"}}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArray(data)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub 文字列を渡しても_2段階配列になる()
            Dim expected As Object()() = {New Object() {"abc"}}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArray("abc")
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub プリミティブ値を渡しても_2段階配列になる()
            Dim expected As Object()() = {New Object() {123}}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArray(123)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub プリミティブ値を渡しても_2段階配列になる_ConvToTwoJaggedArrayAsStringはすべて文字列型になる()
            Dim expected As String()() = {New String() {"123"}}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArrayAsString(123)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub null値を渡したら_長さゼロの2段階配列になる()
            Dim expected As Object()() = {}
            Dim actuals As Object()() = CollectionUtil.ConvToTwoJaggedArray(Nothing)
            Assert.That(actuals, [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class UpdateRecentValuesTest : Inherits CollectionUtilTest

        <TestCase("a", New String() {}, 10, {"a"})>
        <TestCase("x", Nothing, 10, {"x"})>
        <TestCase("a", {"y", "z"}, 10, {"a", "y", "z"})>
        <TestCase("a", {"a", "b"}, 10, {"a", "b"})>
        <TestCase("a", {"A", "B"}, 10, {"a", "B"})>
        <TestCase("abc", {"ABC"}, 10, {"abc"})>
        <TestCase("A", {"c", "a", "b"}, 10, {"A", "c", "b"})>
        <TestCase("a", {"d", "c", "A"}, 10, {"a", "d", "c"})>
        <TestCase("A", {"g", "f", "e"}, 3, {"A", "g", "f"})>
        <TestCase("a", {"h", "i"}, 3, {"a", "h", "i"})>
        Public Sub 最近使った情報を_利用順に更新できる(newValue As String, values As String(), maxCount As Integer, expected As String())
            Dim actual As String() = CollectionUtil.UpdateRecentValues(newValue, values, maxCount)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

    End Class

End Class
