Imports System.Collections.Generic
Imports NUnit.Framework

Public Class TestingLazyInit

    Public Shared ReadOnly names As New List(Of String)
    Public Shared ReadOnly Instance As New TestingLazyInit("TestingLazyInit")

    Private Class Holder
        Public Shared ReadOnly Instance As New TestingLazyInit("TestingLazyInit.Holder")
    End Class

    Private name As String

    Public Sub New(ByVal name As String)
        Me.name = name
        names.Add(name)
    End Sub

    Public Function GetName() As String
        Return name
    End Function

    Public Shared Function HolderInstance() As TestingLazyInit
        Return Holder.Instance
    End Function
End Class

Public MustInherit Class LearningTest

    Public Class コレクション操作のLearning : Inherits LearningTest

        <Test()> Public Sub 配列はICollectionにキャスト出来る_これが良いのか悪いのか不明()

            Dim strings As String() = {"a", "b"}
            Dim ints As Integer() = {1, 2, 3}

            Dim actualStrings As ICollection(Of String) = DirectCast(strings, ICollection(Of String))
            Assert.AreEqual(2, actualStrings.Count)
            Assert.IsTrue(actualStrings.Contains("a"))
            Assert.IsTrue(actualStrings.Contains("b"))

            Dim actualInts As ICollection(Of Integer) = DirectCast(ints, ICollection(Of Integer))
            Assert.AreEqual(3, actualInts.Count)
            Assert.IsTrue(actualInts.Contains(1))
            Assert.IsTrue(actualInts.Contains(2))
            Assert.IsTrue(actualInts.Contains(3))
        End Sub

        Private Class NestedClass
            Public Name As String
            Public Sub New()
            End Sub
            Public Sub New(ByVal name As String)
                Me.Name = name
            End Sub
        End Class
        Private Class NestedChildClass : Inherits NestedClass
            Public FirstName As String
            Public Sub New()
            End Sub
            Public Sub New(ByVal name As String, ByVal firstName As String)
                Me.Name = name
                Me.FirstName = firstName
            End Sub
        End Class

        <Test()> Public Sub サブクラスListを配列にすれば_親クラスListに追加できる()
            Dim baseList As New List(Of NestedClass)
            baseList.Add(New NestedClass("FUGA"))
            Dim children As New List(Of NestedChildClass)
            children.Add(New NestedChildClass("ff", "ww"))

            baseList.AddRange(children.ToArray)
            '↓だとコンパイルエラー
            'fuga.AddRange(children)

            Assert.AreEqual(2, baseList.Count)
            Assert.AreEqual("FUGA", baseList(0).Name)
            Assert.AreEqual("ff", baseList(1).Name)
        End Sub

    End Class

    Public Class String型のLearning : Inherits LearningTest

        <Test()> Public Sub String_Format_数値の整形()
            Assert.AreEqual("001", String.Format("{0:000}", 1))
            Assert.AreEqual("013", String.Format("{0:000}", 13))
        End Sub

        <Test()> Public Sub String_Format_文字列の整形()
            Assert.AreEqual("|       abc|", String.Format("|{0,10}|", "abc"))
            Assert.AreEqual("|abc       |", String.Format("|{0,-10}|", "abc"))
        End Sub

        <Test()> Public Sub Null値のString変換_Convert_ToStringだとStringEmptyになる()
            Dim nullString As Object = Nothing
            Assert.AreEqual("", Convert.ToString(nullString))
        End Sub

        <Test()> Public Sub Null値のString変換_DirectCastだとNull値はNullのまま()
            Dim nullString As Object = Nothing
            Assert.IsNull(DirectCast(nullString, String))
        End Sub

        <Test()> Public Sub 空文字とNullの比較_比較したい値がEmptyかどうか判定するならEquals()
            Dim compareValue As String = Nothing
            Assert.IsFalse("".Equals(compareValue))
            compareValue = ""
            Assert.IsTrue("".Equals(compareValue))
        End Sub

        <Test()> Public Sub 空文字とNullの比較_比較したい値がEmptyかどうか判定するならEquals_EQ演算子だとNG()
            Dim compareValue As String = Nothing
            Assert.IsTrue("" = compareValue)        ' trueになってしまう！！
            compareValue = ""
            Assert.IsTrue("" = compareValue)
        End Sub

        <Test()> Public Sub 空文字とNullの比較_比較したい値がNullかどうか判定するならIs演算子()
            Dim compareValue As String = Nothing
            Assert.IsTrue(compareValue Is Nothing)
            compareValue = ""
            Assert.IsFalse(compareValue Is Nothing)
        End Sub

        <Test()> Public Sub 空文字とNullの比較_比較したい値がNullかどうか判定するならIs演算子_EQ演算子だとNG()
            Dim compareValue As String = Nothing
            Assert.IsTrue(compareValue = Nothing)
            compareValue = ""
            Assert.IsTrue(compareValue = Nothing)   ' trueになってしまう！！
        End Sub

    End Class

    Public Class Char型のLearning : Inherits LearningTest

        <Test()> Public Sub CharIsLetter_半角英と半角カナと全角英と全角カナと全角ひらがなと漢字がtrue()
            Assert.IsTrue(Char.IsLetter("a"c))
            Assert.IsTrue(Char.IsLetter("A"c))
            Assert.IsTrue(Char.IsLetter("ｱ"c))
            Assert.IsTrue(Char.IsLetter("ａ"c))
            Assert.IsTrue(Char.IsLetter("Ａ"c))
            Assert.IsTrue(Char.IsLetter("ア"c))
            Assert.IsTrue(Char.IsLetter("あ"c))
            Assert.IsTrue(Char.IsLetter("亜"c))

            Assert.IsFalse(Char.IsLetter("1"c))
            Assert.IsFalse(Char.IsLetter("*"c))
            Assert.IsFalse(Char.IsLetter("１"c))
            Assert.IsFalse(Char.IsLetter("＊"c))
        End Sub

        <Test()> Public Sub CharIsNumber_半角数と全角数がtrue()
            Assert.IsTrue(Char.IsNumber("1"c))
            Assert.IsTrue(Char.IsNumber("１"c))

            Assert.IsFalse(Char.IsNumber("a"c))
            Assert.IsFalse(Char.IsNumber("A"c))
            Assert.IsFalse(Char.IsNumber("ｱ"c))
            Assert.IsFalse(Char.IsNumber("*"c))
            Assert.IsFalse(Char.IsNumber("ａ"c))
            Assert.IsFalse(Char.IsNumber("Ａ"c))
            Assert.IsFalse(Char.IsNumber("ア"c))
            Assert.IsFalse(Char.IsNumber("あ"c))
            Assert.IsFalse(Char.IsNumber("亜"c))
            Assert.IsFalse(Char.IsNumber("＊"c))
        End Sub

    End Class

    Public Class Path型のLearning : Inherits LearningTest

        <Test()> Public Sub PathCombine_Yenマーク付きでもYenマーク無しでも保管して結合してくれる()
            Assert.AreEqual("c:\hoge\fuga", System.IO.Path.Combine("c:\hoge", "fuga"))
            Assert.AreEqual("c:\hoge\fuga", System.IO.Path.Combine("c:\hoge\", "fuga"))
        End Sub

    End Class

    Public Class HolderLazyInitのLearning : Inherits LearningTest

        <Test()> Public Sub HolderLazyInit_入れ子クラスのクラス変数は_アクセスするまで初期化されない()
            Assert.AreEqual("TestingLazyInit", TestingLazyInit.Instance.GetName)
            ' この時点では、TestingLazyInitクラスのクラス変数しか初期化されていない（入れ子クラスHolderは初期化されていない）
            Assert.AreEqual(1, TestingLazyInit.names.Count)

            ' 入れ子クラスHolderにアクセス
            Assert.AreEqual("TestingLazyInit.Holder", TestingLazyInit.HolderInstance.GetName)
            ' ↑でようやく入れ子クラスHolderのクラス変数も初期化される
            Assert.AreEqual(2, TestingLazyInit.names.Count)
        End Sub

    End Class

    Public Class VB機能_if : Inherits LearningTest

        <Test()> Public Sub 三項演算子_if()
            Dim obj As Object = Nothing
            Assert.AreEqual("2", If(obj Is Nothing, "2", obj.ToString), "実行時エラーにならず実行できる")
        End Sub

        <Test()> Public Sub 三項演算子_ifではなくiifだとエラーになる()
            Try
                Dim obj As Object = Nothing
                Dim actual As Object = IIf(obj Is Nothing, "2", obj.ToString)
                Assert.Fail()
            Catch expected As NullReferenceException
                Assert.IsTrue(True, "iifだと、第二引数第三引数をすべて評価してしまうのでエラーになる")
            End Try
        End Sub

    End Class

    Public Class VB機能_Join_Split : Inherits LearningTest

        <Test()> Public Sub Join_引数がNothingだと_結果はNothing()
            Dim actual As String = Join(Nothing, ",")
            Assert.IsNull(actual, "結果はnull")
        End Sub

        <Test()> Public Sub Join_引数が長さ0のコレクションも_結果はNothing()
            Dim actual As String = Join(New String() {}, ",")
            Assert.IsNull(actual, "結果はnull")
        End Sub

        <Test()> Public Sub Split_引数がNothingだと_長さ1で空文字をもつ_引数空文字と同じ()
            Dim actuals As String() = Split(Nothing, ",")
            Assert.IsNotNull(actuals, "Nullにはならない")
            Assert.AreEqual(1, actuals.Length, "Nothingでも長さ1")
            Assert.AreEqual("", actuals(0), "Nothingでも中身は空文字")
        End Sub

        <Test()> Public Sub Split_引数が空文字だと_長さ1で空文字をもつ()
            Dim actuals As String() = Split("", ",")
            Assert.IsNotNull(actuals, "Nullにはならない")
            Assert.AreEqual(1, actuals.Length, "長さ1")
            Assert.AreEqual("", actuals(0), "中身は空文字")
        End Sub

    End Class

    Public Class VB機能_ほか : Inherits LearningTest

        <Test()> Public Sub Nullableと整数値のEquals評価でNullableを引数にしても正しく評価できる()
            Dim l As Long = 10
            Dim l1 As Long? = 10
            Dim l2 As Long? = Nothing

            Assert.AreEqual(True, l.Equals(l1))
            Assert.AreEqual(False, l.Equals(l2))
        End Sub

        <Test()> Public Sub 匿名型の判別()
            Dim hoge = New With {.Id = 123}
            Assert.AreEqual(True, hoge.GetType.Name.StartsWith("VB$AnonymousType"))
        End Sub

    End Class

End Class

