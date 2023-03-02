Imports Fhi.Fw.Domain
Imports NUnit.Framework

Namespace Util
    Public MustInherit Class ChangeJudgeerTest

#Region "テストVo"
        Private Class TestVo
            Private _id As Integer?
            Private _Name As String
            Private _UpdatedUserId As String
            Private _UpdatedDate As String
            Private _UpdatedTime As String

            Public Property Id() As Integer?
                Get
                    Return _id
                End Get
                Set(ByVal value As Integer?)
                    _id = value
                End Set
            End Property

            Public Property Name() As String
                Get
                    Return _Name
                End Get
                Set(ByVal value As String)
                    _Name = value
                End Set
            End Property

            Public Property UpdatedUserId() As String
                Get
                    Return _UpdatedUserId
                End Get
                Set(ByVal value As String)
                    _UpdatedUserId = value
                End Set
            End Property

            Public Property UpdatedDate() As String
                Get
                    Return _UpdatedDate
                End Get
                Set(ByVal value As String)
                    _UpdatedDate = value
                End Set
            End Property

            Public Property UpdatedTime() As String
                Get
                    Return _UpdatedTime
                End Get
                Set(ByVal value As String)
                    _UpdatedTime = value
                End Set
            End Property

        End Class

        Private Class TestVo2 : Inherits TestVo
            Private _subId As Integer?

            Public Property SubId() As Integer?
                Get
                    Return _subId
                End Get
                Set(ByVal value As Integer?)
                    _subId = value
                End Set
            End Property
        End Class
#End Region
#Region "Nested classes"
        Private Class ImmutableStringPair : Inherits AbstractImmutablePair(Of String, String)

            Public ReadOnly Property Pair1() As String
                Get
                    Return MyBase.PairA
                End Get
            End Property

            Public ReadOnly Property Pair2() As String
                Get
                    Return MyBase.PairB
                End Get
            End Property

            Public Sub New(ByVal pair1 As String, ByVal pair2 As String)
                MyBase.New(pair1, pair2)
            End Sub
        End Class
        Private Class TestingValueObject : Inherits ValueObject
            Public ReadOnly Name As String
            Public ReadOnly ZipCode As String
            Public ReadOnly Address As String
            Public Sub New(name As String, zipCode As String, address As String)
                Me.Name = name
                Me.ZipCode = zipCode
                Me.Address = address
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Name, ZipCode, Address}
            End Function
        End Class
        Private Class TestingNameAttribute : Inherits Attribute
        End Class
        Private Class AttrVo
            Public Property Id As Integer
            <TestingName()>
            Public Property Name As String
        End Class
        Private Class AttrParentVo
            Public Property Id As Integer
            <TestingName()>
            Public Property Vo As AttrVo
        End Class
#End Region

        Private Class KeyPlugin : Implements ChangeJudgeer(Of TestVo2).Plugin

            Public Function MakeKey(ByVal obj As TestVo2) As Object Implements ChangeJudgeer(Of TestVo2).Plugin.MakeKey
                Return EzUtil.MakeKey(obj.Id, obj.SubId)
            End Function
        End Class

        Private Class ChangeJudgeerByKey : Inherits ChangeJudgeer(Of TestVo2)

            Public Sub New()
                MyBase.New(New KeyPlugin)
            End Sub

            Public Sub New(ByVal beforeDatas As ICollection(Of TestVo2))
                MyBase.New(beforeDatas, New KeyPlugin)
            End Sub
        End Class

        Private Function NewTestVo(ByVal id As Integer?, ByVal name As String) As TestVo
            Return NewTestVo(id, name, "HOGE", "2010-11-11", "12:13:14")
        End Function

        Private Function NewTestVo(ByVal id As Integer?, ByVal name As String, ByVal UpdatedUserId As String, ByVal UpdatedDate As String, ByVal UpdatedTime As String) As TestVo
            Dim vo As New TestVo
            vo.Id = id
            vo.Name = name
            vo.UpdatedUserId = UpdatedUserId
            vo.UpdatedDate = UpdatedDate
            vo.UpdatedTime = UpdatedTime
            Return vo
        End Function

        Private Function NewTestVo2(ByVal id As Integer?, ByVal subId As Integer?, ByVal name As String) As TestVo2
            Return NewTestVo2(id, subId, name, "HOGE", "2010-11-11", "12:13:14")
        End Function

        Private Function NewTestVo2(ByVal id As Integer?, ByVal subId As Integer?, ByVal name As String, ByVal UpdatedUserId As String, ByVal UpdatedDate As String, ByVal UpdatedTime As String) As TestVo2
            Dim vo As New TestVo2
            vo.Id = id
            vo.SubId = subId
            vo.Name = name
            vo.UpdatedUserId = UpdatedUserId
            vo.UpdatedDate = UpdatedDate
            vo.UpdatedTime = UpdatedTime
            Return vo
        End Function

        Public Class Initial : Inherits ChangeJudgeerTest

            <Test()> Public Sub HasChanged_変更しない同じ情報ならfalse()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testee.SetUpdatedVos(testingVos)

                Assert.IsFalse(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_値が変更されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                v1.Id = 3
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_値は変わらなくてもvoが追加されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Add(NewTestVo(3, "ccc"))
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_値は変わらなくてもvoが削除されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Remove(v1)
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_値は変わらなくてもvoが削除と追加されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Remove(v1)
                testingVos.Add(NewTestVo(3, "ccc"))
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_変更しない同じ情報のまま_AddBeforeUpdatedVosしただけならfalse()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                Dim v3 As TestVo = NewTestVo(3, "ccc")
                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo)(v3))
                testingVos.Add(v3)

                testee.SetUpdatedVos(testingVos)

                Assert.IsFalse(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_値が変更された後に_AddBeforeUpdatedVosしたならtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                v1.Id = 3

                Dim v3 As TestVo = NewTestVo(3, "ccc")
                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo)(v3))
                testingVos.Add(v3)

                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_値を変更せず_AddBeforeUpdatedVosした値が変更されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                Dim v3 As TestVo = NewTestVo(3, "ccc")
                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo)(v3))
                testingVos.Add(v3)

                v3.Id = 9

                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_値を変更せず_AddBeforeUpdatedVosした後にvoが追加されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                Dim v3 As TestVo = NewTestVo(3, "ccc")
                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo)(v3))
                testingVos.Add(v3)

                testingVos.Add(NewTestVo(4, "ddd"))
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub HasChanged_値を変更せず_AddBeforeUpdatedVosした後にvoが削除されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                Dim v3 As TestVo = NewTestVo(3, "ccc")
                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo)(v3))
                testingVos.Add(v3)

                testingVos.Remove(v1)
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub WasDeleted_voが削除されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Remove(v1)
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.WasDeleted)
            End Sub

            <Test()> Public Sub WasDeleted_voが追加されてもfalse()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Add(NewTestVo(3, "ccc"))
                testee.SetUpdatedVos(testingVos)

                Assert.IsFalse(testee.WasDeleted)
            End Sub

            <Test()> Public Sub WasInserted_voが削除されてもfalse()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Remove(v1)
                testee.SetUpdatedVos(testingVos)

                Assert.IsFalse(testee.WasInserted)
            End Sub

            <Test()> Public Sub WasInserted_voが追加されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Add(NewTestVo(3, "ccc"))
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.WasInserted)
            End Sub

            <Test()> Public Sub WasUpdated_voが削除されてもfalse()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Remove(v1)
                testee.SetUpdatedVos(testingVos)

                Assert.IsFalse(testee.WasUpdated)
            End Sub

            <Test()> Public Sub WasUpdated_voが追加されてもfalse()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                testingVos.Add(NewTestVo(3, "ccc"))
                testee.SetUpdatedVos(testingVos)

                Assert.IsFalse(testee.WasUpdated)
            End Sub

            <Test()> Public Sub WasUpdated_voが変更されたらtrue()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                v2.Name = "ccc"
                testee.SetUpdatedVos(testingVos)

                Assert.IsTrue(testee.WasUpdated)
            End Sub

            <Test()> Public Sub ExtractInsertedVos_列挙順序はSetUpdatedVosの時の並び順()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim testee As New ChangeJudgeer(Of TestVo)(EzUtil.NewList(v1))

                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim v3 As TestVo = NewTestVo(3, "ccc")
                Dim v4 As TestVo = NewTestVo(4, "ddd")
                Dim v5 As TestVo = NewTestVo(5, "eee")
                testee.SetUpdatedVos(EzUtil.NewList(v5, v4, v3, v2, v1))

                Dim actuals As TestVo() = testee.ExtractInsertedVos
                Assert.AreEqual(4, actuals.Length)
                Assert.AreEqual("eee", actuals(0).Name)
                Assert.AreEqual("ddd", actuals(1).Name)
                Assert.AreEqual("ccc", actuals(2).Name)
                Assert.AreEqual("bbb", actuals(3).Name)
            End Sub

            <Test()> Public Sub ExtractDeletedVos_列挙順序はコンストラクタ時_またはSupersedeBeforeUpdateVos時の並び順()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim v3 As TestVo = NewTestVo(3, "ccc")
                Dim v4 As TestVo = NewTestVo(4, "ddd")
                Dim v5 As TestVo = NewTestVo(5, "eee")
                Dim testee As New ChangeJudgeer(Of TestVo)(EzUtil.NewList(v5, v4, v3, v2, v1))

                testee.SetUpdatedVos(EzUtil.NewList(v3))

                Dim actuals As TestVo() = testee.ExtractDeletedVos
                Assert.AreEqual(4, actuals.Length)
                Assert.AreEqual("eee", actuals(0).Name)
                Assert.AreEqual("ddd", actuals(1).Name)
                Assert.AreEqual("bbb", actuals(2).Name)
                Assert.AreEqual("aaa", actuals(3).Name)
            End Sub

            <Test()> Public Sub ExtractUpdatedVos_変更されたデータだけ抽出()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                v2.Name = "ccc"
                testee.SetUpdatedVos(testingVos)

                Dim actuals As ChangeJudgeer(Of TestVo).UpdateInfo() = testee.ExtractUpdatedVos
                Assert.AreEqual(1, actuals.Length)
                Assert.AreEqual("bbb", actuals(0).BeforeVo.Name)
                Assert.AreEqual("ccc", actuals(0).AfterVo.Name)
            End Sub

            <Test()> Public Sub ExtractUpdatedVos_列挙順序はSetUpdatedVosの時の並び順()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim v3 As TestVo = NewTestVo(3, "ccc")
                Dim v4 As TestVo = NewTestVo(4, "ddd")
                Dim v5 As TestVo = NewTestVo(5, "eee")
                Dim testee As New ChangeJudgeer(Of TestVo)(EzUtil.NewList(v1, v2, v3, v4, v5))

                v1.Name = "o"
                v2.Name = "p"
                v3.Name = "q"
                v4.Name = "r"
                v5.Name = "s"
                testee.SetUpdatedVos(EzUtil.NewList(v5, v4, v3, v2, v1))

                Dim actuals As ChangeJudgeer(Of TestVo).UpdateInfo() = testee.ExtractUpdatedVos
                Assert.AreEqual(5, actuals.Length)
                Assert.AreEqual("s", actuals(0).AfterVo.Name)
                Assert.AreEqual("r", actuals(1).AfterVo.Name)
                Assert.AreEqual("q", actuals(2).AfterVo.Name)
                Assert.AreEqual("p", actuals(3).AfterVo.Name)
                Assert.AreEqual("o", actuals(4).AfterVo.Name)
            End Sub

            <Test()> Public Sub SetUpdatedVos_SetUpdatedVosした後のデータ変更は影響しない()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                v2.Name = "ccc"
                testee.SetUpdatedVos(testingVos)

                ' SetUpdatedVos後に変更
                v2.Name = "dddd"

                Dim actuals As ChangeJudgeer(Of TestVo).UpdateInfo() = testee.ExtractUpdatedVos
                Assert.AreEqual(1, actuals.Length)
                Assert.AreEqual("bbb", actuals(0).BeforeVo.Name)
                Assert.AreEqual("ccc", actuals(0).AfterVo.Name)     ' ddddにはならない!!
            End Sub

            <Test()> Public Sub NotifySaveEnd_SetUpdatedVosした後にデータを変更し一度Judgeしたあと_NotifySaveEndしてすぐSetUpdatedVosしたらデータ変更だけ拾う()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                v2.Name = "ccc"
                testee.SetUpdatedVos(testingVos)

                ' SetUpdatedVos後に変更
                v2.Name = "dddd"

                Dim actuals As ChangeJudgeer(Of TestVo).UpdateInfo() = testee.ExtractUpdatedVos
                Assert.AreEqual(1, actuals.Length)
                Assert.AreEqual("bbb", actuals(0).BeforeVo.Name)
                Assert.AreEqual("ccc", actuals(0).AfterVo.Name)     ' ddddにはならない!!

                testee.NotifySaveEnd()

                testee.SetUpdatedVos(testingVos)

                Dim actuals2 As ChangeJudgeer(Of TestVo).UpdateInfo() = testee.ExtractUpdatedVos
                Assert.AreEqual(1, actuals2.Length)
                Assert.AreEqual("ccc", actuals2(0).BeforeVo.Name)
                Assert.AreEqual("dddd", actuals2(0).AfterVo.Name)
            End Sub

            <Test()> Public Sub DetectInstance_ExtractしたVoで元々のVoインスタンスを返す()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testingVos As New List(Of TestVo)(New TestVo() {v1, v2})
                Dim testee As New ChangeJudgeer(Of TestVo)(testingVos)

                v2.Name = "ccc"
                testee.SetUpdatedVos(testingVos)

                ' SetUpdatedVos後に変更
                v2.Name = "dddd"

                Dim actuals As ChangeJudgeer(Of TestVo).UpdateInfo() = testee.ExtractUpdatedVos
                Assert.AreEqual(1, actuals.Length)
                Assert.AreEqual("bbb", actuals(0).BeforeVo.Name)
                Assert.AreEqual("ccc", actuals(0).AfterVo.Name)     ' ddddにはならない!!

                Assert.AreSame(v2, testee.DetectOriginInstance(actuals(0).AfterVo))
                Assert.AreSame(v2, testee.DetectOriginInstance(actuals(0).BeforeVo))
            End Sub

            <Test()> Public Sub AddBeforeUpdatedVos_同一キーなら上書きされる()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim v2 As TestVo = NewTestVo(2, "bbb")
                Dim testee As New ChangeJudgeer(Of TestVo)(EzUtil.NewList(v1, v2))

                v2.Name = "ddd"
                Dim v3 As TestVo = NewTestVo(3, "ccc")
                testee.AddBeforeUpdatedVos(EzUtil.NewList(v2, v3))

                v2.Name = "ddd"
                testee.SetUpdatedVos(EzUtil.NewList(v1, v2, v3))

                Assert.IsFalse(testee.HasChanged)
            End Sub

            <Test()> Public Sub 空インスタンスにAddBeforeUpdatedVosしても処理できる()
                Dim v1 As TestVo = NewTestVo(1, "aaa")
                Dim testee As New ChangeJudgeer(Of TestVo)

                testee.AddBeforeUpdatedVos(New TestVo() {v1})

                v1.Name = "bbb"

                testee.SetUpdatedVos(New TestVo() {v1})

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub String型は不変オブジェクトなので_WasUpdatedはFalse()
                Dim testee As New ChangeJudgeer(Of String)(New String() {"a", "b", "c"})
                testee.SetUpdatedVos(New String() {"a", "b2", "d"})

                Assert.AreEqual(True, testee.WasDeleted)
                Assert.AreEqual(2, testee.ExtractDeletedVos.Length)
                Assert.AreEqual("b", testee.ExtractDeletedVos(0))
                Assert.AreEqual("c", testee.ExtractDeletedVos(1))

                Assert.AreEqual(True, testee.WasInserted)
                Assert.AreEqual(2, testee.ExtractInsertedVos.Length)
                Assert.AreEqual("b2", testee.ExtractInsertedVos(0))
                Assert.AreEqual("d", testee.ExtractInsertedVos(1))

                Assert.AreEqual(False, testee.WasUpdated, "不変オブジェクトだから、変化は掴めない")
            End Sub


            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_変更しない同じ情報ならfalse()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                Assert.IsFalse(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_値が変更されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "fff"), NewTestVo2(2, 0, "bbb")))

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_値は変わらなくてもvoが追加されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc")))

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_値は変わらなくてもvoが削除されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(2, 0, "bbb")))

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_値は変わらなくてもvoが削除と追加されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc")))

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_変更しない同じ情報のまま_AddBeforeUpdatedVosしただけならfalse()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(3, 4, "ccc")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc")))

                Assert.IsFalse(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_値が変更された後に_AddBeforeUpdatedVosしたならtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(3, 4, "ccc")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "fff"), NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc")))

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_値を変更せず_AddBeforeUpdatedVosした値が変更されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(3, 4, "ccc")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "zzzz")))

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_値を変更せず_AddBeforeUpdatedVosした後にvoが追加されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(3, 4, "ccc")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc"), NewTestVo2(4, 9, "ddd")))

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_HasChanged_値を変更せず_AddBeforeUpdatedVosした後にvoが削除されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.AddBeforeUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(3, 4, "ccc")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc")))

                Assert.IsTrue(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_WasDeleted_voが削除されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(2, 0, "bbb")))

                Assert.IsTrue(testee.WasDeleted)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_WasDeleted_voが追加されてもfalse()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc")))

                Assert.IsFalse(testee.WasDeleted)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_WasInserted_voが削除されてもfalse()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(2, 0, "bbb")))

                Assert.IsFalse(testee.WasInserted)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_WasInserted_voが追加されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc")))

                Assert.IsTrue(testee.WasInserted)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_WasUpdated_voが削除されてもfalse()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(2, 0, "bbb")))

                Assert.IsFalse(testee.WasUpdated)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_WasUpdated_voが追加されてもfalse()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb"), NewTestVo2(3, 4, "ccc")))

                Assert.IsFalse(testee.WasUpdated)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_WasUpdated_voが変更されたらtrue()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "ffff"), NewTestVo2(2, 0, "bbb")))

                Assert.IsTrue(testee.WasUpdated)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_キー値を使用するので別々のインスタンスでも判定出来る_ExtractUpdatedVos_変更されたデータだけ抽出()

                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "bbb")))

                testee.SetUpdatedVos(EzUtil.NewList(Of TestVo2)(NewTestVo2(1, 3, "aaa"), NewTestVo2(2, 0, "fff")))

                Dim actuals As ChangeJudgeer(Of TestVo2).UpdateInfo() = testee.ExtractUpdatedVos
                Assert.AreEqual(1, actuals.Length)
                Assert.AreEqual("bbb", actuals(0).BeforeVo.Name)
                Assert.AreEqual("fff", actuals(0).AfterVo.Name)
                Assert.AreEqual(2, actuals(0).AfterVo.Id)
                Assert.AreEqual(0, actuals(0).AfterVo.SubId)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_AddBeforeUpdatedVos_同一キーなら上書きされる()
                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(NewTestVo2(1, 1, "aaa"), NewTestVo2(2, 2, "bbb")))

                Dim v3 As TestVo = NewTestVo(3, "ccc")
                testee.AddBeforeUpdatedVos(EzUtil.NewList(NewTestVo2(2, 2, "ddd"), NewTestVo2(3, 4, "ccc")))

                testee.SetUpdatedVos(EzUtil.NewList(NewTestVo2(1, 1, "aaa"), NewTestVo2(2, 2, "ddd"), NewTestVo2(3, 4, "ccc")))

                Assert.IsFalse(testee.HasChanged)
            End Sub

            <Test()> Public Sub ChangeJudgeerByKey_DetectInstance_ExtractしたVoで元々のVoインスタンスを返す()
                Dim v2 As TestVo2 = NewTestVo2(2, 2, "bbb")
                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(NewTestVo2(1, 1, "aaa"), v2))

                v2.Name = "ccc"
                testee.SetUpdatedVos(EzUtil.NewList(NewTestVo2(1, 1, "aaa"), v2))

                Dim actuals As ChangeJudgeer(Of TestVo2).UpdateInfo() = testee.ExtractUpdatedVos
                Assert.AreEqual(1, actuals.Length)
                Assert.AreEqual("bbb", actuals(0).BeforeVo.Name)
                Assert.AreEqual("ccc", actuals(0).AfterVo.Name)

                Assert.AreSame(v2, testee.DetectOriginInstance(actuals(0).AfterVo))
                Assert.AreSame(v2, testee.DetectOriginInstance(actuals(0).BeforeVo))
            End Sub

            <Test()> Public Sub 不変オブジェクトは_コンストラクタにbeforeVosを指定したら_例外()
                Try
                    Dim testee As New ChangeJudgeer(Of ImmutableStringPair)(EzUtil.NewList(New ImmutableStringPair("A", "B"))) With {.IsTypeImmutable = True}
                    Assert.Fail("不変オブジェクトは、コンストラクタに beforeVos 指定したらエラーになるべき")

                Catch expected As InvalidOperationException
                    Assert.IsTrue(True)
                End Try

            End Sub

            <Test()> Public Sub 不変オブジェクトは_IsTypeImmutableをTrueにしてから_beforeVosを指定する()
                Dim testee As New ChangeJudgeer(Of ImmutableStringPair) With {.IsTypeImmutable = True}
                testee.SupersedeBeforeUpdatedVos(EzUtil.NewList(New ImmutableStringPair("A", "B")))
            End Sub

        End Class

        Public Class [Default] : Inherits ChangeJudgeerTest

            <Test()> Public Sub ExtractInsertedVosした値を_NotifySaveEndでbeforeVoにしていたが_同一インスタンスとなり改変できてしまうので_そうならないように修正する()
                Dim vo1 As TestVo2 = NewTestVo2(1, 0, "A")
                Dim testee As New ChangeJudgeerByKey(EzUtil.NewList(vo1))
                Dim vo2 As TestVo2 = NewTestVo2(2, 0, "B")
                testee.SetUpdatedVos(EzUtil.NewList(vo1, vo2))

                Dim extractedVo2 As TestVo2 = testee.ExtractInsertedVos(0)
                Assert.That(extractedVo2.Name, [Is].EqualTo("B"), "追加分としてvo2の内容が取れる")

                testee.NotifySaveEnd()

                vo2.Name = "B+"
                extractedVo2.Name = "B+"
                testee.SetUpdatedVos(EzUtil.NewList(vo1, vo2))

                Assert.That(testee.WasDeleted, [Is].False)
                Assert.That(testee.WasInserted, [Is].False)
                Assert.That(testee.WasUpdated, [Is].True)
                Assert.That(testee.ExtractUpdatedVos(0).BeforeVo.Name, [Is].EqualTo("B"), "extractedVo2.NameはB+に変更したけど、BeforeはBのまま")
                Assert.That(testee.ExtractUpdatedVos(0).AfterVo.Name, [Is].EqualTo("B+"))
            End Sub

        End Class

        Public Class 配列の中身がプリミティブ値 : Inherits ChangeJudgeerTest

            <Test()> Public Sub Integer等のプリミティブ値が要素だと_インスタンスで判断できないから_変更_しても追加削除になる()
                Dim values As Integer() = {1, 2, 3}
                Dim sut As New ChangeJudgeer(Of Integer)(values)
                values(1) = 9
                sut.SetUpdatedVos(values)
                Assert.That(sut.WasDeleted, [Is].True)
                Assert.That(sut.WasUpdated, [Is].False)
                Assert.That(sut.WasInserted, [Is].True)
                Assert.That(sut.ExtractDeletedVos(0), [Is].EqualTo(2))
                Assert.That(sut.ExtractInsertedVos(0), [Is].EqualTo(9))
            End Sub

            <Test()> Public Sub Integerの配列を要素にすれば_インスタンスで判断できる()
                Dim values As Integer() = {1, 2, 3}
                Dim sut As New ChangeJudgeer(Of Integer())(values)
                values(1) = 9
                sut.SetUpdatedVos(values)
                Assert.That(sut.WasDeleted, [Is].False)
                Assert.That(sut.WasUpdated, [Is].True)
                Assert.That(sut.WasInserted, [Is].False)
                Dim infos As ChangeJudgeer(Of Integer()).UpdateInfo() = sut.ExtractUpdatedVos
                Assert.That(infos(0).BeforeVo, [Is].EquivalentTo(New Integer() {1, 2, 3}))
                Assert.That(infos(0).AfterVo, [Is].EquivalentTo(New Integer() {1, 9, 3}))
                Assert.That(infos.Length, [Is].EqualTo(1))
            End Sub

            <Test()> Public Sub Stringの配列を要素にすれば_インスタンスで判断できる()
                Dim values As String() = {"a", "b", "c"}
                Dim sut As New ChangeJudgeer(Of String())(values)
                values(1) = "z"
                sut.SetUpdatedVos(values)
                Assert.That(sut.WasDeleted, [Is].False)
                Assert.That(sut.WasUpdated, [Is].True)
                Assert.That(sut.WasInserted, [Is].False)
                Dim infos As ChangeJudgeer(Of String()).UpdateInfo() = sut.ExtractUpdatedVos
                Assert.That(infos(0).BeforeVo, [Is].EquivalentTo(New String() {"a", "b", "c"}))
                Assert.That(infos(0).AfterVo, [Is].EquivalentTo(New String() {"a", "c", "z"}))
                Assert.That(infos.Length, [Is].EqualTo(1))
            End Sub

            <Test()> Public Sub Stringの配列を要素にすれば_インスタンスで判断できる_2次元配列()
                Dim data As String()() = {New String() {"a", "b", "c"}, New String() {"l", "m", "n"}, New String() {"x", "y", "z"}}
                Dim sut As New ChangeJudgeer(Of String())(data)
                data(1)(2) = "1"
                sut.SetUpdatedVos(data)
                Assert.That(sut.WasDeleted, [Is].False)
                Assert.That(sut.WasUpdated, [Is].True)
                Assert.That(sut.WasInserted, [Is].False)
                Dim infos As ChangeJudgeer(Of String()).UpdateInfo() = sut.ExtractUpdatedVos
                Assert.That(infos(0).BeforeVo, [Is].EquivalentTo(New String() {"l", "m", "n"}))
                Assert.That(infos(0).AfterVo, [Is].EquivalentTo(New String() {"l", "m", "1"}))
                Assert.That(infos.Length, [Is].EqualTo(1))
            End Sub

            <Test()> Public Sub Object型の配列を要素にすれば_インスタンスで判断できる()
                Dim values As Object() = {"a", "b", "c"}
                Dim sut As New ChangeJudgeer(Of Object())(values)
                values(1) = 2
                sut.SetUpdatedVos(values)
                Assert.That(sut.WasDeleted, [Is].False)
                Assert.That(sut.WasUpdated, [Is].True)
                Assert.That(sut.WasInserted, [Is].False)
                Dim infos As ChangeJudgeer(Of Object()).UpdateInfo() = sut.ExtractUpdatedVos
                Assert.That(infos(0).BeforeVo, [Is].EqualTo(New Object() {"a", "b", "c"}))
                Assert.That(infos(0).AfterVo, [Is].EqualTo(New Object() {"a", 2, "c"}))
                Assert.That(infos.Length, [Is].EqualTo(1))
            End Sub

            <Test()> Public Sub Object型の配列を要素にすれば_インスタンスで判断できる_2次元配列()
                Dim data As Object()() = {New Object() {"a", "b", "c"}, New Object() {"x", "y", "z"}}
                Dim sut As New ChangeJudgeer(Of Object())(data)
                data(1)(2) = 3
                sut.SetUpdatedVos(data)
                Assert.That(sut.WasDeleted, [Is].False)
                Assert.That(sut.WasUpdated, [Is].True)
                Assert.That(sut.WasInserted, [Is].False)
                Dim infos As ChangeJudgeer(Of Object()).UpdateInfo() = sut.ExtractUpdatedVos
                Assert.That(infos(0).BeforeVo, [Is].EqualTo(New Object() {"x", "y", "z"}))
                Assert.That(infos(0).AfterVo, [Is].EqualTo(New Object() {"x", "y", 3}))
                Assert.That(infos.Length, [Is].EqualTo(1))
            End Sub

        End Class

        Public Class 値オブジェクトTest : Inherits ChangeJudgeerTest

            <Test()> Public Sub 同じValueObjectなら_変化なし_となる_Listのインスタンスが変わっても()
                Dim value As TestingValueObject = New TestingValueObject("n1", "z1", "a1")
                Dim before As New List(Of TestingValueObject)({value})
                Dim sut As New ChangeJudgeer(Of TestingValueObject)(before)

                Dim after As New List(Of TestingValueObject)({value})
                sut.SetUpdatedVos(after)

                Assert.That(sut.HasChanged, [Is].False)
            End Sub

            <Test()> Public Sub インスタンス違いのValueObjecでも_中身が同じなら_同値なので_変化なし_となる()
                Dim value1 As TestingValueObject = New TestingValueObject("n1", "z1", "a1")
                Dim value2 As TestingValueObject = New TestingValueObject("n1", "z1", "a1")
                Dim collection As New List(Of TestingValueObject)({value1})
                Dim sut As New ChangeJudgeer(Of TestingValueObject)(collection)

                collection(0) = value2
                sut.SetUpdatedVos(collection)

                Assert.That(sut.HasChanged, [Is].False)
            End Sub

            <Test()> Public Sub 中身が違うValueObjectなら_同値じゃないので_追加あり削除あり_となる_Listのインスタンス同じ()
                Dim value1 As TestingValueObject = New TestingValueObject("n1", "z1", "a1")
                Dim value2 As TestingValueObject = New TestingValueObject("n1", "z2", "a1")
                Dim collection As New List(Of TestingValueObject)({value1})
                Dim sut As New ChangeJudgeer(Of TestingValueObject)(collection)

                collection(0) = value2
                sut.SetUpdatedVos(collection)

                Assert.That(sut.HasChanged, [Is].True)
                Assert.That(sut.WasDeleted, [Is].True)
                Assert.That(sut.WasInserted, [Is].True)
                Assert.That(sut.WasUpdated, [Is].False, "同値じゃないので追加/削除になる")
                Assert.That(sut.ExtractDeletedVos(0), [Is].EqualTo(value1))
                Assert.That(sut.ExtractInsertedVos(0), [Is].EqualTo(value2))
            End Sub

            <Test()> Public Sub 中身が違うValueObjectなら_同値じゃないので_追加あり削除あり_となる_Listのインスタンス違い()
                Dim value1 As TestingValueObject = New TestingValueObject("n1", "z1", "a1")
                Dim value2 As TestingValueObject = New TestingValueObject("n1", "z2", "a1")
                Dim before As New List(Of TestingValueObject)({value1})
                Dim sut As New ChangeJudgeer(Of TestingValueObject)(before)

                Dim after As New List(Of TestingValueObject)({value2})
                sut.SetUpdatedVos(after)

                Assert.That(sut.HasChanged, [Is].True)
                Assert.That(sut.WasDeleted, [Is].True)
                Assert.That(sut.WasInserted, [Is].True)
                Assert.That(sut.WasUpdated, [Is].False, "同値じゃないので追加/削除になる")
                Assert.That(sut.ExtractDeletedVos(0), [Is].EqualTo(value1))
                Assert.That(sut.ExtractInsertedVos(0), [Is].EqualTo(value2))
            End Sub

        End Class

        Public Class 無視属性Test : Inherits ChangeJudgeerTest

            <Test()> Public Sub 値を変更すれば_変化点ありになる()
                Dim vos As AttrVo() = {New AttrVo With {.Id = 2, .Name = "A"},
                                       New AttrVo With {.Id = 3, .Name = "A"}}
                Dim sut As New ChangeJudgeer(Of AttrVo)(vos)
                vos(0).Name = "B"
                vos(1).Name = "C"
                sut.SetUpdatedVos(vos)

                Assert.That(sut.HasChanged, [Is].True)
            End Sub

            <Test()> Public Sub 値を変更しても_そのプロパティの属性を無視属性に指定したら_変化点なしになる()
                Dim vos As AttrVo() = {New AttrVo With {.Id = 2, .Name = "A"},
                                       New AttrVo With {.Id = 3, .Name = "A"}}
                Dim sut As New ChangeJudgeer(Of AttrVo)(vos) With {.IgnorePropertyAttributes = {GetType(TestingNameAttribute)}}
                vos(0).Name = "B"
                vos(1).Name = "C"
                sut.SetUpdatedVos(vos)

                Assert.That(sut.HasChanged, [Is].False, "Nameは無視するので変化点なし")
            End Sub

        End Class

        Public Class 含める属性Test_インスタンス値差分 : Inherits ChangeJudgeerTest

            <Test()> Public Sub インスタンス値の値を変更しても_変化点はありにならない()
                Dim vos As AttrParentVo() = {New AttrParentVo With {.Id = 2, .Vo = New AttrVo With {.Name = "A"}},
                                             New AttrParentVo With {.Id = 3, .Vo = New AttrVo With {.Name = "A"}}}
                Dim sut As New ChangeJudgeer(Of AttrParentVo)(vos)
                vos(0).Vo.Name = "B"
                vos(1).Vo.Name = "C"
                sut.SetUpdatedVos(vos)

                Assert.That(sut.HasChanged, [Is].False)
            End Sub

            <Test()> Public Sub 含める属性を指定したら_インスタンス値の値を変更しても_変化点ありになる()
                Dim vos As AttrParentVo() = {New AttrParentVo With {.Id = 2, .Vo = New AttrVo With {.Name = "A"}},
                                             New AttrParentVo With {.Id = 3, .Vo = New AttrVo With {.Name = "A"}}}
                Dim sut As New ChangeJudgeer(Of AttrParentVo)(vos) With {.IncludePropertyAttributes = {GetType(TestingNameAttribute)}}
                vos(0).Vo.Name = "B"
                vos(1).Vo.Name = "C"
                sut.SetUpdatedVos(vos)

                Assert.That(sut.HasChanged, [Is].True)
            End Sub

        End Class

    End Class
End Namespace