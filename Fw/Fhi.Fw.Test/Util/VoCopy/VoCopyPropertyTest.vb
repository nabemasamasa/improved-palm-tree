Imports NUnit.Framework

Namespace Util.VoCopy
    Public MustInherit Class VoCopyPropertyTest

#Region "Nested classes..."
        Private Class HogeVo
            Private _id As Integer?
            Private _name As String
            Private _birthDay As DateTime?
            Private _isMale As Boolean?

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
                    Return _name
                End Get
                Set(ByVal value As String)
                    _name = value
                End Set
            End Property

            Public Property BirthDay() As Date?
                Get
                    Return _birthDay
                End Get
                Set(ByVal value As Date?)
                    _birthDay = value
                End Set
            End Property

            Public Property IsMale() As Boolean?
                Get
                    Return _isMale
                End Get
                Set(ByVal value As Boolean?)
                    _isMale = value
                End Set
            End Property
        End Class
        Private Class FugaVo
            Private _bd As DateTime?
            Private _m As Boolean?
            Private _nm As String
            Private _i As Integer?

            Public Property Bd() As Date?
                Get
                    Return _bd
                End Get
                Set(ByVal value As Date?)
                    _bd = value
                End Set
            End Property

            Public Property MProperty() As Boolean?
                Get
                    Return _m
                End Get
                Set(ByVal value As Boolean?)
                    _m = value
                End Set
            End Property

            Public Property Nm() As String
                Get
                    Return _nm
                End Get
                Set(ByVal value As String)
                    _nm = value
                End Set
            End Property

            Public Property IProperty() As Integer?
                Get
                    Return _i
                End Get
                Set(ByVal value As Integer?)
                    _i = value
                End Set
            End Property
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()

        End Sub

        Public Class CopyPropetiesTest : Inherits VoCopyPropertyTest

            <Test()> Public Sub XからYへコピーできる()
                Dim sut As New VoCopyProperty(Of HogeVo, FugaVo)(Function(define, x, y) define.Bind(x.Id, y.IProperty).Bind(x.Name, y.Nm).Bind(x.BirthDay, y.Bd).Bind(x.IsMale, y.MProperty))
                Dim xVo As New HogeVo With {.Id = 3, .Name = "Piyo", .BirthDay = CDate("2001/01/01"), .IsMale = True}
                Dim yVo As New FugaVo
                sut.CopyProperties(xVo, yVo)

                Assert.That(yVo.IProperty, [Is].EqualTo(3))
                Assert.That(yVo.Nm, [Is].EqualTo("Piyo"))
                Assert.That(yVo.Bd, [Is].EqualTo(CDate("2001/01/01")))
                Assert.That(yVo.MProperty, [Is].True)

            End Sub

            <Test()> Public Sub YからXへコピーできる()
                Dim sut As New VoCopyProperty(Of HogeVo, FugaVo)(Function(define, x, y) define.Bind(x.Id, y.IProperty).Bind(x.Name, y.Nm).Bind(x.BirthDay, y.Bd).Bind(x.IsMale, y.MProperty))
                Dim fuga As New FugaVo With {.IProperty = 3, .nm = "Piyo", .bd = CDate("2001/01/01"), .mProperty = True}
                Dim hoge As New HogeVo
                sut.CopyProperties(fuga, hoge)

                Assert.That(hoge.Id, [Is].EqualTo(3))
                Assert.That(hoge.Name, [Is].EqualTo("Piyo"))
                Assert.That(hoge.BirthDay, [Is].EqualTo(CDate("2001/01/01")))
                Assert.That(hoge.IsMale, [Is].True)

            End Sub

            <Test()> Public Sub 定義の一部が逆でもコピーできる()
                Dim sut As New VoCopyProperty(Of HogeVo, FugaVo)(Function(define, x, y) define.Bind(x.Id, y.IProperty).Bind(y.Nm, x.Name).Bind(x.BirthDay, y.Bd).Bind(x.IsMale, y.MProperty))
                Dim xVo As New HogeVo With {.Id = 3, .Name = "Piyo", .BirthDay = CDate("2001/01/01"), .IsMale = True}
                Dim yVo As New FugaVo
                sut.CopyProperties(xVo, yVo)

                Assert.That(yVo.IProperty, [Is].EqualTo(3))
                Assert.That(yVo.Nm, [Is].EqualTo("Piyo"))
                Assert.That(yVo.Bd, [Is].EqualTo(CDate("2001/01/01")))
                Assert.That(yVo.MProperty, [Is].True)

            End Sub

            <Test()> Public Sub 定義が順不同でもコピーできる()
                Dim sut As New VoCopyProperty(Of HogeVo, FugaVo)(Function(define, x, y) define.Bind(y.Bd, x.BirthDay).Bind(x.Id, y.IProperty).Bind(x.IsMale, y.MProperty).Bind(y.Nm, x.Name))
                Dim xVo As New HogeVo With {.Id = 3, .Name = "Piyo", .BirthDay = CDate("2001/01/01"), .IsMale = True}
                Dim yVo As New FugaVo
                sut.CopyProperties(xVo, yVo)

                Assert.That(yVo.IProperty, [Is].EqualTo(3))
                Assert.That(yVo.Nm, [Is].EqualTo("Piyo"))
                Assert.That(yVo.Bd, [Is].EqualTo(CDate("2001/01/01")))
                Assert.That(yVo.MProperty, [Is].True)

            End Sub

        End Class

    End Class
End Namespace