Imports Fhi.Fw.Db
Imports NUnit.Framework

Namespace Util
    Public MustInherit Class DecamelizedPropertyAccessorTest

        Private Class GetterOnlyVo
            Public ReadOnly Property HogeStr As String
                Get
                    Return "Str"
                End Get
            End Property
            Public ReadOnly Property HogeInt As Integer
                Get
                    Return 123
                End Get
            End Property
            Public ReadOnly Property HogeDec As Decimal
                Get
                    Return 123.45D
                End Get
            End Property
            Public ReadOnly Property HogeDt As Date
                Get
                    Return CDate("2021/11/11 12:34:56")
                End Get
            End Property
            Public ReadOnly Property HogeEnum As SampleEnum
                Get
                    Return SampleEnum.A
                End Get
            End Property
        End Class
        Private Class SetterOnlyVo
            Public strVal As String
            Public intVal As Integer
            Public decVal As Decimal
            Public dtVal As Date
            Public enumVal As SampleEnum

            Public WriteOnly Property FugaStr As String
                Set(value As String)
                    strVal = value
                End Set
            End Property
            Public WriteOnly Property FugaInt As Integer
                Set(value As Integer)
                    intVal = value
                End Set
            End Property
            Public WriteOnly Property FugaDec As Decimal
                Set(value As Decimal)
                    decVal = value
                End Set
            End Property
            Public WriteOnly Property FugaDt As Date
                Set(value As Date)
                    dtVal = value
                End Set
            End Property
            Public WriteOnly Property FugaEnum As SampleEnum
                Set(value As SampleEnum)
                    enumVal = value
                End Set
            End Property
        End Class

        Private accessor As DecamelizedPropertyAccessor

        <SetUp()> Public Sub SetUp()
            accessor = New DecamelizedPropertyAccessor
        End Sub

        Public Class GetPropertyInfoTest : Inherits DecamelizedPropertyAccessorTest

            <Test()> Public Sub デキャメライズ名でPropertyInfoを取得()
                Dim vo As New SampleVo
                Assert.That(accessor.GetPropertyInfo(vo, "HOGE_ID").Name, [Is].EqualTo("HogeId"))
                Assert.That(accessor.GetPropertyInfo(vo, "HOGE_NAME").Name, [Is].EqualTo("HogeName"))
                Assert.That(accessor.GetPropertyInfo(vo, "HOGE_DECIMAL").Name, [Is].EqualTo("HogeDecimal"))
                Assert.That(accessor.GetPropertyInfo(vo, "HOGE_DATE").Name, [Is].EqualTo("HogeDate"))
            End Sub

            <Test()> Public Sub Getterのみのプロパティは取得できない()
                Dim vo As New GetterOnlyVo
                Try
                    Dim name As String = accessor.GetPropertyInfo(vo, "HOGE_STR").Name
                    Assert.Fail()
                Catch ex As KeyNotFoundException
                    Assert.That(ex.Message, [Is].EqualTo("指定されたキーはディレクトリ内に存在しませんでした。"))
                End Try
            End Sub

            <Test()> Public Sub Setterのみのプロパティは取得できない()
                Dim vo As New SetterOnlyVo
                Try
                    Dim name As String = accessor.GetPropertyInfo(vo, "FUGA_STR").Name
                    Assert.Fail()
                Catch ex As KeyNotFoundException
                    Assert.That(ex.Message, [Is].EqualTo("指定されたキーはディレクトリ内に存在しませんでした。"))
                End Try
            End Sub
        End Class

        Public Class GetGetPropertyInfoTest : Inherits DecamelizedPropertyAccessorTest

            <Test()> Public Sub デキャメライズ名でPropertyInfoを取得できる()
                Dim vo As New SampleVo
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_ID").Name, [Is].EqualTo("HogeId"))
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_NAME").Name, [Is].EqualTo("HogeName"))
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_DECIMAL").Name, [Is].EqualTo("HogeDecimal"))
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_DATE").Name, [Is].EqualTo("HogeDate"))
            End Sub

            <Test()> Public Sub デキャメライズ名でReadOnlyのPropertyInfoを取得できる()
                Dim vo As New GetterOnlyVo
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_STR").Name, [Is].EqualTo("HogeStr"))
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_INT").Name, [Is].EqualTo("HogeInt"))
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_DEC").Name, [Is].EqualTo("HogeDec"))
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_DT").Name, [Is].EqualTo("HogeDt"))
                Assert.That(accessor.GetGetPropertyInfo(vo, "HOGE_ENUM").Name, [Is].EqualTo("HogeEnum"))
            End Sub

            <Test()> Public Sub WriteOnlyPropertyは取得できない()
                Dim vo As New SetterOnlyVo
                Try
                    Dim name As String = accessor.GetGetPropertyInfo(vo, "FUGA_STR").Name
                    Assert.Fail()
                Catch ex As KeyNotFoundException
                    Assert.That(ex.Message, [Is].EqualTo("指定されたキーはディレクトリ内に存在しませんでした。"))
                End Try
            End Sub
        End Class

        Public Class GetSetPropertyInfoTest : Inherits DecamelizedPropertyAccessorTest

            <Test()> Public Sub GetSetPropertyInfo_デキャメライズ名でPropertyInfoを取得できる()
                Dim vo As New SampleVo
                Assert.That(accessor.GetSetPropertyInfo(vo, "HOGE_ID").Name, [Is].EqualTo("HogeId"))
                Assert.That(accessor.GetSetPropertyInfo(vo, "HOGE_NAME").Name, [Is].EqualTo("HogeName"))
                Assert.That(accessor.GetSetPropertyInfo(vo, "HOGE_DECIMAL").Name, [Is].EqualTo("HogeDecimal"))
                Assert.That(accessor.GetSetPropertyInfo(vo, "HOGE_DATE").Name, [Is].EqualTo("HogeDate"))
            End Sub

            <Test()> Public Sub GetSetPropertyInfo_デキャメライズ名でWriteOnlyのPropertyInfoを取得できる()
                Dim vo As New SetterOnlyVo
                Assert.That(accessor.GetSetPropertyInfo(vo, "FUGA_STR").Name, [Is].EqualTo("FugaStr"))
                Assert.That(accessor.GetSetPropertyInfo(vo, "FUGA_INT").Name, [Is].EqualTo("FugaInt"))
                Assert.That(accessor.GetSetPropertyInfo(vo, "FUGA_DEC").Name, [Is].EqualTo("FugaDec"))
                Assert.That(accessor.GetSetPropertyInfo(vo, "FUGA_DT").Name, [Is].EqualTo("FugaDt"))
                Assert.That(accessor.GetSetPropertyInfo(vo, "FUGA_ENUM").Name, [Is].EqualTo("FugaEnum"))
            End Sub

            <Test()> Public Sub GetSetPropertyInfo_ReadOnlyPropertyは取得できない()
                Dim vo As New GetterOnlyVo
                Try
                    Dim name As String = accessor.GetSetPropertyInfo(vo, "HOGE_STR").Name
                    Assert.Fail()
                Catch ex As KeyNotFoundException
                    Assert.That(ex.Message, [Is].EqualTo("指定されたキーはディレクトリ内に存在しませんでした。"))
                End Try
            End Sub
        End Class

        Public Class ContainsTest : Inherits DecamelizedPropertyAccessorTest

            <Test()> Public Sub プロパティ名を含む場合true()
                Dim vo As New SampleVo
                Assert.True(accessor.Contains(vo, "HOGE_ID"))
                Assert.True(accessor.Contains(vo, "HOGE_NAME"))
                Assert.True(accessor.Contains(vo, "HOGE_DECIMAL"))
                Assert.True(accessor.Contains(vo, "HOGE_DATE"))
            End Sub

            <Test()> Public Sub 存在しないプロパティ名ならfalse()
                Dim vo As New SampleVo
                Assert.False(accessor.Contains(vo, "HOGE_SINGLE"))
                Assert.False(accessor.Contains(vo, "CLASS"))
            End Sub

            <Test()> Public Sub ReadOnlyのプロパティ名なら_あってもFalse()
                Dim vo As New GetterOnlyVo
                Assert.False(accessor.Contains(vo, "HOGE_STR"))
                Assert.False(accessor.Contains(vo, "HOGE_INT"))
                Assert.False(accessor.Contains(vo, "HOGE_DEC"))
                Assert.False(accessor.Contains(vo, "HOGE_DT"))
                Assert.False(accessor.Contains(vo, "HOGE_ENUM"))
            End Sub

            <Test()> Public Sub WriteOnlyのプロパティ名なら_あってもFalse()
                Dim vo As New GetterOnlyVo
                Assert.False(accessor.Contains(vo, "FUGA_STR"))
                Assert.False(accessor.Contains(vo, "FUGA_INT"))
                Assert.False(accessor.Contains(vo, "FUGA_DEC"))
                Assert.False(accessor.Contains(vo, "FUGA_DT"))
                Assert.False(accessor.Contains(vo, "FUGA_ENUM"))
            End Sub
        End Class

        Public Class ContainsGetterTest : Inherits DecamelizedPropertyAccessorTest

            <Test()> Public Sub プロパティ名を含む場合_True()
                Dim vo As New SampleVo
                Assert.True(accessor.ContainsGetter(vo, "HOGE_ID"))
                Assert.True(accessor.ContainsGetter(vo, "HOGE_NAME"))
                Assert.True(accessor.ContainsGetter(vo, "HOGE_DECIMAL"))
                Assert.True(accessor.ContainsGetter(vo, "HOGE_DATE"))
            End Sub

            <Test()> Public Sub ReadOnlyのプロパティ名を含む場合_True()
                Dim vo As New GetterOnlyVo
                Assert.True(accessor.ContainsGetter(vo, "HOGE_STR"))
                Assert.True(accessor.ContainsGetter(vo, "HOGE_INT"))
                Assert.True(accessor.ContainsGetter(vo, "HOGE_DEC"))
                Assert.True(accessor.ContainsGetter(vo, "HOGE_DT"))
                Assert.True(accessor.ContainsGetter(vo, "HOGE_ENUM"))
            End Sub

            <Test()> Public Sub 存在しないプロパティ名なら_False()
                Dim vo As New GetterOnlyVo
                Assert.False(accessor.ContainsGetter(vo, "HOGE_SINGLE"))
                Assert.False(accessor.ContainsGetter(vo, "CLASS"))
            End Sub

            <Test()> Public Sub WriteOnlyプロパティ名なら_False()
                Dim vo As New SetterOnlyVo
                Assert.False(accessor.ContainsGetter(vo, "FUGA_STR"))
                Assert.False(accessor.ContainsGetter(vo, "FUGA_INT"))
                Assert.False(accessor.ContainsGetter(vo, "FUGA_DEC"))
                Assert.False(accessor.ContainsGetter(vo, "FUGA_DT"))
                Assert.False(accessor.ContainsGetter(vo, "FUGA_ENUM"))
            End Sub
        End Class

        Public Class ContainsSetterTest : Inherits DecamelizedPropertyAccessorTest

            <Test()> Public Sub プロパティ名を含む場合_True()
                Dim vo As New SampleVo
                Assert.True(accessor.ContainsSetter(vo, "HOGE_ID"))
                Assert.True(accessor.ContainsSetter(vo, "HOGE_NAME"))
                Assert.True(accessor.ContainsSetter(vo, "HOGE_DECIMAL"))
                Assert.True(accessor.ContainsSetter(vo, "HOGE_DATE"))
            End Sub

            <Test()> Public Sub WriteOnlyのプロパティ名を含む場合_True()
                Dim vo As New SetterOnlyVo
                Assert.True(accessor.ContainsSetter(vo, "FUGA_STR"))
                Assert.True(accessor.ContainsSetter(vo, "FUGA_INT"))
                Assert.True(accessor.ContainsSetter(vo, "FUGA_DEC"))
                Assert.True(accessor.ContainsSetter(vo, "FUGA_DT"))
                Assert.True(accessor.ContainsSetter(vo, "FUGA_ENUM"))
            End Sub

            <Test()> Public Sub 存在しないプロパティ名なら_False()
                Dim vo As New SetterOnlyVo
                Assert.False(accessor.ContainsSetter(vo, "FUGA_SINGLE"))
                Assert.False(accessor.ContainsSetter(vo, "CLASS"))
            End Sub

            <Test()> Public Sub ReadOnlyプロパティ名なら_False()
                Dim vo As New GetterOnlyVo
                Assert.False(accessor.ContainsSetter(vo, "HOGE_STR"))
                Assert.False(accessor.ContainsSetter(vo, "HOGE_INT"))
                Assert.False(accessor.ContainsSetter(vo, "HOGE_DEC"))
                Assert.False(accessor.ContainsSetter(vo, "HOGE_DT"))
                Assert.False(accessor.ContainsSetter(vo, "HOGE_ENUM"))
            End Sub
        End Class

        Public Class GetValueTest : Inherits DecamelizedPropertyAccessorTest

            <Test()> Public Sub デキャメライズ名で値を取得できる()
                Dim vo As New SampleVo With {.HogeId = 1, .HogeName = "name", .HogeDecimal = 123.456D, .HogeDate = CDate("2012/03/03 12:34:56"), .HogeEnum = SampleEnum.A}
                Assert.That(accessor.GetValue(vo, "HOGE_ID", Nothing), [Is].EqualTo(1))
                Assert.That(accessor.GetValue(vo, "HOGE_NAME", Nothing), [Is].EqualTo("name"))
                Assert.That(accessor.GetValue(vo, "HOGE_DECIMAL", Nothing), [Is].EqualTo(123.456D))
                Assert.That(accessor.GetValue(vo, "HOGE_DATE", Nothing), [Is].EqualTo(CDate("2012/03/03 12:34:56")))
                Assert.That(accessor.GetValue(vo, "HOGE_ENUM", Nothing), [Is].EqualTo(SampleEnum.A))
            End Sub

            <Test()> Public Sub デキャメライズ名で値を取得できる_ReadOnlyプロパティでも可能()
                Dim vo As New GetterOnlyVo
                Assert.That(accessor.GetValue(vo, "HOGE_STR", Nothing), [Is].EqualTo("Str"))
                Assert.That(accessor.GetValue(vo, "HOGE_INT", Nothing), [Is].EqualTo(123))
                Assert.That(accessor.GetValue(vo, "HOGE_DEC", Nothing), [Is].EqualTo(123.45D))
                Assert.That(accessor.GetValue(vo, "HOGE_DT", Nothing), [Is].EqualTo(CDate("2021/11/11 12:34:56")))
                Assert.That(accessor.GetValue(vo, "HOGE_ENUM", Nothing), [Is].EqualTo(SampleEnum.A))
            End Sub
        End Class

        Public Class SetValueTest : Inherits DecamelizedPropertyAccessorTest

            <Test()> Public Sub デキャメライズ名で値を設定する()
                Dim vo As New SampleVo With {.HogeId = 1, .HogeName = "name", .HogeDecimal = 123.456D, .HogeDate = CDate("2012/03/03 12:34:56"), .HogeEnum = SampleEnum.B}
                accessor.SetValue(vo, "HOGE_ID", 2, Nothing)
                accessor.SetValue(vo, "HOGE_NAME", "second", Nothing)
                accessor.SetValue(vo, "HOGE_DECIMAL", 987.654D, Nothing)
                accessor.SetValue(vo, "HOGE_DATE", CDate("2012/04/04 01:23:45"), Nothing)
                accessor.SetValue(vo, "HOGE_ENUM", SampleEnum.B, Nothing)

                Assert.That(accessor.GetValue(vo, "HOGE_ID", Nothing), [Is].EqualTo(2))
                Assert.That(accessor.GetValue(vo, "HOGE_NAME", Nothing), [Is].EqualTo("second"))
                Assert.That(accessor.GetValue(vo, "HOGE_DECIMAL", Nothing), [Is].EqualTo(987.654D))
                Assert.That(accessor.GetValue(vo, "HOGE_DATE", Nothing), [Is].EqualTo(CDate("2012/04/04 01:23:45")))
                Assert.That(accessor.GetValue(vo, "HOGE_ENUM", Nothing), [Is].EqualTo(SampleEnum.B))
            End Sub

            <Test()> Public Sub デキャメライズ名で値を設定する_WriteOnlyプロパティでも可()
                Dim vo As New SetterOnlyVo
                accessor.SetValue(vo, "FUGA_STR", "moji", Nothing)
                accessor.SetValue(vo, "FUGA_INT", 321, Nothing)
                accessor.SetValue(vo, "FUGA_DEC", 33.4D, Nothing)
                accessor.SetValue(vo, "FUGA_DT", CDate("2021/11/10 10:20:30"), Nothing)
                accessor.SetValue(vo, "FUGA_ENUM", SampleEnum.B, Nothing)

                Assert.That(vo.strVal, [Is].EqualTo("moji"))
                Assert.That(vo.intVal, [Is].EqualTo(321))
                Assert.That(vo.decVal, [Is].EqualTo(33.4D))
                Assert.That(vo.dtVal, [Is].EqualTo(CDate("2021/11/10 10:20:30")))
                Assert.That(vo.enumVal, [Is].EqualTo(SampleEnum.B))
            End Sub

            <Test()> Public Sub 数値に空文字ならNothingを設定する()
                Dim vo As New SampleVo With {.HogeId = 1, .HogeName = "name", .HogeDecimal = 123.456D, .HogeDate = CDate("2012/03/03 12:34:56")}
                accessor.SetValue(vo, "HOGE_ID", "", Nothing)
                accessor.SetValue(vo, "HOGE_DECIMAL", "", Nothing)

                Assert.IsNull(accessor.GetValue(vo, "HOGE_ID", Nothing))
                Assert.IsNull(accessor.GetValue(vo, "HOGE_DECIMAL", Nothing))
            End Sub

            <Test()> Public Sub 数値に文字なら例外()
                Dim vo As New SampleVo With {.HogeId = 1, .HogeName = "name", .HogeDecimal = 123.456D, .HogeDate = CDate("2012/03/03 12:34:56")}
                Try
                    accessor.SetValue(vo, "HOGE_ID", "aa", Nothing)
                Catch expect As ArgumentException
                    Assert.IsTrue(True)
                End Try
            End Sub
        End Class

    End Class
End Namespace