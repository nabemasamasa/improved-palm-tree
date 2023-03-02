Imports Fhi.Fw.Domain
Imports Fhi.Fw.TestUtil.DebugString
Imports Fhi.Fw.Util
Imports NUnit.Framework

Namespace Db
    Public MustInherit Class DbAccessHelperTest

        Private Shared Function BuildTable(Of T)(name As String, values As Object()) As DataTable
            Dim result As DataTable = New DataTable
            AddColumnToTable(Of T)(result, name, values)
            Return result
        End Function

        Private Shared Sub AddColumnToTable(Of T)(dt As DataTable, columnName As String, values As Object())
            dt.Columns.Add(columnName, GetType(T))
            If dt.Rows.Count = 0 Then
                For Each value As Object In values
                    dt.Rows.Add(value)
                Next
            Else
                For i = 0 To dt.Rows.Count - 1
                    dt.Rows(i)(columnName) = values(i)
                Next
            End If
        End Sub

        Public Class [Default] : Inherits DbAccessHelperTest

            <Test()> Public Sub ConvDataTableToList_値型に変換_Integer()
                Dim dt As DataTable = BuildTable(Of Integer)("one", {1000, 9})

                Dim actuals As List(Of Integer) = DbAccessHelper.ConvDataTableToList(Of Integer)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual(1000, actuals(0))
                Assert.AreEqual(9, actuals(1))
            End Sub

            <Test()> Public Sub ConvDataTableToList_結果nullをNullable型に変換_nullで取得()
                Dim dt As DataTable = BuildTable(Of Integer)("one", {Nothing})

                Dim actuals As List(Of Integer?) = DbAccessHelper.ConvDataTableToList(Of Integer?)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.IsNull(actuals(0))
            End Sub

            <Test()> Public Sub ConvDataTableToList_結果nullを値型に変換_値型の初期値になる()
                Dim dt As DataTable = BuildTable(Of Integer)("one", {Nothing})

                Dim actuals As List(Of Integer) = DbAccessHelper.ConvDataTableToList(Of Integer)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(0, actuals(0))
            End Sub

            <Test()> Public Sub ConvDataTableToList_Stringに変換()
                Dim dt As DataTable = BuildTable(Of String)("one", {"hoge", "fuga"})

                Dim actuals As List(Of String) = DbAccessHelper.ConvDataTableToList(Of String)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("hoge", actuals(0))
                Assert.AreEqual("fuga", actuals(1))
            End Sub

            <Test()> Public Sub ConvDataTableToList_値型またはString型に変換する場合_DataTableの列名は何でも同じ()
                Dim dt As DataTable = BuildTable(Of Integer)("fugafugahogehogefuga", {999})

                Dim actuals As List(Of Integer) = DbAccessHelper.ConvDataTableToList(Of Integer)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(999, actuals(0))
            End Sub

            <Test()> Public Sub ConvDataTableToList_値型またはString型に変換する場合_１列目を取得する()
                Dim dt As DataTable = BuildTable(Of Integer)("String", {888})
                AddColumnToTable(Of Integer)(dt, "Integer", {99})

                Dim actuals As List(Of Integer) = DbAccessHelper.ConvDataTableToList(Of Integer)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(888, actuals(0), "複数列あっても1列目の値")
            End Sub

        End Class

        Public Class ConvDataTableToList_戻り値Vo : Inherits DbAccessHelperTest

            <Test()> Public Sub ConvDataTableToList_VOに変換()
                Dim dt As DataTable = BuildTable(Of Integer)("HOGE_ID", {3})
                AddColumnToTable(Of DateTime)(dt, "HOGE_DATE", {DateTime.Parse("2010/02/03 04:55:06")})
                AddColumnToTable(Of Decimal)(dt, "HOGE_DECIMAL", {12.345@})
                AddColumnToTable(Of String)(dt, "HOGE_NAME", {"aiueo"})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(3, actuals(0).HogeId)
                Assert.AreEqual("aiueo", actuals(0).HogeName)
                Assert.AreEqual(12.345@, actuals(0).HogeDecimal)
                Assert.AreEqual(DateTime.Parse("2010/02/03 04:55:06"), actuals(0).HogeDate)
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_int_String()
                Dim dt As DataTable = BuildTable(Of Integer)("HOGE_NAME", {3})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual("3", actuals(0).HogeName)
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_String_int()
                Dim dt As DataTable = BuildTable(Of String)("HOGE_ID", {"3"})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(3, actuals(0).HogeId)
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_Decimal_int_小数無し()
                Dim dt As DataTable = BuildTable(Of Decimal)("HOGE_ID", {400@})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(400, actuals(0).HogeId)
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_Decimal_int_小数有りは例外()
                Dim dt As DataTable = BuildTable(Of Decimal)("HOGE_ID", {141.421356@})

                Try
                    DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                    Assert.Fail()
                Catch expect As ArgumentException
                    Assert.AreEqual("列名 HOGE_ID の値は 141.421356 だが、VO.HogeId は Int32 型", expect.Message)
                End Try
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_int_Decimal()
                Dim dt As DataTable = BuildTable(Of Integer)("HOGE_DECIMAL", {3})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(3@, actuals(0).HogeDecimal)
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_Double_Decimal()
                Dim dt As DataTable = BuildTable(Of Double)("HOGE_DECIMAL", {3.14#})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(3.14@, actuals(0).HogeDecimal)
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_Double_Decimal_精度()
                Dim dt As DataTable = BuildTable(Of Double)("HOGE_DECIMAL", {3.141592653589#})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(3.141592653589@, actuals(0).HogeDecimal)
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_String_Decimal()
                Dim dt As DataTable = BuildTable(Of String)("HOGE_DECIMAL", {"1.732"})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual(1.732@, actuals(0).HogeDecimal)
            End Sub

        End Class

        Public Class ConvDataTableToList_戻り値ValueObject : Inherits DbAccessHelperTest

            Private Class Id : Inherits PrimitiveValueObject(Of Integer)
                Public Sub New(ByVal value As Integer)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class Name : Inherits PrimitiveValueObject(Of String)
                Public Sub New(ByVal value As String)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class Date2 : Inherits PrimitiveValueObject(Of DateTime)
                Public Sub New(ByVal value As DateTime)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class Money : Inherits PrimitiveValueObject(Of Decimal)
                Public Sub New(ByVal value As Decimal)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class SubId : Inherits PrimitiveValueObject(Of Byte?)
                Public Sub New(ByVal value As Byte?)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class SampleVo
                Public Property HogeId As Id
                Public Property HogeName As Name
                Public Property HogeDate As Date2
                Public Property HogeDecimal As Money
                Public Property HogeEnum As SampleEnum
            End Class
            Private Overloads Function ToString(records As IEnumerable(Of SampleVo)) As String
                Dim maker As New DebugStringMaker(Of SampleVo)(
                    Function(defineBy As IDebugStringRuleBinder, vo As SampleVo) _
                        defineBy.Bind(vo.HogeId, vo.HogeName, vo.HogeDate, vo.HogeDecimal, vo.HogeEnum))
                Return maker.MakeString(records)
            End Function

            <Test()> Public Sub ConvDataTableToList_VOに変換()
                Dim dt As DataTable = BuildTable(Of Integer)("HOGE_ID", {3})
                AddColumnToTable(Of DateTime)(dt, "HOGE_DATE", {DateTime.Parse("2010/02/03 04:55:06")})
                AddColumnToTable(Of Decimal)(dt, "HOGE_DECIMAL", {12.345@})
                AddColumnToTable(Of String)(dt, "HOGE_NAME", {"aiueo"})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.That(ToString(actuals), [Is].EqualTo(
                            "HogeId HogeName HogeDate             HogeDecimal HogeEnum" & vbCrLf & _
                            "     3 'aiueo'  '2010/02/03 4:55:06'      12.345        0"))
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_int_String()
                Dim dt As DataTable = BuildTable(Of Integer)("HOGE_NAME", {3})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.That(ToString(actuals), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf & _
                            "null   '3'      null     null               0"))
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_String_int()
                Dim dt As DataTable = BuildTable(Of String)("HOGE_ID", {"3"})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.That(ToString(actuals), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf & _
                            "     3 null     null     null               0"))
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_Decimal_int_小数無し()
                Dim dt As DataTable = BuildTable(Of Decimal)("HOGE_ID", {400@})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.That(ToString(actuals), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf & _
                            "   400 null     null     null               0"))
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_Decimal_int_小数有りは例外()
                Dim dt As DataTable = BuildTable(Of Decimal)("HOGE_ID", {141.421356@})

                Try
                    DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                    Assert.Fail()
                Catch expect As ArgumentException
                    Assert.AreEqual("列名 HOGE_ID の値は 141.421356 だが、VO.HogeId は Int32 型", expect.Message)
                End Try
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_int_Decimal()
                Dim dt As DataTable = BuildTable(Of Integer)("HOGE_DECIMAL", {3})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.That(ToString(actuals), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf & _
                            "null   null     null               3        0"))
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_Double_Decimal()
                Dim dt As DataTable = BuildTable(Of Double)("HOGE_DECIMAL", {3.14#})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.That(ToString(actuals), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf & _
                            "null   null     null            3.14        0"))
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_Double_Decimal_精度()
                Dim dt As DataTable = BuildTable(Of Double)("HOGE_DECIMAL", {3.141592653589#})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.That(ToString(actuals), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf & _
                            "null   null     null         3.141…        0"))
            End Sub

            <Test()> Public Sub ConvDataTableToList_DataTable_VO_が_String_Decimal()
                Dim dt As DataTable = BuildTable(Of String)("HOGE_DECIMAL", {"1.732"})

                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt)
                Assert.That(ToString(actuals), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf & _
                            "null   null     null           1.732        0"))
            End Sub

            <Test()> Public Sub Int型をIntのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {123})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Id)(dt, Nothing)(0), [Is].EqualTo(New Id(123)))
            End Sub

            <Test()> Public Sub Int型をNullableByteのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {12})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of SubId)(dt, Nothing)(0), [Is].EqualTo(New SubId(12)))
            End Sub

        End Class

        Public Class ConvDataTableToList_戻り値ValueObject2 : Inherits DbAccessHelperTest

            Private Class IdNull : Inherits PrimitiveValueObject(Of Integer?)
                Public Sub New(ByVal value As Integer?)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class Name : Inherits PrimitiveValueObject(Of String)
                Public Sub New(ByVal value As String)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class DateNull : Inherits PrimitiveValueObject(Of DateTime?)
                Public Sub New(ByVal value As DateTime?)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class DecimalNull : Inherits PrimitiveValueObject(Of Decimal?)
                Public Sub New(ByVal value As Decimal?)
                    MyBase.New(value)
                End Sub
            End Class

            <Test()> Public Sub Int型をNullableIntのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {234})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of IdNull)(dt, Nothing)(0), [Is].EqualTo(New IdNull(234)))
            End Sub

            <Test()> Public Sub Null値をNullableIntのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {Nothing})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of IdNull)(dt, Nothing)(0), [Is].EqualTo(New IdNull(Nothing)))
            End Sub

            <Test()> Public Sub Null値をNullableIntのPrimitiveValueObject型で取得できる_DBNull()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {DBNull.Value})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of IdNull)(dt, Nothing)(0), [Is].EqualTo(New IdNull(Nothing)))
            End Sub

            <Test()> Public Sub String型をStringのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of String)("name", {"ai"})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Name)(dt, Nothing)(0), [Is].EqualTo(New Name("ai")))
            End Sub

            <Test()> Public Sub Null値をStringのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of String)("name", {Nothing})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Name)(dt, Nothing)(0), [Is].EqualTo(New Name(Nothing)))
            End Sub

            <Test()> Public Sub Null値をStringのPrimitiveValueObject型で取得できる_DBNull()
                Dim dt As DataTable = BuildTable(Of String)("name", {DBNull.Value})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Name)(dt, Nothing)(0), [Is].EqualTo(New Name(Nothing)))
            End Sub

            <Test()> Public Sub DateTime型をNullableDateTimeのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of DateTime)("Date", {#12/23/2012#})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of DateNull)(dt, Nothing)(0), [Is].EqualTo(New DateNull(#12/23/2012#)))
            End Sub

            <Test()> Public Sub Null値をNullableDateTimeのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of DateTime)("Date", {Nothing})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of DateNull)(dt, Nothing)(0), [Is].EqualTo(New DateNull(Nothing)))
            End Sub

            <Test()> Public Sub Null値をNullableDateTimeのPrimitiveValueObject型で取得できる_DBNull()
                Dim dt As DataTable = BuildTable(Of DateTime)("Date", {DBNull.Value})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of DateNull)(dt, Nothing)(0), [Is].EqualTo(New DateNull(Nothing)))
            End Sub

            <Test()> Public Sub Decimal型をNullableDecimalのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of Decimal)("money", {345D})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of DecimalNull)(dt, Nothing)(0), [Is].EqualTo(New DecimalNull(345D)))
            End Sub

            <Test()> Public Sub Null値をNullableDecimalのPrimitiveValueObject型で取得できる()
                Dim dt As DataTable = BuildTable(Of Decimal)("money", {Nothing})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of DecimalNull)(dt, Nothing)(0), [Is].EqualTo(New DecimalNull(Nothing)))
            End Sub

            <Test()> Public Sub Null値をNullableDecimalのPrimitiveValueObject型で取得できる_DBNull()
                Dim dt As DataTable = BuildTable(Of Decimal)("money", {DBNull.Value})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of DecimalNull)(dt, Nothing)(0), [Is].EqualTo(New DecimalNull(Nothing)))
            End Sub

        End Class

        Public Class trimsEndOfValueToForce_項目値の末尾を強制的にTrimする : Inherits DbAccessHelperTest

            Private dt As DataTable

            <SetUp()> Public Sub SetUp()
                dt = New DataTable
                dt.Columns.Add("HOGE_NAME", GetType(String))
                dt.Rows.Add("hoge  ")
                dt.Rows.Add("fuga  ")
                dt.Rows.Add("piyo  ")
            End Sub

            <Test()> Public Sub falseならTrimせず_そのまま()
                Dim actuals As List(Of String) = DbAccessHelper.ConvDataTableToList(Of String)(dt, Nothing, trimsEndOfValueToForce:=False)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(3, actuals.Count)
                Assert.AreEqual("hoge  ", actuals(0), "サンプル的に1件目だけ")
            End Sub

            <Test()> Public Sub 型がStringでもtrue指定ならTrimされる()
                Dim actuals As List(Of String) = DbAccessHelper.ConvDataTableToList(Of String)(dt, Nothing, trimsEndOfValueToForce:=True)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(3, actuals.Count)
                Assert.AreEqual("hoge", actuals(0))
                Assert.AreEqual("fuga", actuals(1))
                Assert.AreEqual("piyo", actuals(2))
            End Sub

            <Test()> Public Sub 型がVoでもtrue指定ならTrimされる()
                Dim actuals As List(Of SampleVo) = DbAccessHelper.ConvDataTableToList(Of SampleVo)(dt, Nothing, trimsEndOfValueToForce:=True)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(3, actuals.Count)
                Assert.AreEqual("hoge", actuals(0).HogeName)
                Assert.AreEqual("fuga", actuals(1).HogeName)
                Assert.AreEqual("piyo", actuals(2).HogeName)
            End Sub

        End Class

        Public Class ConvDataTableToList_VOではない型での取得_Test : Inherits DbAccessHelperTest

            <Test()> Public Sub Long型をLong型で取得できる()
                Dim dt As DataTable = BuildTable(Of Long)("id", {123L})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Long)(dt, Nothing)(0), [Is].EqualTo(123L))
            End Sub

            <Test()> Public Sub Long型をInt型で取得できる()
                Dim dt As DataTable = BuildTable(Of Long)("id", {123L})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Integer)(dt, Nothing)(0), [Is].EqualTo(123))
            End Sub

            <Test()> Public Sub Int型をInt型で取得できる()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {123})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Integer)(dt, Nothing)(0), [Is].EqualTo(123))
            End Sub

            <Test()> Public Sub Int型をLong型で取得できる()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {123})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Long)(dt, Nothing)(0), [Is].EqualTo(123L))
            End Sub

            <Test()> Public Sub Int型をString型で取得できる()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {123})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of String)(dt, Nothing)(0), [Is].EqualTo("123"))
            End Sub

            <Test()> Public Sub String型をInt型で取得できる()
                Dim dt As DataTable = BuildTable(Of String)("id", {"23"})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Integer)(dt, Nothing)(0), [Is].EqualTo(23))
            End Sub

            <Test()> Public Sub Int型をByte型で取得できる()
                Dim dt As DataTable = BuildTable(Of Integer)("id", {123})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Byte)(dt, Nothing)(0), [Is].EqualTo(123))
            End Sub

            <Test()> Public Sub Byte型をInt型で取得できる()
                Dim dt As DataTable = BuildTable(Of Byte)("id", {123})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of Integer)(dt, Nothing)(0), [Is].EqualTo(123))
            End Sub

        End Class

        Public Class 複数コンストラクタのPrimitiveValueObjectTest : Inherits VoPropertyMarkerTest

            Private Class HogeVo
                Public Property Fuga As MultiConstructor
            End Class
            Private Class MultiConstructor : Inherits PrimitiveValueObject(Of String)
                Public Sub New(array As IEnumerable(Of String))
                    MyBase.New(Join(array.ToArray, "/"))
                End Sub
                Public Sub New(ByVal value As String)
                    MyBase.New(value)
                End Sub
            End Class

            <Test()> Public Sub 複数コンストラクタがあってもPrimitiveValueObjectを適宜生成できる()
                Dim dt As DataTable = BuildTable(Of String)("Fuga", {"23/23/23"})
                Assert.That(DbAccessHelper.ConvDataTableToList(Of HogeVo)(dt, Nothing)(0).Fuga, [Is].EqualTo(New MultiConstructor("23/23/23")))
            End Sub

        End Class

        Public Class ExtractColumnNamesTest : Inherits DbAccessHelperTest

            <Test()>
            Public Sub DataTableの列名を抽出できる()
                Dim dt As DataTable = BuildTable(Of Integer)("HOGE_ID", {})
                AddColumnToTable(Of DateTime)(dt, "HOGE_DATE", {})
                AddColumnToTable(Of Decimal)(dt, "HOGE_DECIMAL", {})
                AddColumnToTable(Of String)(dt, "HOGE_NAME", {})

                Assert.That(DbAccessHelper.ExtractColumnNames(dt), [Is].EquivalentTo({"HOGE_ID", "HOGE_DATE", "HOGE_DECIMAL", "HOGE_NAME"}))
            End Sub

            <Test()>
            Public Sub DataTableの列名を抽出できる_一列()
                Dim dt As DataTable = BuildTable(Of Integer)("ID", {})

                Assert.That(DbAccessHelper.ExtractColumnNames(dt), [Is].EquivalentTo({"ID"}))
            End Sub

        End Class

    End Class
End Namespace