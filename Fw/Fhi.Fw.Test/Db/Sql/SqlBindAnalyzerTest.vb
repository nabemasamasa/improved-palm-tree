Imports System
Imports System.Collections.Generic
Imports NUnit.Framework
Imports Fhi.Fw.Db.Sql
Imports Fhi.Fw.Domain

Namespace Db.Sql

    <TestFixture()> Public MustInherit Class SqlBindAnalyzerTest

        Public Class 一般Test : Inherits SqlBindAnalyzerTest

            <Test()> Public Sub paramがnullならbindしない()
                Dim vo As HogeVo = Nothing
                Dim bind As New SqlBindAnalyzer("select * from hoge", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(0, actuals.Count)
            End Sub

            <Test()> Public Sub Idをbindする()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Id", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual("paramName = @Id, value = 123, dbType = ", actuals(0).ToString)
            End Sub

            <Test()> Public Sub 数字含むプロパティ名をbindする()
                Dim vo As New HogeVo
                vo.ExName7 = "fuga"
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@ExName7", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual("paramName = @ExName7, value = fuga, dbType = AnsiString", actuals(0).ToString)
            End Sub

            <Test()> Public Sub シングルクォートに囲まれた中のbind名にはbindしない()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where name='@id'", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(0, actuals.Count)
            End Sub

            <Test()> Public Sub シングルクォートの中のダブルクォートは無視される()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where name='""@id'", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(0, actuals.Count)
            End Sub

            <Test()> Public Sub ダブルクォートに囲まれた中のbind名にはbindしない()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where name=""@id""", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(0, actuals.Count)
            End Sub

            <Test()> Public Sub ダブルクォートの中のシングルクォートは無視される()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where name=""'@id""", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(0, actuals.Count)
            End Sub

            <Test()> Public Sub Idが二つあってもbindは一つにする()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Id and id=@Id", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual("paramName = @Id, value = 123, dbType = ", actuals(0).ToString)
            End Sub

            <Test()> Public Sub paramがprimitiveならAtValueにbindする()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", 123)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual("paramName = @Value, value = 123, dbType = ", actuals(0).ToString)
            End Sub

            <Test()> Public Sub paramがDecimalならAtValueにbindする()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", CDec("12.345"))
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual("paramName = @Value, value = 12.345, dbType = ", actuals(0).ToString)
            End Sub

            <Test()> Public Sub paramがStringならAtValueにbindする()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", "asd")
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(1, actuals.Count)
                Assert.AreEqual("paramName = @Value, value = asd, dbType = AnsiString", actuals(0).ToString)
            End Sub

            <Test()> Public Sub Vo配列のIdをBindできる()
                Dim vos As HogeVo() = {New HogeVo With {.Id = 123}, New HogeVo With {.Id = 456}}
                Dim bind As New SqlBindAnalyzer("select * from hoge where name=@Value#0$Id or name=@Value#1$Id", vos)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual("paramName = @Value#0$Id, value = 123, dbType = ", actuals(0).ToString)
                Assert.AreEqual("paramName = @Value#1$Id, value = 456, dbType = ", actuals(1).ToString)
                Assert.AreEqual(2, actuals.Count)
            End Sub

            <Test()> Public Sub bindパラメータ名があるのにparamがnullなら例外()
                Dim vo As HogeVo = Nothing
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Id", vo)
                Try
                    bind.Analyze()
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub bindパラメータ名に該当するプロパティがなかったら例外()
                Dim vo As New HogeVo
                vo.Id = 123

                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@id", vo)
                Try
                    bind.Analyze()
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub 配列Stringにbindする()

                Dim vo As New FugaVo
                vo.StrArray = New String() {"x", "y", "z"}
                Dim bind As New SqlBindAnalyzer("... @StrArray#0,@StrArray#1,@StrArray#2 ,,,", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(3, actuals.Count)
                Assert.AreEqual("paramName = @StrArray#0, value = x, dbType = AnsiString", actuals(0).ToString)
                Assert.AreEqual("paramName = @StrArray#1, value = y, dbType = AnsiString", actuals(1).ToString)
                Assert.AreEqual("paramName = @StrArray#2, value = z, dbType = AnsiString", actuals(2).ToString)
            End Sub

            <Test()> Public Sub ListStringにbindする()

                Dim vo As New FugaVo
                vo.StrList = New List(Of String)(New String() {"x", "y"})
                Dim bind As New SqlBindAnalyzer("... @StrList#0,@StrList#1 ,,,", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("paramName = @StrList#0, value = x, dbType = AnsiString", actuals(0).ToString)
                Assert.AreEqual("paramName = @StrList#1, value = y, dbType = AnsiString", actuals(1).ToString)
            End Sub

            <Test()> Public Sub 配列Voのプロパティ値にbindする()

                Dim vo As New FugaVo
                vo.HogeArray = New HogeVo() {New HogeVo, New HogeVo}
                vo.HogeArray(0).Name = "a"
                vo.HogeArray(1).Name = "b"
                Dim bind As New SqlBindAnalyzer("... @HogeArray#0$Name,@HogeArray#1$Name ,,,", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("paramName = @HogeArray#0$Name, value = a, dbType = AnsiString", actuals(0).ToString)
                Assert.AreEqual("paramName = @HogeArray#1$Name, value = b, dbType = AnsiString", actuals(1).ToString)
            End Sub

            <Test()> Public Sub 配列Voの配列Voのプロパティ値にbindする()

                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(0).HogeArray = New HogeVo() {New HogeVo, New HogeVo}
                vo.FugaArray(0).HogeArray(0).Name = "f"
                vo.FugaArray(0).HogeArray(1).Name = "g"
                Dim bind As New SqlBindAnalyzer("... @FugaArray#0$HogeArray#0$Name,@FugaArray#0$HogeArray#1$Name ,,,", vo)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("paramName = @FugaArray#0$HogeArray#0$Name, value = f, dbType = AnsiString", actuals(0).ToString)
                Assert.AreEqual("paramName = @FugaArray#0$HogeArray#1$Name, value = g, dbType = AnsiString", actuals(1).ToString)
            End Sub

            <Test()> Public Sub 配列Voが直接指定された倍意でもプロパティ値にbindする()

                Dim vos As FugaVo() = New FugaVo() {New FugaVo}
                vos(0).HogeArray = New HogeVo() {New HogeVo, New HogeVo}
                vos(0).HogeArray(0).Name = "f"
                vos(0).HogeArray(1).Name = "g"
                Dim bind As New SqlBindAnalyzer("... @Value#0$HogeArray#0$Name,@Value#0$HogeArray#1$Name ,,,", vos)
                bind.Analyze()

                Dim actuals As New List(Of String)
                bind.AddAllTo(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("paramName = @Value#0$HogeArray#0$Name, value = f, dbType = AnsiString", actuals(0).ToString)
                Assert.AreEqual("paramName = @Value#0$HogeArray#1$Name, value = g, dbType = AnsiString", actuals(1).ToString)
            End Sub

        End Class

        Private Class ParamVo
            Private _buhinNos As String()

            Public Property BuhinNos() As String()
                Get
                    Return _buhinNos
                End Get
                Set(ByVal value As String())
                    _buhinNos = value
                End Set
            End Property

            Private _buhinNo As String

            Public Property BuhinNo() As String
                Get
                    Return _buhinNo
                End Get
                Set(ByVal value As String)
                    _buhinNo = value
                End Set
            End Property
        End Class

        Public Class パラメータSQL埋めTest : Inherits SqlBindAnalyzerTest

            <Test()> Public Sub paramがnullならbindしない()
                Dim vo As HogeVo = Nothing
                Dim bind As New SqlBindAnalyzer("select * from hoge", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 数値型をbindする()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Id", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id=123", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub String型をbindする()
                Dim vo As New HogeVo With {.Name = "abc"}

                Dim bind As New SqlBindAnalyzer("select * from hoge where name=@Name", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where name='abc'", bind.MakeParametersBoundSql(), "String型はシングルクォートで囲む")
            End Sub

            <Test()> Public Sub Date型をbindする()
                Dim vo As New HogeVo With {.Last = CDate("2001/02/03 4:05:06")}

                Dim bind As New SqlBindAnalyzer("select * from hoge where last<=@Last", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where last<='2001/02/03 4:05:06'", bind.MakeParametersBoundSql(), "Date型はシングルクォートで囲む")
            End Sub

            <Test()> Public Sub 数字含むプロパティ名をbindする()
                Dim vo As New HogeVo
                vo.ExName7 = "fuga"
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@ExName7", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id='fuga'", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub シングルクォートに囲まれた中のbind名にはbindしない()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where name='@id'", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where name='@id'", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub シングルクォートの中のダブルクォートは無視される()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where name='""@id'", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where name='""@id'", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub ダブルクォートに囲まれた中のbind名にはbindしない()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where name=""@id""", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where name=""@id""", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub ダブルクォートの中のシングルクォートは無視される()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where name=""'@id""", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where name=""'@id""", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub Idが二つあってもbindは一つにする()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Id and id=@Id", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id=123 and id=123", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub paramがprimitiveならAtValueにbindする()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", 123)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id=123", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub paramがDecimalならAtValueにbindする()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", CDec("12.345"))
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id=12.345", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub paramがStringならAtValueにbindする()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", "asd")
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id='asd'", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 埋め込みパラメータ値がString型で_シングルクォートがあれば二重化()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", "fuga'")
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id='fuga'''", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 配列Stringにbindする()

                Dim vo As New FugaVo
                vo.StrArray = New String() {"x", "y", "z"}
                Dim bind As New SqlBindAnalyzer("... @StrArray#0,@StrArray#1,@StrArray#2 ,,,", vo)
                bind.Analyze()

                Assert.AreEqual("... 'x','y','z' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub ListStringにbindする()

                Dim vo As New FugaVo
                vo.StrList = New List(Of String)(New String() {"x", "y"})
                Dim bind As New SqlBindAnalyzer("... @StrList#0,@StrList#1 ,,,", vo)
                bind.Analyze()

                Assert.AreEqual("... 'x','y' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub ListStringにbindする_2桁も()

                Dim vo As New FugaVo
                vo.StrList = New List(Of String)(New String() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"})
                Dim bind As New SqlBindAnalyzer("... @StrList#0,@StrList#1,@StrList#2,@StrList#3,@StrList#4,@StrList#5,@StrList#6,@StrList#7,@StrList#8,@StrList#9,@StrList#10 ,,,", vo)
                bind.Analyze()

                Assert.AreEqual("... '0','1','2','3','4','5','6','7','8','9','10' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 配列Voのプロパティ値にbindする()

                Dim vo As New FugaVo
                vo.HogeArray = New HogeVo() {New HogeVo, New HogeVo}
                vo.HogeArray(0).Name = "a"
                vo.HogeArray(1).Name = "b"
                Dim bind As New SqlBindAnalyzer("... @HogeArray#0$Name,@HogeArray#1$Name ,,,", vo)
                bind.Analyze()

                Assert.AreEqual("... 'a','b' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 配列Voの配列Voのプロパティ値にbindする()

                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(0).HogeArray = New HogeVo() {New HogeVo, New HogeVo}
                vo.FugaArray(0).HogeArray(0).Name = "f"
                vo.FugaArray(0).HogeArray(1).Name = "g"
                Dim bind As New SqlBindAnalyzer("... @FugaArray#0$HogeArray#0$Name,@FugaArray#0$HogeArray#1$Name ,,,", vo)
                bind.Analyze()

                Assert.AreEqual("... 'f','g' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 配列Voが直接指定された倍意でもプロパティ値にbindする()

                Dim vos As FugaVo() = New FugaVo() {New FugaVo}
                vos(0).HogeArray = New HogeVo() {New HogeVo, New HogeVo}
                vos(0).HogeArray(0).Name = "f"
                vos(0).HogeArray(1).Name = "g"
                Dim bind As New SqlBindAnalyzer("... @Value#0$HogeArray#0$Name,@Value#0$HogeArray#1$Name ,,,", vos)
                bind.Analyze()

                Assert.AreEqual("... 'f','g' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub コレクション型パラメータが12番目以降に適切に埋め込めない問題を修正()
                Dim vo As New ParamVo With {.BuhinNos = New String() {"A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12"}}
                Dim bind As New SqlBindAnalyzer("select * from hoge where buhin_no in (@BuhinNos#0,@BuhinNos#1,@BuhinNos#2,@BuhinNos#3,@BuhinNos#4,@BuhinNos#5,@BuhinNos#6,@BuhinNos#7,@BuhinNos#8,@BuhinNos#9,@BuhinNos#10,@BuhinNos#11)", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where buhin_no in ('A1','A2','A3','A4','A5','A6','A7','A8','A9','A10','A11','A12')", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 単埋め込みパラメータと複埋め込みパラメータがあると_複に単が埋め込まれる問題を修正()
                Dim vo As New ParamVo With {.BuhinNo = "14%", .BuhinNos = New String() {"ABC", "XYZ"}}
                Dim bind As New SqlBindAnalyzer("select * from hoge where buhin_no in (@BuhinNos#0,@BuhinNos#1) and buhin_no=@BuhinNo", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where buhin_no in ('ABC','XYZ') and buhin_no='14%'", bind.MakeParametersBoundSql())
            End Sub

        End Class

        Public Class パラメータSQL埋めTest_Criteria : Inherits SqlBindAnalyzerTest

            Private vo As HogeVo
            Private criteria As CriteriaBinder

            <SetUp()> Public Sub SetUp()
                vo = New HogeVo()
                criteria = New Criteria(Of HogeVo)(vo)
            End Sub

            <Test()> Public Sub Criteriaを利用してSQLを発行できる()
                criteria.Equal(vo.Name, "名乗れよ").Not.Equal(vo.Id, 123)

                Dim bind As New SqlBindAnalyzer("SELECT * FROM HOGE WHERE NAME = @Name0 AND NOT ID = @Id1", criteria)
                bind.Analyze()

                Assert.That(bind.MakeParametersBoundSql(), [Is].EqualTo("SELECT * FROM HOGE WHERE NAME = '名乗れよ' AND NOT ID = 123"))
            End Sub

            <Test()> Public Sub Criteriaを利用してSQLを発行できる_配列()
                criteria.Any(vo.Name, {"クミン", "ターメリック", "レッドチリ"})

                Dim bind As New SqlBindAnalyzer("SELECT * FROM HOGE WHERE NAME IN (@Name0#0,@Name0#1,@Name0#2)", criteria)
                bind.Analyze()

                Assert.That(bind.MakeParametersBoundSql(), [Is].EqualTo("SELECT * FROM HOGE WHERE NAME IN ('クミン','ターメリック','レッドチリ')"))
            End Sub

        End Class

#Region "Nested ValueObject..."
        Private Class Id : Inherits PrimitiveValueObject(Of Integer)
            Public Sub New(ByVal value As Integer)
                MyBase.New(value)
            End Sub
        End Class
        Private Class Money : Inherits PrimitiveValueObject(Of Decimal)
            Public Sub New(ByVal value As Decimal)
                MyBase.New(value)
            End Sub
        End Class
        Private Class Name : Inherits PrimitiveValueObject(Of String)
            Public Sub New(ByVal value As String)
                MyBase.New(value)
            End Sub
        End Class
        Private Class Date2 : Inherits PrimitiveValueObject(Of DateTime)
            Public Sub New(ByVal value As String)
                Me.New(CDate(value))
            End Sub
            Public Sub New(ByVal value As DateTime)
                MyBase.New(value)
            End Sub
        End Class
        Private Class Entity
            Public Property Id As Id
            Public Property Name As Name
            Public Property Last As Date2
            Public Property NameArray As Name()
            Public Property NameList As List(Of Name)
        End Class
        Private Class MyValueObject : Inherits ValueObject
            Private _Id As Id
            Private _Name As Name
            Private _Last As Date2
            Public ReadOnly Property Id As Id
                Get
                    Return _Id
                End Get
            End Property
            Public ReadOnly Property Name As Name
                Get
                    Return _Name
                End Get
            End Property
            Public ReadOnly Property Last As Date2
                Get
                    Return _Last
                End Get
            End Property
            Public Sub New(ByVal id As Id, ByVal name As Name, ByVal last As Date2)
                _Id = id
                _Name = name
                _Last = last
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Throw New NotImplementedException()
            End Function
        End Class

#End Region
        Public Class PrimitiveValueObjectパラメータSQL埋めTest : Inherits SqlBindAnalyzerTest

            <Test()> Public Sub paramがnullならbindしない()
                Dim anId As Id = Nothing
                Dim bind As New SqlBindAnalyzer("select * from hoge", anId)
                bind.Analyze()

                Assert.AreEqual("select * from hoge", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub AtValueにbindする_Int()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", New Id(123))
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id=123", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub AtValueにbindする_Decimal()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", New Money(12.345D))
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id=12.345", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub AtValueにbindする_String()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", New Name("asd"))
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id='asd'", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub AtValueにbindする_Date()
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Value", New Date2("2012/03/04 05:06:07"))
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id='2012/03/04 5:06:07'", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 配列Nameにbindする()

                Dim names As Name() = {New Name("x"), New Name("y"), New Name("z")}
                Dim bind As New SqlBindAnalyzer("... @Value#0,@Value#1,@Value#2 ,,,", names)
                bind.Analyze()

                Assert.AreEqual("... 'x','y','z' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub ListNameにbindする()

                Dim names As New List(Of Name)({New Name("x"), New Name("y")})
                Dim bind As New SqlBindAnalyzer("... @Value#0,@Value#1 ,,,", names)
                bind.Analyze()

                Assert.AreEqual("... 'x','y' ,,,", bind.MakeParametersBoundSql())
            End Sub

        End Class

        Public Class EntityパラメータSQL埋めTest : Inherits SqlBindAnalyzerTest

            <Test()> Public Sub paramがnullならbindしない()
                Dim vo As Entity = Nothing
                Dim bind As New SqlBindAnalyzer("select * from hoge", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 数値型をbindする()
                Dim vo As New Entity With {.Id = New Id(123)}
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Id", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id=123", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub String型をbindする()
                Dim vo As New Entity With {.Name = New Name("abc")}

                Dim bind As New SqlBindAnalyzer("select * from hoge where name=@Name", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where name='abc'", bind.MakeParametersBoundSql(), "String型はシングルクォートで囲む")
            End Sub

            <Test()> Public Sub Date型をbindする()
                Dim vo As New Entity With {.Last = New Date2("2001/02/03 4:05:06")}

                Dim bind As New SqlBindAnalyzer("select * from hoge where last<=@Last", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where last<='2001/02/03 4:05:06'", bind.MakeParametersBoundSql(), "Date型はシングルクォートで囲む")
            End Sub

            <Test()> Public Sub 配列Stringにbindする()

                Dim vo As New Entity
                vo.NameArray = {New Name("x"), New Name("y"), New Name("z")}
                Dim bind As New SqlBindAnalyzer("... @NameArray#0,@NameArray#1,@NameArray#2 ,,,", vo)
                bind.Analyze()

                Assert.AreEqual("... 'x','y','z' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub ListStringにbindする()

                Dim vo As New Entity
                vo.NameList = New List(Of Name)({New Name("x"), New Name("y")})
                Dim bind As New SqlBindAnalyzer("... @NameList#0,@NameList#1 ,,,", vo)
                bind.Analyze()

                Assert.AreEqual("... 'x','y' ,,,", bind.MakeParametersBoundSql())
            End Sub

        End Class

        Public Class ValueObjectパラメータSQL埋めTest : Inherits SqlBindAnalyzerTest

            <Test()> Public Sub paramがnullならbindしない()
                Dim vo As MyValueObject = Nothing
                Dim bind As New SqlBindAnalyzer("select * from hoge", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub 数値型をbindする()
                Dim vo As New MyValueObject(New Id(123), Nothing, Nothing)
                Dim bind As New SqlBindAnalyzer("select * from hoge where id=@Id", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where id=123", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub String型をbindする()
                Dim vo As New MyValueObject(Nothing, New Name("abc"), Nothing)

                Dim bind As New SqlBindAnalyzer("select * from hoge where name=@Name", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where name='abc'", bind.MakeParametersBoundSql(), "String型はシングルクォートで囲む")
            End Sub

            <Test()> Public Sub Date型をbindする()
                Dim vo As New MyValueObject(Nothing, Nothing, New Date2("2001/02/03 4:05:06"))

                Dim bind As New SqlBindAnalyzer("select * from hoge where last<=@Last", vo)
                bind.Analyze()

                Assert.AreEqual("select * from hoge where last<='2001/02/03 4:05:06'", bind.MakeParametersBoundSql(), "Date型はシングルクォートで囲む")
            End Sub

            <Test()> Public Sub 配列ValueObjectにbindする()

                Dim array As MyValueObject() = {New MyValueObject(Nothing, New Name("x"), Nothing),
                                                New MyValueObject(Nothing, New Name("y"), Nothing),
                                                New MyValueObject(Nothing, New Name("z"), Nothing)}
                Dim bind As New SqlBindAnalyzer("... @Value#0$Name,@Value#1$Name,@Value#2$Name ,,,", array)
                bind.Analyze()

                Assert.AreEqual("... 'x','y','z' ,,,", bind.MakeParametersBoundSql())
            End Sub

            <Test()> Public Sub ListValueObjectにbindする()

                Dim vo As New List(Of MyValueObject)({New MyValueObject(Nothing, New Name("x"), Nothing),
                                                      New MyValueObject(Nothing, New Name("y"), Nothing)})
                Dim bind As New SqlBindAnalyzer("... @Value#0$Name,@Value#1$Name ,,,", vo)
                bind.Analyze()

                Assert.AreEqual("... 'x','y' ,,,", bind.MakeParametersBoundSql())
            End Sub

        End Class

    End Class
End Namespace
