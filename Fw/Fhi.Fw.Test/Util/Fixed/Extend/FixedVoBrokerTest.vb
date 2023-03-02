Imports Fhi.Fw.Domain
Imports NUnit.Framework

Namespace Util.Fixed.Extend
    Public MustInherit Class FixedVoBrokerTest

#Region "Testing Fixed Classes"
        Protected Class MyBasic
            Public Property Seq() As String
            Public Property Children() As MyRoot()
        End Class

        Protected Class MyRoot
            Public Property Num() As String
            Public Property Customer() As MyCustomer
            Public Property CustomerArray() As MyCustomer()
            Public Property Insu() As MyInsu()
        End Class

        Protected Class MyCustomer
            Public Property Tanto() As String
            Public Property Weight() As Decimal?
            Public Property Id() As Integer
        End Class

        Protected Class MyInsu
            Public Property Insu() As String
        End Class

        Protected Class TypeVo
            Public Property [Int]() As Integer
            Public Property IntArray() As Integer()
            Public Property IntList() As List(Of Integer)
            Public Property NullableInt() As Integer?
            Public Property NullableIntArray() As Integer?()
            Public Property NullableIntList() As List(Of Integer?)
            Public Property Dec() As Decimal
            Public Property DecArray() As Decimal()
            Public Property DecList() As List(Of Decimal)
            Public Property NullableDec() As Decimal?
            Public Property NullableDecArray() As Decimal?()
            Public Property NullableDecList() As List(Of Decimal?)
            Public Property [Str]() As String
            Public Property StrArray() As String()
            Public Property StrList() As List(Of String)
            Public Property Header() As String
            Public Property Footer() As String
            Public Property StrPvo() As StringPvo
            Public Property IntPvo() As IntegerPvo
        End Class

        Protected Class ExTypeVo : Inherits TypeVo
            Public Property TypeVoChild() As TypeVo
            Public Property TypeVoArray() As TypeVo()
            Public Property TypeVoList() As List(Of TypeVo)
        End Class

        Private Class MyFixedStringVo
            Public Property First() As String
            Public Property Second() As String
            Public Property Third() As String
        End Class

        Protected Class StringPvo : Inherits PrimitiveValueObject(Of String)
            Public Sub New(value As String)
                MyBase.New(value)
            End Sub
        End Class

        Protected Class IntegerPvo : Inherits PrimitiveValueObject(Of Integer)
            Public Sub New(value As Integer)
                MyBase.New(value)
            End Sub
        End Class
#End Region

#Region "implements IFixedRule"
        Protected Class CustomerRuleImpl : Implements IFixedRule(Of MyCustomer)

            Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As MyCustomer) Implements IFixedRule(Of MyCustomer).Configure
                defineBy.Zenkaku(vo.Tanto, 12).Number(vo.Weight, 10, 2)
            End Sub
        End Class

        Protected Class InsuRuleImpl : Implements IFixedRule(Of MyInsu)

            Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As MyInsu) Implements IFixedRule(Of MyInsu).Configure
                defineBy.Hankaku(vo.Insu, 2)
            End Sub
        End Class

        Protected Class RootRuleImpl : Implements IFixedRule(Of MyRoot)

            Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As MyRoot) Implements IFixedRule(Of MyRoot).Configure
                defineBy.Hankaku(vo.Num, 4).GroupRepeat(vo.CustomerArray, New CustomerRuleImpl, 3).GroupRepeat(vo.Insu, New InsuRuleImpl, 5)
            End Sub
        End Class

        Protected Class BasicRuleImpl : Implements IFixedRule(Of MyBasic)

            Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As MyBasic) Implements IFixedRule(Of MyBasic).Configure
                defineBy.Hankaku(vo.Seq, 2).GroupRepeat(vo.Children, New RootRuleImpl, 2)
            End Sub
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()
            ' nop
        End Sub

        Public Class LearningTest : Inherits FixedVoBrokerTest
            <Test()> Public Sub ListOfNullableIntならNothingをAddできる()
                Dim cut As New List(Of Integer?)
                cut.Add(Nothing)
                Assert.IsTrue(True)
            End Sub

            Public Sub dotNET35では_IList型だとNothingをAddできない()
                Dim cut As IList = New List(Of Integer?)
                Try
                    cut.Add(Nothing)
                    Assert.Fail()

                Catch expected As ArgumentException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub dotNET4では_IList型でもNothingをAddできるようになる()
                Dim cut As IList = New List(Of Integer?)
                cut.Add(Nothing)
                Assert.IsTrue(True)
            End Sub
        End Class

        Public Class コンストラクタTest : Inherits FixedVoBrokerTest

            <Test()> Public Sub Hankaku_半角文字で初期化()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Hankaku(vo.Str, 3))

                Assert.AreEqual("   ", broker.FixedString, "半角3文字")
                Assert.AreEqual("", broker.Vo.Str)
            End Sub

            <Test()> Public Sub Hankaku_decoratorを指定すると装飾されて初期化される()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Hankaku(vo.Str, 3, decorateToString:=Function(value) "bbb"))

                Assert.AreEqual("bbb", broker.FixedString, "半角3文字")
                Assert.AreEqual("", broker.Vo.Str)
            End Sub

            <Test()> Public Sub Hankaku_PrimitiveValueObjectを半角文字で初期化()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Hankaku(vo.StrPvo, 3))

                Assert.That(broker.FixedString, [Is].EqualTo("   "), "半角3文字")
                Assert.IsNull(broker.Vo.StrPvo)
            End Sub

            <Test()> Public Sub Zenkaku_全角文字で初期化()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Zenkaku(vo.Str, 3))

                Assert.AreEqual("　　　", broker.FixedString, "全角3文字")
                Assert.AreEqual("", broker.Vo.Str)
            End Sub

            <Test()> Public Sub Zenkaku_PrimitiveValueObjectを半角文字で初期化()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Zenkaku(vo.StrPvo, 3))

                Assert.That(broker.FixedString, [Is].EqualTo("　　　"), "全角3文字")
                Assert.IsNull(broker.Vo.StrPvo)
            End Sub

            <Test()> Public Sub Number_プリミティブ型なら0埋めで初期化_Int型()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Number(vo.Int, 3, 0))

                Assert.AreEqual("000", broker.FixedString, "プリミティブ型だから0埋め")
                Assert.AreEqual(0, broker.Vo.Int)
            End Sub

            <Test()> Public Sub Number_プリミティブ型なら0埋めで初期化_Decimal型()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Number(vo.Dec, 4, 2))

                Assert.AreEqual("0000", broker.FixedString, "プリミティブ型だから0埋め")
                Assert.AreEqual(0D, broker.Vo.Dec)
            End Sub

            <Test()> Public Sub Number_Nullable型なら空白で初期化()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Number(vo.NullableInt, 3, 0))

                Assert.AreEqual("   ", broker.FixedString, "空白で初期化")
                Assert.IsNull(broker.Vo.NullableInt)
            End Sub

            <Test()> Public Sub Number_PrimitiveValueObjectなら空白で初期化()
                Dim broker As New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.Number(vo.IntPvo, 3, 0))

                Assert.That(broker.FixedString, [Is].EqualTo("   "), "空白で初期化")
                Assert.IsNull(broker.Vo.IntPvo)
            End Sub

            <Test()> Public Sub Constructor_インスタンス生成直後に固定長文字列は初期化されている_Group項目が配列じゃなくても動作する()
                Dim subRule As New FixedRule(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                            defineBy.Zenkaku(vo.Str, 12).Number(vo.NullableDec, 10, 2))
                Dim voString As New FixedVoBroker(Of ExTypeVo)(Function(defineBy As IFixedRuleLocator, vo As ExTypeVo) _
                                                                    defineBy.Hankaku(vo.Header, 3).Group(vo.TypeVoChild, subRule))

                Assert.AreEqual("   " _
                    & "　　　　　　　　　　　　" & "          " _
                    , voString.FixedString)
            End Sub

            <Test()> Public Sub Constructor_インスタンス生成直後に固定長文字列は初期化されている_2階層()
                Dim customerRule As New FixedRule(Of MyCustomer)(Function(defineBy As IFixedRuleLocator, vo As MyCustomer) _
                                                                              defineBy.Zenkaku(vo.Tanto, 12).Number(vo.Weight, 10, 2))
                Dim insuRule As New FixedRule(Of MyInsu)(Function(defineBy As IFixedRuleLocator, vo As MyInsu) _
                                                             defineBy.Hankaku(vo.Insu, 2))
                Dim voString As New FixedVoBroker(Of MyRoot)(Function(defineBy As IFixedRuleLocator, vo As MyRoot) _
                                                                 defineBy.Hankaku(vo.Num, 4).GroupRepeat(vo.CustomerArray, customerRule, 3).GroupRepeat(vo.Insu, insuRule, 5))
                Assert.AreEqual("    " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "  " & "  " & "  " & "  " & "  " _
                    , voString.FixedString)
            End Sub

            <Test()> Public Sub Constructor_インスタンス生成直後に固定長文字列は初期化されている_3階層()
                Dim voString As New FixedVoBroker(Of MyBasic)(New BasicRuleImpl)

                Assert.AreEqual("  " _
                    & "    " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "  " & "  " & "  " & "  " & "  " _
                    & "    " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "  " & "  " & "  " & "  " & "  " _
                    , voString.FixedString)
            End Sub

            <Test()> Public Sub Constructor_インスタンス生成直後にVoは初期化されている_2階層()
                Dim hoge As New FixedVoBroker(Of MyRoot)(New RootRuleImpl)

                Assert.AreEqual("", hoge.Vo.Num)
                Assert.AreEqual(3, hoge.Vo.CustomerArray.Length)
                With hoge.Vo.CustomerArray(0)
                    Assert.AreEqual("", .Tanto)
                    Assert.IsNull(.Weight)
                End With
                With hoge.Vo.CustomerArray(1)
                    Assert.AreEqual("", .Tanto)
                    Assert.IsNull(.Weight)
                End With
                With hoge.Vo.CustomerArray(2)
                    Assert.AreEqual("", .Tanto)
                    Assert.IsNull(.Weight)
                End With
                Assert.AreEqual(5, hoge.Vo.Insu.Length)
                Assert.AreEqual("", hoge.Vo.Insu(0).Insu)
                Assert.AreEqual("", hoge.Vo.Insu(1).Insu)
                Assert.AreEqual("", hoge.Vo.Insu(2).Insu)
                Assert.AreEqual("", hoge.Vo.Insu(3).Insu)
                Assert.AreEqual("", hoge.Vo.Insu(4).Insu)
            End Sub

            <Test()> Public Sub Constructor_インスタンス生成直後にVoは初期化されている_List()
                Dim subRule As New FixedRule(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                            defineBy.Zenkaku(vo.Str, 12).Number(vo.NullableDec, 10, 2))
                Dim broker As New FixedVoBroker(Of ExTypeVo)(Function(defineBy As IFixedRuleLocator, vo As ExTypeVo) _
                                                                 defineBy.Hankaku(vo.Header, 2).GroupRepeat(vo.TypeVoList, subRule, 3))

                Assert.AreEqual("", broker.Vo.Header)
                Assert.AreEqual(3, broker.Vo.TypeVoList.Count)
                With broker.Vo.TypeVoList(0)
                    Assert.AreEqual("", .Str)
                    Assert.IsNull(.NullableDec)
                End With
                With broker.Vo.TypeVoList(1)
                    Assert.AreEqual("", .Str)
                    Assert.IsNull(.NullableDec)
                End With
                With broker.Vo.TypeVoList(2)
                    Assert.AreEqual("", .Str)
                    Assert.IsNull(.NullableDec)
                End With
            End Sub
        End Class

        Public Class 全体Test : Inherits FixedVoBrokerTest

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する()
                Dim voString As New FixedVoBroker(Of MyRoot)(New RootRuleImpl)
                With voString.Vo
                    .Num = "abc"
                    .CustomerArray(1).Tanto = "tanto0"
                    .CustomerArray(2).Weight = 123.4@
                    .Insu(1).Insu = "a"
                    .Insu(2).Insu = " b"
                End With

                voString.ApplyVoToFixedString()

                Assert.AreEqual("abc " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "ｔａｎｔｏ０　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "0000012340" _
                    & "  " & "a " & " b" & "  " & "  " _
                    , voString.FixedString)
            End Sub

            <Test()> Public Sub ApplyFixedStringToVo_固定長文字列の内容をVoに反映する()
                Dim voString As New FixedVoBroker(Of MyRoot)(New RootRuleImpl)

                voString.FixedString = " xy " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "ｔａｎｔｏ９　　　　　　" & "0000000000" _
                    & "　　　　　　　　　　　　" & "0000012340" _
                    & "  " & "a " & " b" & "cd" & "  "

                voString.ApplyFixedStringToVo()

                With voString.Vo
                    Assert.AreEqual(" xy", .Num)
                    Assert.AreEqual(3, .CustomerArray.Length)
                    With .CustomerArray(0)
                        Assert.AreEqual("", .Tanto)
                        Assert.IsNull(.Weight)
                    End With
                    With .CustomerArray(1)
                        Assert.AreEqual("ｔａｎｔｏ９", .Tanto)
                        Assert.AreEqual(0@, .Weight)
                    End With
                    With .CustomerArray(2)
                        Assert.AreEqual("", .Tanto)
                        Assert.AreEqual(123.4@, .Weight)
                    End With
                    Assert.AreEqual(5, .Insu.Length)
                    Assert.AreEqual("", .Insu(0).Insu)
                    Assert.AreEqual("a", .Insu(1).Insu)
                    Assert.AreEqual(" b", .Insu(2).Insu)
                    Assert.AreEqual("cd", .Insu(3).Insu)
                    Assert.AreEqual("", .Insu(4).Insu)
                End With

            End Sub

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する_3階層()
                Dim voString As New FixedVoBroker(Of MyBasic)(New BasicRuleImpl)

                voString.Vo.Children(0).CustomerArray(2).Tanto = "abc"
                voString.Vo.Children(1).CustomerArray(1).Weight = 1234.5678@

                voString.ApplyVoToFixedString()

                Assert.AreEqual("  " _
                    & "    " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "ａｂｃ　　　　　　　　　" & "          " _
                    & "  " & "  " & "  " & "  " & "  " _
                    & "    " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "0000123456" _
                    & "　　　　　　　　　　　　" & "          " _
                    & "  " & "  " & "  " & "  " & "  " _
                     , voString.FixedString)
            End Sub

            <Test()> Public Sub ApplyFixedStringToVo_Voで設定した内容を_固定長文字列で上書きする()
                Dim voString As New FixedVoBroker(Of MyRoot)(New RootRuleImpl)

                voString.Vo.CustomerArray(0).Tanto = "aa"

                voString.ApplyVoToFixedString()

                voString.FixedString = "    " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "ｂｂ　　　　　　　　　　" & "          " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "  " & "  " & "  " & "  " & "  "

                voString.ApplyFixedStringToVo()

                Assert.AreEqual("", voString.Vo.CustomerArray(0).Tanto)
                Assert.AreEqual("ｂｂ", voString.Vo.CustomerArray(1).Tanto)
            End Sub

            <Test()> Public Sub NewVoInstanceAndInitialize_Voのインスタンスを再生成する()
                Dim voString As New FixedVoBroker(Of MyRoot)(New RootRuleImpl)

                voString.FixedString = " xy " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "ｔａｎｔｏ９　　　　　　" & "0000000000" _
                    & "　　　　　　　　　　　　" & "0000012340" _
                    & "  " & "a " & " b" & "cd" & "  "

                voString.ApplyFixedStringToVo()
                Dim actual1 As MyRoot = voString.Vo

                voString.NewVoInstanceAndInitialize()

                voString.FixedString = " vw " _
                    & "　　　　　　　　　　　　" & "          " _
                    & "ｔａｎｔｏ９　　　　　　" & "0000000000" _
                    & "　　　　　　　　　　　　" & "0000012340" _
                    & "  " & "a " & " b" & "cd" & "  "

                voString.ApplyFixedStringToVo()
                Dim actual2 As MyRoot = voString.Vo

                Assert.AreNotSame(actual1, actual2)
                Assert.AreNotSame(actual1.CustomerArray, actual2.CustomerArray)
                Assert.AreNotSame(actual1.Insu, actual2.Insu)

                Assert.AreEqual(" xy", actual1.Num)
                Assert.AreEqual(" vw", actual2.Num)
            End Sub
        End Class

        Public Class 数値型変換Test : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of TypeVo)

            Public Overrides Sub SetUp()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                          defineBy.Number(vo.Int, 3, 0).Number(vo.NullableInt, 3, 0) _
                                                          .Number(vo.Dec, 4, 2).Number(vo.NullableDec, 4, 2))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_は数値項目でも空文字で初期化()
                Assert.AreEqual("000" & "   " & "0000" & "    ", broker.FixedString)
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_数値項目はゼロ埋めで初期化()
                broker.IsZeroPaddingIfNull = True
                Assert.AreEqual("000" & "000" & "0000" & "0000", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(0, .Int, "int型は0で初期化")
                    Assert.IsNull(.NullableInt, "null可はnullで初期化")
                    Assert.AreEqual(0, .Dec, "Decimal型は0で初期化")
                    Assert.IsNull(.NullableDec, "null可はnullで初期化")
                End With
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_Voの数値項目は0で初期化()
                broker.IsZeroPaddingIfNull = True
                With broker.Vo
                    Assert.AreEqual(0, .Int)
                    Assert.AreEqual(0, .NullableInt, "null可でも、0で初期化")
                    Assert.AreEqual(0, .Dec)
                    Assert.AreEqual(0, .NullableDec, "null可でも、0で初期化")
                End With
            End Sub

            <Test()> Public Sub FixedString_Voの内容を変更して_ApplyVoToFixedStringで_FixedStringに反映する()
                With broker.Vo
                    .Int = 1
                    .NullableInt = 12
                    .Dec = 23.45D
                    .NullableDec = 0.01D
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("001" & "012" & "2345" & "0001", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_FixedStringの内容を変更して_ApplyFixedStringToVoで_Voに反映する()
                broker.FixedString = "001" & "012" & "2345" & "0001"
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(1, .Int)
                    Assert.AreEqual(12, .NullableInt)
                    Assert.AreEqual(23.45D, .Dec)
                    Assert.AreEqual(0.01D, .NullableDec)
                End With
            End Sub

            <Test()> Public Sub Int型3桁に整数5桁を設定したら_1000の位より上の桁が切れる()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                          defineBy.Number(vo.Int, 3, 0).Number(vo.NullableInt, 3, 0))
                With broker.Vo
                    .Int = 12345
                    .NullableInt = 67890
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("345" & "890", broker.FixedString, "1000の位より上の桁が切れる")
            End Sub

            <Test()> Public Sub Dec型4桁内小数2桁に整数4桁と小数以下4桁を設定したら_100の位より上の桁と小数第三位以下の桁が切れる()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                          defineBy.Number(vo.Dec, 4, 2).Number(vo.NullableDec, 4, 2))
                With broker.Vo
                    .Dec = 0.1234D
                    .NullableDec = 5678D
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("0012" & "7800", broker.FixedString, "100の位より上の桁と小数第三位以下の桁が切れる")
            End Sub

        End Class

        Public Class Int型配列で繰り返しTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of TypeVo)

            Public Overrides Sub SetUp()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                          defineBy.NumberRepeat(vo.IntArray, 3, 0, 2).NumberRepeat(vo.NullableIntArray, 3, 0, 2))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("000" & "000" & "   " & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_数値項目はゼロ埋めで初期化()
                broker.IsZeroPaddingIfNull = True
                Assert.AreEqual("000" & "000" & "000" & "000", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(2, .IntArray.Length)
                    Assert.AreEqual(0, .IntArray(0), "int型は0で初期化")
                    Assert.AreEqual(0, .IntArray(1))
                    Assert.AreEqual(2, .NullableIntArray.Length)
                    Assert.IsNull(.NullableIntArray(0), "null可はnullで初期化")
                    Assert.IsNull(.NullableIntArray(1))
                End With
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_Voの数値項目は0で初期化()
                broker.IsZeroPaddingIfNull = True
                With broker.Vo
                    Assert.AreEqual(2, .IntArray.Length)
                    Assert.AreEqual(0, .IntArray(0))
                    Assert.AreEqual(0, .IntArray(1))
                    Assert.AreEqual(2, .NullableIntArray.Length)
                    Assert.AreEqual(0, .NullableIntArray(0), "null可でも、0で初期化")
                    Assert.AreEqual(0, .NullableIntArray(1))
                End With
            End Sub

            <Test()> Public Sub FixedString_Voの内容を変更して_ApplyVoToFixedStringで_FixedStringに反映する()
                With broker.Vo
                    .IntArray(0) = 1
                    .IntArray(1) = 12
                    .NullableIntArray(0) = 987
                    .NullableIntArray(1) = 2
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("001" & "012" & "987" & "002", broker.FixedString)
            End Sub

            <Test()> Public Sub FixedString_Voの内容を変更して_ApplyVoToFixedStringで_FixedStringに反映する_NullableはNull値だと空白()
                With broker.Vo
                    .IntArray(0) = 1
                    .IntArray(1) = 12
                    .NullableIntArray(0) = Nothing
                    .NullableIntArray(1) = 2
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("001" & "012" & "   " & "002", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_FixedStringの内容を変更して_ApplyFixedStringToVoで_Voに反映する()
                broker.FixedString = "001" & "012" & "987" & "002"
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(1, .IntArray(0))
                    Assert.AreEqual(12, .IntArray(1))
                    Assert.AreEqual(987, .NullableIntArray(0))
                    Assert.AreEqual(2, .NullableIntArray(1))
                End With
            End Sub

        End Class

        Public Class ListOfInt型で繰り返しTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of TypeVo)

            Public Overrides Sub SetUp()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                          defineBy.NumberRepeat(vo.IntList, 3, 0, 2).NumberRepeat(vo.NullableIntList, 3, 0, 2))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("000" & "000" & "   " & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_数値項目はゼロ埋めで初期化()
                broker.IsZeroPaddingIfNull = True
                Assert.AreEqual("000" & "000" & "000" & "000", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(2, .IntList.Count)
                    Assert.AreEqual(0, .IntList(0), "int型は0で初期化")
                    Assert.AreEqual(0, .IntList(1))
                    Assert.AreEqual(2, .NullableIntList.Count)
                    Assert.IsNull(.NullableIntList(0), "null可はnullで初期化")
                    Assert.IsNull(.NullableIntList(1))
                End With
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_Voの数値項目は0で初期化()
                broker.IsZeroPaddingIfNull = True
                With broker.Vo
                    Assert.AreEqual(2, .IntList.Count)
                    Assert.AreEqual(0, .IntList(0))
                    Assert.AreEqual(0, .IntList(1))
                    Assert.AreEqual(2, .NullableIntList.Count)
                    Assert.AreEqual(0, .NullableIntList(0), "null可でも、0で初期化")
                    Assert.AreEqual(0, .NullableIntList(1))
                End With
            End Sub

            <Test()> Public Sub FixedString_Voの内容を変更して_ApplyVoToFixedStringで_FixedStringに反映する()
                With broker.Vo
                    .IntList(0) = 1
                    .IntList(1) = 12
                    .NullableIntList(0) = 987
                    .NullableIntList(1) = 2
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("001" & "012" & "987" & "002", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_FixedStringの内容を変更して_ApplyFixedStringToVoで_Voに反映する()
                broker.FixedString = "001" & "012" & "987" & "002"
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(1, .IntList(0))
                    Assert.AreEqual(12, .IntList(1))
                    Assert.AreEqual(987, .NullableIntList(0))
                    Assert.AreEqual(2, .NullableIntList(1))
                End With
            End Sub

        End Class

        Public Class Dec型配列で繰り返しTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of TypeVo)

            Public Overrides Sub SetUp()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                          defineBy.NumberRepeat(vo.DecArray, 3, 2, 2).NumberRepeat(vo.NullableDecArray, 3, 1, 2))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("000" & "000" & "   " & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_数値項目はゼロ埋めで初期化()
                broker.IsZeroPaddingIfNull = True
                Assert.AreEqual("000" & "000" & "000" & "000", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(2, .DecArray.Length)
                    Assert.AreEqual(0, .DecArray(0), "int型は0で初期化")
                    Assert.AreEqual(0, .DecArray(1))
                    Assert.AreEqual(2, .NullableDecArray.Length)
                    Assert.IsNull(.NullableDecArray(0), "null可はnullで初期化")
                    Assert.IsNull(.NullableDecArray(1))
                End With
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_Voの数値項目は0で初期化()
                broker.IsZeroPaddingIfNull = True
                With broker.Vo
                    Assert.AreEqual(2, .DecArray.Length)
                    Assert.AreEqual(0, .DecArray(0))
                    Assert.AreEqual(0, .DecArray(1))
                    Assert.AreEqual(2, .NullableDecArray.Length)
                    Assert.AreEqual(0, .NullableDecArray(0), "null可でも、0で初期化")
                    Assert.AreEqual(0, .NullableDecArray(1))
                End With
            End Sub

            <Test()> Public Sub FixedString_Voの内容を変更して_ApplyVoToFixedStringで_FixedStringに反映する()
                With broker.Vo
                    .DecArray(0) = 1.23D
                    .DecArray(1) = 45.6D
                    .NullableDecArray(0) = 12.3D
                    .NullableDecArray(1) = 4.56D
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("123" & "560" & "123" & "045", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_FixedStringの内容を変更して_ApplyFixedStringToVoで_Voに反映する()
                broker.FixedString = "123" & "560" & "123" & "045"
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(1.23D, .DecArray(0))
                    Assert.AreEqual(5.6D, .DecArray(1))
                    Assert.AreEqual(12.3, .NullableDecArray(0))
                    Assert.AreEqual(4.5, .NullableDecArray(1))
                End With
            End Sub

        End Class

        Public Class ListOfDec型で繰り返しTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of TypeVo)

            Public Overrides Sub SetUp()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                          defineBy.NumberRepeat(vo.DecList, 3, 2, 2).NumberRepeat(vo.NullableDecList, 3, 1, 2))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("000" & "000" & "   " & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_数値項目はゼロ埋めで初期化()
                broker.IsZeroPaddingIfNull = True
                Assert.AreEqual("000" & "000" & "000" & "000", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(2, .DecList.Count)
                    Assert.AreEqual(0, .DecList(0), "int型は0で初期化")
                    Assert.AreEqual(0, .DecList(1))
                    Assert.AreEqual(2, .NullableDecList.Count)
                    Assert.IsNull(.NullableDecList(0), "null可はnullで初期化")
                    Assert.IsNull(.NullableDecList(1))
                End With
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_IsZeroPaddingIfNullがTrueなら_Voの数値項目は0で初期化()
                broker.IsZeroPaddingIfNull = True
                With broker.Vo
                    Assert.AreEqual(2, .DecList.Count)
                    Assert.AreEqual(0, .DecList(0))
                    Assert.AreEqual(0, .DecList(1))
                    Assert.AreEqual(2, .NullableDecList.Count)
                    Assert.AreEqual(0, .NullableDecList(0), "null可でも、0で初期化")
                    Assert.AreEqual(0, .NullableDecList(1))
                End With
            End Sub

            <Test()> Public Sub FixedString_Voの内容を変更して_ApplyVoToFixedStringで_FixedStringに反映する()
                With broker.Vo
                    .DecList(0) = 1.23D
                    .DecList(1) = 45.6D
                    .NullableDecList(0) = 12.3D
                    .NullableDecList(1) = 4.56D
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("123" & "560" & "123" & "045", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_FixedStringの内容を変更して_ApplyFixedStringToVoで_Voに反映する()
                broker.FixedString = "123" & "560" & "123" & "045"
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(1.23D, .DecList(0))
                    Assert.AreEqual(5.6D, .DecList(1))
                    Assert.AreEqual(12.3D, .NullableDecList(0))
                    Assert.AreEqual(4.5D, .NullableDecList(1))
                End With
            End Sub

        End Class

        Public Class String型配列で繰り返しTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of TypeVo)

            Public Overrides Sub SetUp()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.HankakuRepeat(vo.StrArray, 3, 3))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("   " & "   " & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(3, .StrArray.Length)
                    Assert.AreEqual("", .StrArray(0))
                    Assert.AreEqual("", .StrArray(1))
                    Assert.AreEqual("", .StrArray(2))
                End With
            End Sub

            <Test()> Public Sub FixedString_Voの内容を変更して_ApplyVoToFixedStringで_FixedStringに反映する()
                With broker.Vo
                    .StrArray(0) = "1234"
                    .StrArray(1) = Nothing
                    .StrArray(2) = "b"
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("123" & "   " & "b  ", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_FixedStringの内容を変更して_ApplyFixedStringToVoで_Voに反映する()
                broker.FixedString = "123" & "   " & "b  "
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual("123", .StrArray(0))
                    Assert.AreEqual("", .StrArray(1))
                    Assert.AreEqual("b", .StrArray(2))
                End With
            End Sub

            <Test()> Public Sub ApplyFixedStringToVo_配列長さの違う別インスタンスにしても_正しい長さの配列になる()
                broker.Vo.StrArray = New String() {}

                broker.FixedString = "123" & "   " & "b  "
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(3, .StrArray.Length)
                    Assert.AreEqual("123", .StrArray(0))
                    Assert.AreEqual("", .StrArray(1))
                    Assert.AreEqual("b", .StrArray(2))
                End With
            End Sub

        End Class

        Public Class ListOfString型で繰り返しTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of TypeVo)

            Public Overrides Sub SetUp()
                broker = New FixedVoBroker(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) defineBy.HankakuRepeat(vo.StrList, 3, 3))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("   " & "   " & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(3, .StrList.Count)
                    Assert.AreEqual("", .StrList(0))
                    Assert.AreEqual("", .StrList(1))
                    Assert.AreEqual("", .StrList(2))
                End With
            End Sub

            <Test()> Public Sub FixedString_Voの内容を変更して_ApplyVoToFixedStringで_FixedStringに反映する()
                With broker.Vo
                    .StrList(0) = "b"
                    .StrList(1) = Nothing
                    .StrList(2) = "4567"
                End With
                broker.ApplyVoToFixedString()

                Assert.AreEqual("b  " & "   " & "456", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_FixedStringの内容を変更して_ApplyFixedStringToVoで_Voに反映する()
                broker.FixedString = "b  " & "   " & "456"
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual("b", .StrList(0))
                    Assert.AreEqual("", .StrList(1))
                    Assert.AreEqual("456", .StrList(2))
                End With
            End Sub

            <Test()> Public Sub ApplyFixedStringToVo_List値の長さを変更してもRepeat数になる()
                broker.Vo.StrList = New List(Of String)

                broker.FixedString = "b  " & "   " & "456"
                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(3, .StrList.Count)
                    Assert.AreEqual("b", .StrList(0))
                    Assert.AreEqual("", .StrList(1))
                    Assert.AreEqual("456", .StrList(2))
                End With
            End Sub

        End Class

        Public Class VoのTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of ExTypeVo)

            Public Overrides Sub SetUp()
                Dim subRule As New FixedRule(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                            defineBy.Hankaku(vo.Str, 2).Number(vo.Int, 2, 0).Hankaku(vo.StrPvo, 2).Number(vo.IntPvo, 2, 0).Number(vo.NullableDec, 3, 2))
                broker = New FixedVoBroker(Of ExTypeVo)(Function(defineBy As IFixedRuleLocator, vo As ExTypeVo) _
                                                          defineBy.Group(vo.TypeVoChild, subRule))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("  " & "00" & "  " & "  " & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual("", .TypeVoChild.Str)
                    Assert.AreEqual(0, .TypeVoChild.Int)
                    Assert.IsNull(.TypeVoChild.StrPvo)
                    Assert.IsNull(.TypeVoChild.IntPvo)
                    Assert.IsNull(.TypeVoChild.NullableDec)
                End With
            End Sub

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する()
                With broker.Vo
                    .TypeVoChild = New TypeVo() With {.Str = "ab", .Int = 12, .StrPvo = New StringPvo("cd"), .IntPvo = New IntegerPvo(89), .NullableDec = 4.56D}
                End With

                broker.ApplyVoToFixedString()

                Assert.AreEqual("ab" & "12" & "cd" & "89" & "456" _
                                , broker.FixedString)
            End Sub

            <Test()> Public Sub ApplyFixedStringToVo_固定長文字列の内容をVoに反映する()
                broker.FixedString = "ab" & "12" & "cd" & "89" & "456"

                broker.ApplyFixedStringToVo()

                With broker.Vo
                    With .TypeVoChild
                        Assert.AreEqual("ab", .Str)
                        Assert.AreEqual(12, .Int)
                        Assert.That(.StrPvo, [Is].EqualTo(New StringPvo("cd")))
                        Assert.That(.IntPvo, [Is].EqualTo(New IntegerPvo(89)))
                        Assert.AreEqual(4.56, .NullableDec)
                    End With
                End With
            End Sub

            <Test()> Public Sub ApplyFixedStringToVo_VoにNullを設定されてても新たにインスタンスを作成し直す()
                broker.Vo.TypeVoChild = Nothing

                broker.FixedString = "ab" & "12" & "cd" & "89" & "456"

                broker.ApplyFixedStringToVo()

                With broker.Vo
                    With .TypeVoChild
                        Assert.AreEqual("ab", .Str)
                        Assert.AreEqual(12, .Int)
                        Assert.That(.StrPvo, [Is].EqualTo(New StringPvo("cd")))
                        Assert.That(.IntPvo, [Is].EqualTo(New IntegerPvo(89)))
                        Assert.AreEqual(4.56, .NullableDec)
                    End With
                End With
            End Sub
        End Class

        Public Class Vo型配列で繰り返しTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of ExTypeVo)

            Public Overrides Sub SetUp()
                Dim subRule As New FixedRule(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                            defineBy.Hankaku(vo.Str, 2).Number(vo.Int, 2, 0).Number(vo.NullableDec, 3, 2))
                broker = New FixedVoBroker(Of ExTypeVo)(Function(defineBy As IFixedRuleLocator, vo As ExTypeVo) _
                                                          defineBy.GroupRepeat(vo.TypeVoArray, subRule, 3))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("  " & "00" & "   " & "  " & "00" & "   " & "  " & "00" & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(3, .TypeVoArray.Length)
                    Assert.AreEqual("", .TypeVoArray(0).Str)
                    Assert.AreEqual(0, .TypeVoArray(0).Int)
                    Assert.IsNull(.TypeVoArray(0).NullableDec)
                    Assert.AreEqual("", .TypeVoArray(1).Str)
                    Assert.AreEqual("", .TypeVoArray(2).Str)
                End With
            End Sub

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する()
                With broker.Vo
                    .TypeVoArray = New TypeVo() {New TypeVo() With {.Str = "ab", .Int = 12, .NullableDec = 4.56D}, _
                                                 New TypeVo(), _
                                                 New TypeVo()}
                End With

                broker.ApplyVoToFixedString()

                Assert.AreEqual("ab" & "12" & "456" _
                                & "  " & "00" & "   " _
                                & "  " & "00" & "   " _
                                , broker.FixedString)
            End Sub

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する_要素が少ない1のとき_補填される()
                With broker.Vo
                    .TypeVoArray = New TypeVo() {New TypeVo With {.Str = "ab", .Int = 12, .NullableDec = 4.56D}}
                End With

                broker.ApplyVoToFixedString()

                Assert.AreEqual("ab" & "12" & "456" _
                                & "  " & "00" & "   " _
                                & "  " & "00" & "   " _
                                , broker.FixedString, "2番目3番目の要素は補填される")
            End Sub

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する_要素が多いときは例外()
                With broker.Vo
                    .TypeVoArray = New TypeVo() {New TypeVo With {.Str = "ab"}, _
                                                 New TypeVo With {.Str = "cd"}, _
                                                 New TypeVo With {.Str = "ef"}, _
                                                 New TypeVo With {.Str = "gh"}}
                End With

                Try
                    broker.ApplyVoToFixedString()
                    Assert.Fail("例外になるべき")
                Catch actual As InvalidOperationException
                    Assert.AreEqual("固定長サイズを超えた箇所に値を設定できない. 繰返し数は適切か？ name='TypeVoArray[3].Str'", actual.Message)
                End Try
            End Sub

            <Test()> Public Sub ApplyFixedStringToVo_固定長文字列の内容をVoに反映する()
                broker.FixedString = "ab" & "12" & "456" _
                                    & "  " & "00" & "   " _
                                    & "  " & "00" & "   "

                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(3, .TypeVoArray.Length)
                    With .TypeVoArray(0)
                        Assert.AreEqual("ab", .Str)
                        Assert.AreEqual(12, .Int)
                        Assert.AreEqual(4.56, .NullableDec)
                    End With
                    With .TypeVoArray(1)
                        Assert.AreEqual("", .Str)
                        Assert.AreEqual(0, .Int)
                        Assert.IsNull(.NullableDec)
                    End With
                    With .TypeVoArray(2)
                        Assert.AreEqual("", .Str)
                        Assert.AreEqual(0, .Int)
                        Assert.IsNull(.NullableDec)
                    End With
                End With
            End Sub

        End Class

        Public Class List型Voで繰り返しTest : Inherits FixedVoBrokerTest

            Private broker As FixedVoBroker(Of ExTypeVo)

            Public Overrides Sub SetUp()
                Dim subRule As New FixedRule(Of TypeVo)(Function(defineBy As IFixedRuleLocator, vo As TypeVo) _
                                                            defineBy.Hankaku(vo.Str, 2).Number(vo.Int, 2, 0).Number(vo.NullableDec, 3, 2))
                broker = New FixedVoBroker(Of ExTypeVo)(Function(defineBy As IFixedRuleLocator, vo As ExTypeVo) _
                                                          defineBy.GroupRepeat(vo.TypeVoList, subRule, 3))
            End Sub

            <Test()> Public Sub FixedString_コンストラクタ直後_はnull可の数値項目なら空文字で初期化()
                Assert.AreEqual("  " & "00" & "   " & "  " & "00" & "   " & "  " & "00" & "   ", broker.FixedString)
            End Sub

            <Test()> Public Sub Vo_コンストラクタ直後_はVo自身のインスタンス生成直後と同じ()
                With broker.Vo
                    Assert.AreEqual(3, .TypeVoList.Count)
                    Assert.AreEqual("", .TypeVoList(0).Str)
                    Assert.AreEqual(0, .TypeVoList(0).Int)
                    Assert.IsNull(.TypeVoList(0).NullableDec)
                    Assert.AreEqual("", .TypeVoList(1).Str)
                    Assert.AreEqual("", .TypeVoList(2).Str)
                End With
            End Sub

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する()
                With broker.Vo
                    .TypeVoList = EzUtil.NewList(New TypeVo() With {.Str = "ab", .Int = 12, .NullableDec = 4.56D}, _
                                                 New TypeVo(), _
                                                 New TypeVo())
                End With

                broker.ApplyVoToFixedString()

                Assert.AreEqual("ab" & "12" & "456" _
                                & "  " & "00" & "   " _
                                & "  " & "00" & "   " _
                                , broker.FixedString)
            End Sub

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する_要素が少ない1のとき_補填される()
                With broker.Vo
                    .TypeVoList = EzUtil.NewList(New TypeVo() With {.Str = "ab", .Int = 12, .NullableDec = 4.56D})
                End With

                broker.ApplyVoToFixedString()

                Assert.AreEqual("ab" & "12" & "456" _
                                & "  " & "00" & "   " _
                                & "  " & "00" & "   " _
                                , broker.FixedString, "2番目3番目の要素は補填される")
            End Sub

            <Test()> Public Sub ApplyVoToFixedString_Voの内容を変更して固定長文字列に反映する_要素が多いときは例外()
                With broker.Vo
                    .TypeVoList = EzUtil.NewList(New TypeVo() With {.Str = "ab"}, _
                                                 New TypeVo() With {.Str = "cd"}, _
                                                 New TypeVo() With {.Str = "ef"}, _
                                                 New TypeVo() With {.Str = "gh"})
                End With

                Try
                    broker.ApplyVoToFixedString()
                    Assert.Fail("例外になるべき")
                Catch actual As InvalidOperationException
                    Assert.AreEqual("固定長サイズを超えた箇所に値を設定できない. 繰返し数は適切か？ name='TypeVoList[3].Str'", actual.Message)
                End Try
            End Sub

            <Test()> Public Sub ApplyFixedStringToVo_固定長文字列の内容をVoに反映する()
                broker.FixedString = "ab" & "12" & "456" _
                                    & "  " & "00" & "   " _
                                    & "  " & "00" & "   "

                broker.ApplyFixedStringToVo()

                With broker.Vo
                    Assert.AreEqual(3, .TypeVoList.Count)
                    With .TypeVoList(0)
                        Assert.AreEqual("ab", .Str)
                        Assert.AreEqual(12, .Int)
                        Assert.AreEqual(4.56, .NullableDec)
                    End With
                    With .TypeVoList(1)
                        Assert.AreEqual("", .Str)
                        Assert.AreEqual(0, .Int)
                        Assert.IsNull(.NullableDec)
                    End With
                    With .TypeVoList(2)
                        Assert.AreEqual("", .Str)
                        Assert.AreEqual(0, .Int)
                        Assert.IsNull(.NullableDec)
                    End With
                End With
            End Sub
        End Class

        Public Class 装飾のテスト : Inherits FixedVoBrokerTest

            <Test()>
            Public Sub DecorateToVoを指定した場合_固定長文字列からVoへの変換時に指定した装飾処理が行われる()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 3, decorateToVo:=Function(str) str.Replace("12", "54")))

                broker.FixedString = "123"

                broker.ApplyFixedStringToVo()

                Assert.That(broker.Vo.First, [Is].EqualTo("543"))
            End Sub

            <Test()>
            Public Sub DecorateToStringを指定した場合_Voから固定長文字列への変換時に指定した装飾処理が行われる()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 10, decorateToString:=Function(v) StringUtil.ToUpper(StringUtil.ToString(v))))

                broker.Vo.First = "piyo"

                broker.ApplyVoToFixedString()

                Assert.That(broker.FixedString, [Is].EqualTo("PIYO      "), "自動的に空白は埋められる")
            End Sub

            <Test()>
            Public Sub 長さを指定しなければ_残りの文字列を切り取ることができる()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 4).Hankaku(vo.Second))

                broker.FixedString = "hogefugapiyo"

                broker.ApplyFixedStringToVo()

                Assert.That(broker.Vo.First, [Is].EqualTo("hoge"))
                Assert.That(broker.Vo.Second, [Is].EqualTo("fugapiyo"))
            End Sub

            <Test()>
            Public Sub 長さが指定されていないプロパティに対応する位置に文字列がない場合_空文字がセットされる()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 4).Hankaku(vo.Second))

                broker.FixedString = "hoge"

                broker.ApplyFixedStringToVo()

                Assert.That(broker.Vo.First, [Is].EqualTo("hoge"))
                Assert.That(broker.Vo.Second, [Is].EqualTo(""))
            End Sub

            <Test()>
            Public Sub 値は空白文字がトリムされた状態でVoに格納される()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 6).Hankaku(vo.Second))

                broker.FixedString = "hoge  fuga  piyo  "

                broker.ApplyFixedStringToVo()

                Assert.That(broker.Vo.First, [Is].EqualTo("hoge"))
                Assert.That(broker.Vo.Second, [Is].EqualTo("fuga  piyo"))
            End Sub

            <Test()>
            Public Sub 長さを指定してなくても_DecorateToVoを指定すれば装飾処理が行われる()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 6).Hankaku(vo.Second, decorateToVo:=Function(str) StringUtil.ToUpper(str) & "  suffix"))

                broker.FixedString = "hoge  fuga  piyo  "

                broker.ApplyFixedStringToVo()

                Assert.That(broker.Vo.First, [Is].EqualTo("hoge"))
                Assert.That(broker.Vo.Second, [Is].EqualTo("FUGA  PIYO  suffix"))
            End Sub

            <Test()>
            Public Sub Voを固定長文字列に変換できる()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 6).Hankaku(vo.Second, 6).Hankaku(vo.Third, 6))

                broker.Vo.First = "hoge"
                broker.Vo.Second = "fuga"
                broker.Vo.Third = "piyo"

                broker.ApplyVoToFixedString()

                Assert.That(broker.FixedString, [Is].EqualTo("hoge  fuga  piyo  "))
            End Sub

            <Test()>
            Public Sub 最後の桁数を指定しなければ_そのプロパティの値を出力する()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 6).Hankaku(vo.Second, 6).Hankaku(vo.Third))

                broker.Vo.First = "hoge"
                broker.Vo.Second = "fuga"
                broker.Vo.Third = "piyo"

                broker.ApplyVoToFixedString()

                Assert.That(broker.FixedString, [Is].EqualTo("hoge  fuga  piyo"))
            End Sub

            <Test()>
            Public Sub プロパティに値がなければ_長さ0の文字列として固定長文字列が出力される()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 6).Hankaku(vo.Second))

                broker.Vo.First = "hoge"
                broker.Vo.Second = Nothing

                broker.ApplyVoToFixedString()

                Assert.That(broker.FixedString, [Is].EqualTo("hoge  "))
            End Sub

            <Test()>
            Public Sub 長さを指定してなくても_DecorateToStringを指定すれば装飾処理が行われる()
                Dim broker As New FixedVoBroker(Of MyFixedStringVo)(Function(defineBy As IFixedRuleLocator, vo As MyFixedStringVo) _
                                                            defineBy.Hankaku(vo.First, 6).Hankaku(vo.Second, decorateTostring:=Function(v) "foo"))

                broker.Vo.First = "hoge"
                broker.Vo.Second = "fuga"

                broker.ApplyVoToFixedString()

                Assert.That(broker.FixedString, [Is].EqualTo("hoge  foo"))
            End Sub

        End Class

    End Class
End Namespace