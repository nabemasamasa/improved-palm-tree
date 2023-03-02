Imports Fhi.Fw.TestUtil.DebugString
Imports NUnit.Framework

Namespace Util.Csv
    Public MustInherit Class CsvVoBrokerTest

#Region "CsvVo"
        Private Class SimpleVo
            Private _id As Integer?
            Private _name As String
            Private _aDouble As Double?
            Private _aDecimal As Decimal?
            Private _aDate As DateTime?

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

            Public Property ADouble() As Double?
                Get
                    Return _aDouble
                End Get
                Set(ByVal value As Double?)
                    _aDouble = value
                End Set
            End Property

            Public Property ADecimal() As Decimal?
                Get
                    Return _aDecimal
                End Get
                Set(ByVal value As Decimal?)
                    _aDecimal = value
                End Set
            End Property

            Public Property ADate() As Date?
                Get
                    Return _aDate
                End Get
                Set(ByVal value As Date?)
                    _aDate = value
                End Set
            End Property
        End Class
        Private Class ParentVo
            Private _id As Integer?
            Private _child As ChildVo
            Private _children As List(Of ChildVo)

            Public Property Id() As Integer?
                Get
                    Return _id
                End Get
                Set(ByVal value As Integer?)
                    _id = value
                End Set
            End Property

            Public Property Child() As ChildVo
                Get
                    Return _child
                End Get
                Set(ByVal value As ChildVo)
                    _child = value
                End Set
            End Property

            Public Property Children() As List(Of ChildVo)
                Get
                    Return _children
                End Get
                Set(ByVal value As List(Of ChildVo))
                    _children = value
                End Set
            End Property
        End Class
        Private Class ChildVo
            Private _name As String
            Private _address As String

            Public Property Name() As String
                Get
                    Return _name
                End Get
                Set(ByVal value As String)
                    _name = value
                End Set
            End Property

            Public Property Address() As String
                Get
                    Return _address
                End Get
                Set(ByVal value As String)
                    _address = value
                End Set
            End Property
        End Class
        Private Class ParentVo2
            Private _id As Integer?
            Private _child As ChildVo
            Private _name As String

            Public Property Id() As Integer?
                Get
                    Return _id
                End Get
                Set(ByVal value As Integer?)
                    _id = value
                End Set
            End Property

            Public Property Child() As ChildVo
                Get
                    Return _child
                End Get
                Set(ByVal value As ChildVo)
                    _child = value
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
        End Class
        Private Class CollectionVo
            Private _id As Integer?
            Private _nameList As List(Of String)
            Private _nameArray As String()
            Private _address As String

            Public Property Id() As Integer?
                Get
                    Return _id
                End Get
                Set(ByVal value As Integer?)
                    _id = value
                End Set
            End Property

            Public Property NameList() As List(Of String)
                Get
                    Return _nameList
                End Get
                Set(ByVal value As List(Of String))
                    _nameList = value
                End Set
            End Property

            Public Property NameArray() As String()
                Get
                    Return _nameArray
                End Get
                Set(ByVal value As String())
                    _nameArray = value
                End Set
            End Property

            Public Property Address() As String
                Get
                    Return _address
                End Get
                Set(ByVal value As String)
                    _address = value
                End Set
            End Property
        End Class
#End Region

#Region "Rule Implements"
        Private Class SimpleVoCsvRule : Implements ICsvRule(Of SimpleVo)
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As SimpleVo) Implements ICsvRule(Of SimpleVo).Configure
                defineBy.Field(vo.Id).Field(vo.Name)
            End Sub
        End Class
        Private Class SimpleVoDoubleRule : Implements ICsvRule(Of SimpleVo)
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As SimpleVo) Implements ICsvRule(Of SimpleVo).Configure
                defineBy.Field(vo.Name).Field(vo.ADouble)
            End Sub
        End Class
        Private Class SimpleVoDecimalRule : Implements ICsvRule(Of SimpleVo)
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As SimpleVo) Implements ICsvRule(Of SimpleVo).Configure
                defineBy.Field(vo.Name).Field(vo.ADecimal)
            End Sub
        End Class
        Private Class SimpleVoDateRule : Implements ICsvRule(Of SimpleVo)
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As SimpleVo) Implements ICsvRule(Of SimpleVo).Configure
                defineBy.Field(vo.Name).Field(vo.ADate)
            End Sub
        End Class
        Private Class ParentVoNamesRule : Implements ICsvRule(Of ParentVo)
            Private ReadOnly repeat As Integer
            Public Sub New(ByVal repeat As Integer)
                Me.repeat = repeat
            End Sub
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As ParentVo) Implements ICsvRule(Of ParentVo).Configure
                defineBy.Field(vo.Id).GroupRepeat(vo.Children, New ChildVoNameRule, repeat)
            End Sub
        End Class
        Private Class ChildVoNameRule : Implements ICsvRule(Of ChildVo)
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As ChildVo) Implements ICsvRule(Of ChildVo).Configure
                defineBy.Field(vo.Name)
            End Sub
        End Class
        Private Class ParentVoNameAddressesRule : Implements ICsvRule(Of ParentVo)
            Private ReadOnly repeat As Integer
            Public Sub New(ByVal repeat As Integer)
                Me.repeat = repeat
            End Sub
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As ParentVo) Implements ICsvRule(Of ParentVo).Configure
                defineBy.Field(vo.Id).GroupRepeat(vo.Children, New ChildVoNameAddressRule, repeat)
            End Sub
        End Class
        Private Class ParentChildRule : Implements ICsvRule(Of ParentVo2)
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As ParentVo2) Implements ICsvRule(Of ParentVo2).Configure
                defineBy.Field(vo.Id).Group(vo.Child, New ChildVoNameAddressRule).Field(vo.Name)
            End Sub
        End Class
        Private Class ChildVoNameAddressRule : Implements ICsvRule(Of ChildVo)
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As ChildVo) Implements ICsvRule(Of ChildVo).Configure
                defineBy.Field(vo.Name, vo.Address)
            End Sub
        End Class
        Private Class CollectionVoListRule : Implements ICsvRule(Of CollectionVo)
            Private ReadOnly repeat As Integer
            Public Sub New(ByVal repeat As Integer)
                Me.repeat = repeat
            End Sub
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As CollectionVo) Implements ICsvRule(Of CollectionVo).Configure
                defineBy.Field(vo.Id).FieldRepeat(vo.NameList, repeat).Field(vo.Address)
            End Sub
        End Class
        Private Class CollectionVoArrayRule : Implements ICsvRule(Of CollectionVo)
            Private ReadOnly repeat As Integer
            Public Sub New(ByVal repeat As Integer)
                Me.repeat = repeat
            End Sub
            Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As CollectionVo) Implements ICsvRule(Of CollectionVo).Configure
                defineBy.Field(vo.Id).FieldRepeat(vo.NameArray, repeat).Field(vo.Address)
            End Sub
        End Class
#End Region

        Private Function NewSimple(ByVal id As Integer?, ByVal name As String) As SimpleVo
            Dim vo As New SimpleVo
            vo.Id = id
            vo.Name = name
            Return vo
        End Function

        Private Function NewParentNames(ByVal id As Integer?, ByVal ParamArray names As String()) As ParentVo
            Dim vo As New ParentVo
            vo.Id = id
            vo.Children = New List(Of ChildVo)
            For Each name As String In names
                Dim vo2 As New ChildVo
                vo2.Name = name
                vo.Children.Add(vo2)
            Next
            Return vo
        End Function

        Private Function NewCollectionListVo(ByVal id As Integer?, ByVal address As String, ByVal ParamArray names As String()) As CollectionVo
            Dim vo As New CollectionVo
            vo.Id = id
            vo.Address = address
            vo.NameList = New List(Of String)(names)
            Return vo
        End Function

        Private Function NewCollectionArrayVo(ByVal id As Integer?, ByVal address As String, ByVal ParamArray names As String()) As CollectionVo
            Dim vo As New CollectionVo
            vo.Id = id
            vo.Address = address
            vo.NameArray = names
            Return vo
        End Function

        Public Class [Default] : Inherits CsvVoBrokerTest

            <Test()> Public Sub Vos_値を設定したらCsvStringsが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule)
                broker.Vos = EzUtil.NewList(NewSimple(3, "a"), NewSimple(20, "b")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(2, broker.CsvStrings.Length)
                Assert.AreEqual("3,a", broker.CsvStrings(0))
                Assert.AreEqual("20,b", broker.CsvStrings(1))
            End Sub

            <Test()> Public Sub CsvStrings_値を設定したらVosが出来る_Integer()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule)
                broker.CsvStrings = New String() {"-6,a", _
                                                  "12,b"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(2, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(-6, .Id)
                    Assert.AreEqual("a", .Name)
                End With
                With broker.Vos(1)
                    Assert.AreEqual(12, .Id)
                    Assert.AreEqual("b", .Name)
                End With
            End Sub

            <Test()> Public Sub Vos_値を設定したらCsvStringsが出来る_ラムダ式()
                Dim broker As New CsvVoBroker(Of SimpleVo)(Function(defineBy As ICsvRuleLocator, vo As SimpleVo) defineBy.Field(vo.Id).Field(vo.Name))
                broker.Vos = EzUtil.NewList(NewSimple(3, "a"), NewSimple(20, "b")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(2, broker.CsvStrings.Length)
                Assert.AreEqual("3,a", broker.CsvStrings(0))
                Assert.AreEqual("20,b", broker.CsvStrings(1))
            End Sub

            <Test()> Public Sub CsvStrings_値を設定したらVosが出来る_ラムダ式_Integer()
                Dim broker As New CsvVoBroker(Of SimpleVo)(Function(defineBy As ICsvRuleLocator, vo As SimpleVo) defineBy.Field(vo.Id).Field(vo.Name))
                broker.CsvStrings = New String() {"-6,a", _
                                                  "12,b"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(2, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(-6, .Id)
                    Assert.AreEqual("a", .Name)
                End With
                With broker.Vos(1)
                    Assert.AreEqual(12, .Id)
                    Assert.AreEqual("b", .Name)
                End With
            End Sub

            <Test()> Public Sub CsvStrings_値を設定したらVosが出来る_Double()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoDoubleRule)
                broker.CsvStrings = New String() {"a,3.14", _
                                                  "b,1.4142"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(2, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual("a", .Name)
                    Assert.AreEqual(3.14#, .ADouble)
                End With
                With broker.Vos(1)
                    Assert.AreEqual("b", .Name)
                    Assert.AreEqual(1.4142#, .ADouble)
                End With
            End Sub

            <Test()> Public Sub CsvStrings_値を設定したらVosが出来る_Decimal()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoDecimalRule)
                broker.CsvStrings = New String() {"a,3.141592653589", _
                                                  "b,1.41421356"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(2, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual("a", .Name)
                    Assert.AreEqual(3.141592653589@, .ADecimal)
                End With
                With broker.Vos(1)
                    Assert.AreEqual("b", .Name)
                    Assert.AreEqual(1.41421356@, .ADecimal)
                End With
            End Sub

            <Test()> Public Sub CsvStrings_値を設定したらVosが出来る_DateTime()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoDateRule)
                broker.CsvStrings = New String() {"a,2011/09/23", _
                                                  "b,21:47:38", _
                                                  "c,2011/09/22 11:22:33"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(3, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual("a", .Name)
                    Assert.AreEqual(CDate("2011/09/23"), .ADate)
                End With
                With broker.Vos(1)
                    Assert.AreEqual("b", .Name)
                    Assert.AreEqual(CDate("21:47:38"), .ADate)
                End With
                With broker.Vos(2)
                    Assert.AreEqual("c", .Name)
                    Assert.AreEqual(CDate("2011/09/22 11:22:33"), .ADate)
                End With
            End Sub

            <Test()> Public Sub Vos_値を設定したらCsvStringsが出来る_タブ区切り()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.TAB)
                broker.Vos = EzUtil.NewList(NewSimple(3, "a"), NewSimple(20, "b")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(2, broker.CsvStrings.Length)
                Assert.AreEqual("3" & vbTab & "a", broker.CsvStrings(0))
                Assert.AreEqual("20" & vbTab & "b", broker.CsvStrings(1))
            End Sub

            <Test()> Public Sub CsvStrings_値を設定したらVosが出来る_タブ区切り()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.TAB)
                broker.CsvStrings = New String() {"-6" & vbTab & "a", _
                                                  "12" & vbTab & "b"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(2, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(-6, .Id)
                    Assert.AreEqual("a", .Name)
                End With
                With broker.Vos(1)
                    Assert.AreEqual(12, .Id)
                    Assert.AreEqual("b", .Name)
                End With
            End Sub

            Private Const DQ As String = """"

            <Test()> Public Sub CsvStrings_値をダブルコートで囲った空文字列を設定したら_囲まれてない空文字列のVosが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)

                broker.CsvStrings = New String() {"0," & DQ & "" & DQ}
                Assert.AreEqual("", broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値をダブルコートで囲った文字列を設定したら_囲まれてない文字列のVosが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)

                broker.CsvStrings = New String() {"0," & DQ & "A" & DQ}
                Assert.AreEqual("A", broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値をダブルコートで囲ったカンマ入り文字列を設定したら_囲まれてないカンマ入りのVosが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)

                broker.CsvStrings = New String() {"0," & DQ & "A,B" & DQ}
                Assert.AreEqual("A,B", broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値をダブルコートで囲ったダブルコート入り文字列を設定したら_囲まれてないダブルコート入りのVosが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)

                broker.CsvStrings = New String() {"0," & DQ & "A" & DQ & DQ & DQ}
                Assert.AreEqual("A" & DQ, broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値をダブルコートで囲ったダブルコート入り文字列を設定したら_囲まれてないダブルコート入りのVosが出来る_2個()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)

                broker.CsvStrings = New String() {"0," & DQ & "A" & DQ & DQ & DQ & DQ & DQ}
                Assert.AreEqual("A" & DQ & DQ, broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値をダブルコートで囲って_カンマの前にダブルコートを入れた文字列を設定しても_そのカンマ前ダブルコートを終端と誤らずに_Vosが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)

                broker.CsvStrings = New String() {"0," & DQ & "A" & DQ & DQ & ",B" & DQ}
                Assert.AreEqual("A" & DQ & ",B", broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値をダブルコートで囲った改行入り文字列を設定したら_囲まれてない改行入り文字列のVosが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)

                broker.CsvStrings = New String() {"0," & DQ & "A" & vbCrLf & ",B" & DQ}
                Assert.AreEqual("A" & vbCrLf & ",B", broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値変換_数値項目が空なら値はNothing()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {",aaa"}
                Assert.AreEqual(Nothing, broker.Vos(0).Id)
            End Sub

            <Test()> Public Sub CsvStrings_値変換_数値項目が空白でも値はNothing()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"   ,bbb"}
                Assert.AreEqual(Nothing, broker.Vos(0).Id)
            End Sub

            <Test()> Public Sub CsvStrings_値変換_数値項目が文字なら値はNothing()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"ccc,ddd"}
                Assert.AreEqual(Nothing, broker.Vos(0).Id)
            End Sub

            <Test()> Public Sub CsvStrings_値変換_String項目が空なら値は空文字()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"1,"}
                Assert.AreEqual("", broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値変換_String項目が空白でも値は空文字()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"1,    "}
                Assert.AreEqual("", broker.Vos(0).Name)
            End Sub

            <Test()> Public Sub CsvStrings_値変換_日付項目が空なら値はNothing()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoDateRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"1,aaa"}
                Assert.AreEqual(Nothing, broker.Vos(0).ADate)
            End Sub

            <Test()> Public Sub CsvStrings_値変換_日付項目が空白でも値はNothing()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoDateRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"1,   "}
                Assert.AreEqual(Nothing, broker.Vos(0).ADate)
            End Sub

            <Test()> Public Sub CsvStrings_値変換_日付項目が文字なら値はNothing()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoDateRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"1,ddd"}
                Assert.AreEqual(Nothing, broker.Vos(0).ADate)
            End Sub

            <Test()> Public Sub Vosで_区切り文字を含む文字列をを設定したら_ダブルコートで囲んだ区切り文字入りCsvStringsが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)
                broker.Vos = EzUtil.NewList(NewSimple(3, "a,b"), NewSimple(20, "x" & vbTab & "y")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(2, broker.CsvStrings.Length)
                Assert.AreEqual("3," & DQ & "a,b" & DQ, broker.CsvStrings(0))
                Assert.AreEqual("20," & "x" & vbTab & "y", broker.CsvStrings(1))
            End Sub

            <Test()> Public Sub Vosで_改行を含む文字列をを設定したら_ダブルコートで囲んだ改行入りCsvStringsが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.COMMA)
                broker.Vos = EzUtil.NewList(NewSimple(3, "a" & vbCrLf & "b"), NewSimple(20, "x" & vbCrLf & "y")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(2, broker.CsvStrings.Length)
                Assert.AreEqual("3," & DQ & "a" & vbCrLf & "b" & DQ, broker.CsvStrings(0))
                Assert.AreEqual("20," & DQ & "x" & vbCrLf & "y" & DQ, broker.CsvStrings(1))
            End Sub

            <Test()> Public Sub Vosで_区切り文字を含む文字列をを設定したら_ダブルコートで囲んだ区切り文字入りCsvStringsが出来る_TAB()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.TAB)
                broker.Vos = EzUtil.NewList(NewSimple(3, "a" & vbTab & "b"), NewSimple(20, "x,y")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(2, broker.CsvStrings.Length)
                Assert.AreEqual("3" & vbTab & DQ & "a" & vbTab & "b" & DQ, broker.CsvStrings(0))
                Assert.AreEqual("20" & vbTab & "x,y", broker.CsvStrings(1))
            End Sub

            <Test()> Public Sub GroupRuleで繰り返し_Vosで_繰り返し値()
                Dim broker As New CsvVoBroker(Of ParentVo)(New ParentVoNamesRule(3), CsvSeparator.COMMA)
                broker.Vos = EzUtil.NewList(NewParentNames(2, "a", "b", "c")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length)
                Assert.AreEqual("2,a,b,c", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub GroupRuleで繰り返し_Vosで_繰り返し値_ラムダ式()
                Dim broker As New CsvVoBroker(Of ParentVo)(Function(define, vo) define.Field(vo.Id).GroupRepeat(Of ChildVo)(vo.Children, Function(define1, vo1) define1.Field(vo1.Name), 3), CsvSeparator.COMMA)

                broker.Vos = EzUtil.NewList(NewParentNames(2, "a", "b", "c")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length)
                Assert.AreEqual("2,a,b,c", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub GroupRuleで繰り返し_Vosで_繰り返し値_繰返し数に満たなければ補完される()
                Dim broker As New CsvVoBroker(Of ParentVo)(New ParentVoNamesRule(4), CsvSeparator.COMMA)
                broker.Vos = EzUtil.NewList(NewParentNames(3, "x", "y")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length)
                Assert.AreEqual("3,x,y,,", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub GroupRuleで繰り返し_CsvStringsで_繰り返し値_単一値()
                Dim broker As New CsvVoBroker(Of ParentVo)(New ParentVoNamesRule(3), CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"4,l,m,n"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(1, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(4, .Id)
                    Assert.AreEqual(3, .Children.Count)
                    Assert.AreEqual("l", .Children(0).Name)
                    Assert.AreEqual("m", .Children(1).Name)
                    Assert.AreEqual("n", .Children(2).Name)
                End With
            End Sub

            <Test()> Public Sub GroupRuleで繰り返し_CsvStringsで_繰り返し値_複数値()
                Dim broker As New CsvVoBroker(Of ParentVo)(New ParentVoNameAddressesRule(2), CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"5,a,b,c,d"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(1, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(5, .Id)
                    Assert.AreEqual(2, .Children.Count)
                    Assert.AreEqual("a", .Children(0).Name)
                    Assert.AreEqual("b", .Children(0).Address)
                    Assert.AreEqual("c", .Children(1).Name)
                    Assert.AreEqual("d", .Children(1).Address)
                End With
            End Sub

            <Test()> Public Sub RepeatRuleで繰り返し_Vosで_繰り返し値_List複数値()
                Dim broker As New CsvVoBroker(Of CollectionVo)(New CollectionVoListRule(3), CsvSeparator.COMMA)
                broker.Vos = EzUtil.NewList(NewCollectionListVo(7, "a", "x", "y", "z")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length)
                Assert.AreEqual("7,x,y,z,a", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub RepeatRuleで繰り返し_Vosで_繰り返し値_List複数値_繰返し数を超えた要素は無視される()
                Dim broker As New CsvVoBroker(Of CollectionVo)(New CollectionVoListRule(2), CsvSeparator.COMMA)
                broker.Vos = EzUtil.NewList(NewCollectionListVo(7, "a", "n1", "n2", "n3", "n4")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length)
                Assert.AreEqual("7,n1,n2,a", broker.CsvStrings(0), "4要素渡されても繰返し数は2なので")
            End Sub

            <Test()> Public Sub RepeatRuleで繰り返し_CsvStringsで_繰り返し値_List複数値()
                Dim broker As New CsvVoBroker(Of CollectionVo)(New CollectionVoListRule(3), CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"5,h,i,j,k"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(1, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(5, .Id)
                    Assert.AreEqual(Nothing, .NameArray)
                    Assert.AreEqual(3, .NameList.Count)
                    Assert.AreEqual("h", .NameList(0))
                    Assert.AreEqual("i", .NameList(1))
                    Assert.AreEqual("j", .NameList(2))
                    Assert.AreEqual("k", .Address)
                End With
            End Sub

            <Test()> Public Sub RepeatRuleで繰り返し_Vosで_繰り返し値_配列複数値()
                Dim broker As New CsvVoBroker(Of CollectionVo)(New CollectionVoArrayRule(3), CsvSeparator.COMMA)
                broker.Vos = EzUtil.NewList(NewCollectionArrayVo(7, "a", "r", "y", "z")).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length)
                Assert.AreEqual("7,r,y,z,a", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub RepeatRuleで繰り返し_CsvStringsで_繰り返し値_配列複数値()
                Dim broker As New CsvVoBroker(Of CollectionVo)(New CollectionVoArrayRule(3), CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"5,l,m,n,o"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(1, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(5, .Id)
                    Assert.AreEqual(Nothing, .NameList)
                    Assert.AreEqual(3, .NameArray.Length)
                    Assert.AreEqual("l", .NameArray(0))
                    Assert.AreEqual("m", .NameArray(1))
                    Assert.AreEqual("n", .NameArray(2))
                    Assert.AreEqual("o", .Address)
                End With
            End Sub

            <Test()> Public Sub CsvStrings_行によって個数の違う値を設定してもVosが出来る()
                Dim broker As New CsvVoBroker(Of SimpleVo)(New SimpleVoCsvRule, CsvSeparator.TAB)
                broker.CsvStrings = New String() {"-1" & vbTab & "n1", _
                                                  "-2"}
                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(2, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(-1, .Id)
                    Assert.AreEqual("n1", .Name)
                End With
                With broker.Vos(1)
                    Assert.AreEqual(-2, .Id)
                    Assert.AreEqual(Nothing, .Name)
                End With
            End Sub

            <Test()> Public Sub CsvStrings_Voを直接内包するCSV定義()

                Dim broker As New CsvVoBroker(Of ParentVo2)(New ParentChildRule, CsvSeparator.COMMA)
                broker.CsvStrings = New String() {"-1,a,b,c"}

                Assert.IsNotNull(broker.Vos)
                Assert.AreEqual(1, broker.Vos.Length)
                With broker.Vos(0)
                    Assert.AreEqual(-1, .Id)
                    Assert.AreEqual("a", .Child.Name)
                    Assert.AreEqual("b", .Child.Address)
                    Assert.AreEqual("c", .Name)
                End With
            End Sub

            <Test()> Public Sub Vos_Voを直接内包するCSV定義()

                Dim broker As New CsvVoBroker(Of ParentVo2)(New ParentChildRule, CsvSeparator.COMMA)
                Dim parent As New ParentVo2 With {.Id = 2, .Name = "hoge", .Child = New ChildVo With {.Name = "fuga", .Address = "piyo"}}
                broker.Vos = EzUtil.NewList(parent).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length)
                Assert.AreEqual("2,fuga,piyo,hoge", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub Vos_Voを直接内包するCSV定義_ラムダ式()

                Dim broker As New CsvVoBroker(Of ParentVo2)(Function(define, vo) define.Field(vo.Id).Group(Of ChildVo)(vo.Child, Function(define1, vo1) define1.Field(vo1.Name, vo1.Address)).Field(vo.Name), CsvSeparator.COMMA)
                Dim parent As New ParentVo2 With {.Id = 2, .Name = "hoge", .Child = New ChildVo With {.Name = "fuga", .Address = "piyo"}}
                broker.Vos = EzUtil.NewList(parent).ToArray

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length)
                Assert.AreEqual("2,fuga,piyo,hoge", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub IsOutputTitle_タイトル出力する()
                Dim broker As New CsvVoBroker(Of ParentVo2)(Function(defineBy As ICsvRuleLocator, vo As ParentVo2) defineBy.Field(vo.Id, vo.Name), CsvSeparator.COMMA)
                broker.IsOutputTitle = True
                broker.Vos = New ParentVo2() {}

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length, "データ0行だからタイトル行のみ")
                Assert.AreEqual("Id,Name", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub IsOutputTitle_繰り返し単項目もタイトル出力する_配列_繰返し数1()
                Dim broker As New CsvVoBroker(Of CollectionVo)(Function(defineBy As ICsvRuleLocator, vo As CollectionVo) defineBy.Field(vo.Id).FieldRepeat(vo.NameArray, 1), CsvSeparator.COMMA)
                broker.IsOutputTitle = True
                broker.Vos = New CollectionVo() {}

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length, "データ0行だからタイトル行のみ")
                Assert.AreEqual("Id,NameArray(0)", broker.CsvStrings(0), "繰返し数1でも添え字が付く")
            End Sub

            <Test()> Public Sub IsOutputTitle_繰り返し単項目もタイトル出力する_配列_繰返し数2()
                Dim broker As New CsvVoBroker(Of CollectionVo)(Function(defineBy As ICsvRuleLocator, vo As CollectionVo) defineBy.Field(vo.Id).FieldRepeat(vo.NameArray, 2), CsvSeparator.COMMA)
                broker.IsOutputTitle = True
                broker.Vos = New CollectionVo() {}

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length, "データ0行だからタイトル行のみ")
                Assert.AreEqual("Id,NameArray(0),NameArray(1)", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub IsOutputTitle_繰り返し単項目もタイトル出力する_List()
                Dim broker As New CsvVoBroker(Of CollectionVo)(Function(defineBy As ICsvRuleLocator, vo As CollectionVo) defineBy.Field(vo.Id).FieldRepeat(vo.NameList, 3), CsvSeparator.COMMA)
                broker.IsOutputTitle = True
                broker.Vos = New CollectionVo() {}

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length, "データ0行だからタイトル行のみ")
                Assert.AreEqual("Id,NameList(0),NameList(1),NameList(2)", broker.CsvStrings(0))
            End Sub

            <Test()> Public Sub IsOutputTitle_グループもタイトル出力する()
                Dim childRule As New CsvRule(Of ChildVo)(Function(defineBy As ICsvRuleLocator, vo As ChildVo) defineBy.Field(vo.Address, vo.Name))
                Dim broker As New CsvVoBroker(Of ParentVo)(Function(defineBy As ICsvRuleLocator, vo As ParentVo) defineBy.Field(vo.Id).Group(vo.Child, childRule), CsvSeparator.COMMA)
                broker.IsOutputTitle = True
                broker.Vos = New ParentVo() {}

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length, "データ0行だからタイトル行のみ")
                Assert.AreEqual("Id,Child.Address,Child.Name", broker.CsvStrings(0), "繰り返し設定は無いから、添え字はつかない")
            End Sub

            <Test()> Public Sub IsOutputTitle_繰り返し一グループもタイトル出力する_繰返し数1()
                Dim childrenRule As New CsvRule(Of ChildVo)(Function(defineBy As ICsvRuleLocator, vo As ChildVo) defineBy.Field(vo.Address, vo.Name))
                Dim broker As New CsvVoBroker(Of ParentVo)(Function(defineBy As ICsvRuleLocator, vo As ParentVo) defineBy.Field(vo.Id).GroupRepeat(vo.Children, childrenRule, 1), CsvSeparator.COMMA)
                broker.IsOutputTitle = True
                broker.Vos = New ParentVo() {}

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length, "データ0行だからタイトル行のみ")
                Assert.AreEqual("Id,Children(0).Address,Children(0).Name", broker.CsvStrings(0), "繰返し数1でも添え字が付く")
            End Sub

            <Test()> Public Sub IsOutputTitle_繰り返し一グループもタイトル出力する_繰返し数2()
                Dim childrenRule As New CsvRule(Of ChildVo)(Function(defineBy As ICsvRuleLocator, vo As ChildVo) defineBy.Field(vo.Address, vo.Name))
                Dim broker As New CsvVoBroker(Of ParentVo)(Function(defineBy As ICsvRuleLocator, vo As ParentVo) defineBy.Field(vo.Id).GroupRepeat(vo.Children, childrenRule, 2), CsvSeparator.COMMA)
                broker.IsOutputTitle = True
                broker.Vos = New ParentVo() {}

                Assert.IsNotNull(broker.CsvStrings)
                Assert.AreEqual(1, broker.CsvStrings.Length, "データ0行だからタイトル行のみ")
                Assert.AreEqual("Id,Children(0).Address,Children(0).Name,Children(1).Address,Children(1).Name", broker.CsvStrings(0))
            End Sub

        End Class

        Public Class AddTitleOnlyTest : Inherits CsvVoBrokerTest

            Private Overloads Function ToString(vos As IEnumerable(Of SimpleVo)) As String
                Return (New DebugStringMaker(Of SimpleVo)(Function(define, vo) define.Bind(vo.Id, vo.Name))).MakeString(vos)
            End Function

            Private Overloads Function ToString(vos As IEnumerable(Of ParentVo)) As String
                Return (New DebugStringMaker(Of ParentVo)(Function(define, vo) define.Bind(vo.Id, vo.Child.Name, vo.Child.Address))).MakeString(vos)
            End Function

            <Test()>
            Public Sub AddTitleOnlyで列名のみの項目を設定できる_ToCsv()
                Dim broker As New CsvVoBroker(Of SimpleVo)(Function(define, vo)
                                                               define.Field(vo.Id).TitleOnly("からっぽ").Field(vo.Name)
                                                               Return define
                                                           End Function)
                broker.Vos = {New SimpleVo With {.Id = 1, .Name = "hoge"},
                              New SimpleVo With {.Id = 2, .Name = "fuga"}}

                Assert.That(broker.CsvStrings, [Is].EquivalentTo({"1,,hoge", "2,,fuga"}))
            End Sub

            <Test()>
            Public Sub AddTitleOnlyで列名のみの項目を設定できる_ToCsv_OutputTiltle()
                Dim broker As New CsvVoBroker(Of SimpleVo)(Function(define, vo)
                                                               define.Field(vo.Id).TitleOnly("からっぽ").Field(vo.Name)
                                                               Return define
                                                           End Function)
                broker.IsOutputTitle = True
                broker.Vos = {New SimpleVo With {.Id = 1, .Name = "hoge"},
                              New SimpleVo With {.Id = 2, .Name = "fuga"}}

                Assert.That(broker.CsvStrings, [Is].EquivalentTo(
                            {"Id,からっぽ,Name",
                             "1,,hoge",
                             "2,,fuga"}))
            End Sub

            <Test()>
            Public Sub AddTitleOnlyで列名のみの項目を設定できる_ToVo()
                Dim broker As New CsvVoBroker(Of SimpleVo)(Function(define, vo)
                                                               define.Field(vo.Id).TitleOnly("からっぽ").Field(vo.Name)
                                                               Return define
                                                           End Function)
                broker.CsvStrings = {"1,,hoge", "2,,fuga"}

                Assert.That(ToString(broker.Vos), [Is].EqualTo(
                            "Id Name  " & vbCrLf &
                            " 1 'hoge'" & vbCrLf &
                            " 2 'fuga'"))
            End Sub

            <Test()>
            Public Sub AddTitleOnly_グループに対して列名のみの項目を設定できる_ToCsv()
                Dim broker As New CsvVoBroker(Of ParentVo)(Function(define, vo)
                                                               define.Field(vo.Id)
                                                               define.TitleOnly("から")
                                                               define.Group(Of ChildVo)(vo.Child, Function(define1, vo1)
                                                                                                      define1.Field(vo1.Name)
                                                                                                      define1.TitleOnly("こども")
                                                                                                      define1.Field(vo1.Address)
                                                                                                      Return define1
                                                                                                  End Function)
                                                               Return define
                                                           End Function)

                broker.Vos = {New ParentVo With {.Id = 1, .Child = New ChildVo With {.Name = "hoge", .Address = "fuga"}},
                              New ParentVo With {.Id = 2, .Child = New ChildVo With {.Name = "piyo", .Address = "musa"}}}

                Assert.That(broker.CsvStrings, [Is].EquivalentTo(
                            {"1,,hoge,,fuga",
                             "2,,piyo,,musa"}))
            End Sub

            <Test()>
            Public Sub AddTitleOnly_グループに対して列名のみの項目を設定できる_ToCsv_OutputTiltle()
                Dim broker As New CsvVoBroker(Of ParentVo)(Function(define, vo)
                                                               define.Field(vo.Id)
                                                               define.TitleOnly("から")
                                                               define.Group(Of ChildVo)(vo.Child, Function(define1, vo1)
                                                                                                      define1.Field(vo1.Name)
                                                                                                      define1.TitleOnly("こども")
                                                                                                      define1.Field(vo1.Address)
                                                                                                      Return define1
                                                                                                  End Function)
                                                               Return define
                                                           End Function)

                broker.IsOutputTitle = True

                broker.Vos = {New ParentVo With {.Id = 1, .Child = New ChildVo With {.Name = "hoge", .Address = "fuga"}},
                              New ParentVo With {.Id = 2, .Child = New ChildVo With {.Name = "piyo", .Address = "musa"}}}

                Assert.That(broker.CsvStrings, [Is].EquivalentTo(
                            {"Id,から,Child.Name,こども,Child.Address",
                             "1,,hoge,,fuga",
                             "2,,piyo,,musa"}))
            End Sub

            <Test()>
            Public Sub AddTitleOnly_グループに対して列名のみの項目を設定できる_ToVo()
                Dim broker As New CsvVoBroker(Of ParentVo)(Function(define, vo)
                                                               define.Field(vo.Id)
                                                               define.TitleOnly("から")
                                                               define.Group(Of ChildVo)(vo.Child, Function(define1, vo1)
                                                                                                      define1.Field(vo1.Name)
                                                                                                      define1.TitleOnly("こども")
                                                                                                      define1.Field(vo1.Address)
                                                                                                      Return define1
                                                                                                  End Function)
                                                               Return define
                                                           End Function)

                broker.CsvStrings = {"1,,hoge,,fuga", "2,,piyo,,musa"}

                Assert.That(ToString(broker.Vos), [Is].EqualTo(
                            "Id Name   Address" & vbCrLf &
                            " 1 'hoge' 'fuga' " & vbCrLf &
                            " 2 'piyo' 'musa' "))
            End Sub

        End Class

        Public Class AddDecorationTest : Inherits CsvVoBrokerTest

            Private Overloads Function ToString(vos As IEnumerable(Of SimpleVo)) As String
                Return (New DebugStringMaker(Of SimpleVo)(Function(define, vo) define.Bind(vo.Id, vo.Name))).MakeString(vos)
            End Function

            <Test()>
            Public Sub CSV出力時に指定したPropertyに対して装飾を施す事ができる_ToCsv()
                Dim broker As New CsvVoBroker(Of SimpleVo)(Function(define, vo)
                                                               define.FieldWithDecorator(vo.Id,
                                                                                     toCsvDecorator:=Function(value) "id" & StringUtil.ToString(value),
                                                                                     toVoDecorator:=Function(str) StringUtil.ToInteger(str.Substring(2)))
                                                               define.Field(vo.Name)
                                                               Return define
                                                           End Function)
                broker.Vos = {New SimpleVo With {.Id = 1, .Name = "hoge"},
                              New SimpleVo With {.Id = 2, .Name = "fuga"}}

                Assert.That(broker.CsvStrings, [Is].EquivalentTo({"id1,hoge", "id2,fuga"}))
            End Sub

            <Test()>
            Public Sub CSV出力時に指定したPropertyに対して装飾を施す事ができる_ToVo()
                Dim broker As New CsvVoBroker(Of SimpleVo)(Function(define, vo)
                                                               define.FieldWithDecorator(vo.Id,
                                                                                     toCsvDecorator:=Function(value) "ddddd" & StringUtil.ToString(value),
                                                                                     toVoDecorator:=Function(str) StringUtil.ToInteger(str.Substring(5)))
                                                               define.Field(vo.Name)
                                                               Return define
                                                           End Function)
                broker.CsvStrings = {"ddddd1,hohoho",
                                     "abcde2,ffffff"}

                Assert.That(ToString(broker.Vos), [Is].EqualTo(
                            "Id Name    " & vbCrLf &
                            " 1 'hohoho'" & vbCrLf &
                            " 2 'ffffff'"))
            End Sub

            <Test()>
            Public Sub FieldWithDecoratorなのに装飾が何も指定してなければ_例外を返す()

                Try
                    Dim broker As New CsvVoBroker(Of SimpleVo)(Function(define, vo)
                                                                   define.FieldWithDecorator(vo.Id)
                                                                   define.Field(vo.Name)
                                                                   Return define
                                                               End Function)
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("WithDecoratorなのに装飾処理が何も指定されてない"))
                End Try
            End Sub

        End Class

    End Class
End Namespace