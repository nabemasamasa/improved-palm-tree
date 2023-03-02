Imports System.IO
Imports NUnit.Framework
Imports System.Xml

' 内部クラスのテストメソッドも実行させる JUnitの @RunWith(Enclosed.class) を行いたいけど、NUnitの方法がわからない
' わからないので、
' 1. ベースとなる XmlUtilTest を MustInherit
' 2. 内部クラスにしたかったクラスを XmlUtilTest と同列に配置し、XmlUtilTest を継承
' 3. 2のクラスにテストメソッドを追加
' 4. XmlUtilTest をテスト実行すると、継承先のテストがすべて実行される
Public MustInherit Class XmlUtilTest

#Region "Testing vo"
    Protected Class PrimitiveVo
        Private _no As Integer

        Public Property No() As Integer
            Get
                Return _no
            End Get
            Set(ByVal value As Integer)
                _no = value
            End Set
        End Property
    End Class
    Protected Class IdNameVo
        Private _id As Integer?
        Private _name As String

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
    End Class
    Protected Class IdNameArrayVo : Inherits IdNameVo
        Private _addresses As String()

        Public Property Addresses() As String()
            Get
                Return _addresses
            End Get
            Set(ByVal value As String())
                _addresses = value
            End Set
        End Property
    End Class
    Protected Class IdNameArrayIntVo : Inherits IdNameVo
        Private _empNos As Integer()

        Public Property EmpNos() As Integer()
            Get
                Return _empNos
            End Get
            Set(ByVal value As Integer())
                _empNos = value
            End Set
        End Property
    End Class
    Protected Class IdNameArrayIntNullableVo : Inherits IdNameVo
        Private _empNos As Integer?()

        Public Property EmpNos() As Integer?()
            Get
                Return _empNos
            End Get
            Set(ByVal value As Integer?())
                _empNos = value
            End Set
        End Property
    End Class
    Protected Class ChildVo : Inherits IdNameVo

    End Class
    Protected Class IdNameListChildVo : Inherits IdNameVo
        Private _children As List(Of ChildVo)

        Public Property Children() As List(Of ChildVo)
            Get
                Return _children
            End Get
            Set(ByVal value As List(Of ChildVo))
                _children = value
            End Set
        End Property
    End Class
    Protected Class IdNameArrayChildVo : Inherits IdNameVo
        Private _children As ChildVo()

        Public Property Children() As ChildVo()
            Get
                Return _children
            End Get
            Set(ByVal value As ChildVo())
                _children = value
            End Set
        End Property
    End Class
    Protected Class IdNameInChildVo : Inherits IdNameVo
        Private _child As ChildVo

        Public Property Child() As ChildVo
            Get
                Return _child
            End Get
            Set(ByVal value As ChildVo)
                _child = value
            End Set
        End Property
    End Class
    Protected Class ImmutableVo
        Private _id As Integer?
        Private _name As String
        Private _accessDateTime As DateTime?
        Private _value As Decimal?

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

        Public Property AccessDateTime() As Date?
            Get
                Return _accessDateTime
            End Get
            Set(ByVal value As Date?)
                _accessDateTime = value
            End Set
        End Property

        Public Property Value() As Decimal?
            Get
                Return _value
            End Get
            Set(ByVal value As Decimal?)
                _value = value
            End Set
        End Property
    End Class
#End Region

    Protected Const HEADER_STRINGWRITER As String = "<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>"
    Protected Const HEADER_UTF_8 As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
End Class

Public Class XmlUnitGeneralTest : Inherits XmlUtilTest

End Class

Public Class PopulateVoToDocTest : Inherits XmlUtilTest
    Private doc As XmlDocument
    Private writer As StringWriter

    <SetUp()> Public Sub SetUp()
        doc = New XmlDocument
        Dim declaration As XmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", "yes")
        doc.AppendChild(declaration)
        writer = New StringWriter
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のInt_有効な値ならそのまま出力()
        Dim vo As New PrimitiveVo With {.No = 2}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<PrimitiveVo>" & vbCrLf _
                        & "  <No>2</No>" & vbCrLf _
                        & "</PrimitiveVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のInt_Nothingなら既定値になるので0が出力される()
        Dim vo As New PrimitiveVo With {.No = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<PrimitiveVo>" & vbCrLf _
                        & "  <No>0</No>" & vbCrLf _
                        & "</PrimitiveVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のNullableInt及びString_有効な値ならそのまま出力()
        Dim vo As New IdNameVo With {.Id = 2, .Name = "n2"}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameVo>" & vbCrLf _
                        & "  <Id>2</Id>" & vbCrLf _
                        & "  <Name>n2</Name>" & vbCrLf _
                        & "</IdNameVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のNullableInt_Nothingならタグは出力されない()
        Dim vo As New IdNameVo With {.Id = Nothing, .Name = "n"}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameVo>" & vbCrLf _
                        & "  <Name>n</Name>" & vbCrLf _
                        & "</IdNameVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のNullableInt_NothingだけどFlagTrueなら空要素タグが出力される()
        Dim vo As New IdNameVo With {.Id = Nothing, .Name = "n"}
        XmlUtil.PopulateVoToDoc(vo, doc, True)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameVo>" & vbCrLf _
                        & "  <Id />" & vbCrLf _
                        & "  <Name>n</Name>" & vbCrLf _
                        & "</IdNameVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString_Nothingならタグは出力されない()
        Dim vo As New IdNameVo With {.Id = 3, .Name = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameVo>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "</IdNameVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString_NothingだけどFlagTrueなら空要素タグが出力される()
        Dim vo As New IdNameVo With {.Id = 3, .Name = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc, True)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameVo>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name />" & vbCrLf _
                        & "</IdNameVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString_空文字なら開始終了タグが出力される()
        Dim vo As New IdNameVo With {.Id = 4, .Name = ""}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameVo>" & vbCrLf _
                        & "  <Id>4</Id>" & vbCrLf _
                        & "  <Name>" & vbCrLf _
                        & "  </Name>" & vbCrLf _
                        & "</IdNameVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo_値があれば出力される()
        Dim vo As New IdNameInChildVo With {.Id = 4, .Child = New ChildVo With {.Name = "CN"}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameInChildVo>" & vbCrLf _
                        & "  <Child>" & vbCrLf _
                        & "    <Name>CN</Name>" & vbCrLf _
                        & "  </Child>" & vbCrLf _
                        & "  <Id>4</Id>" & vbCrLf _
                        & "</IdNameInChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo_Nothingならタグは出力されない()
        Dim vo As New IdNameInChildVo With {.Id = 3}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameInChildVo>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "</IdNameInChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo_NothingだけどFlagTrueなら空要素タグが出力される()
        Dim vo As New IdNameInChildVo With {.Id = 3}
        XmlUtil.PopulateVoToDoc(vo, doc, True)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameInChildVo>" & vbCrLf _
                        & "  <Child />" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name />" & vbCrLf _
                        & "</IdNameInChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo_空インスタンスなら開始終了タグが出力される()
        Dim vo As New IdNameInChildVo With {.Id = 4, .Child = New ChildVo}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameInChildVo>" & vbCrLf _
                        & "  <Child>" & vbCrLf _
                        & "  </Child>" & vbCrLf _
                        & "  <Id>4</Id>" & vbCrLf _
                        & "</IdNameInChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Voが空インスタンスなら開始終了タグが出力される()
        Dim vo As New IdNameVo With {.Id = Nothing, .Name = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameVo>" & vbCrLf _
                        & "</IdNameVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下にNullableInt及び不変オブジェクトのみ_ファイル出力ならちゃんとUTF8()
        FileUtil.DeleteFileIfExist("hoge.xml")

        Dim vo As New IdNameVo With {.Id = 2, .Name = "n2"}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save("hoge.xml")

        Try
            Assert.AreEqual(HEADER_UTF_8 & vbCrLf & "<IdNameVo>" & vbCrLf _
                            & "  <Id>2</Id>" & vbCrLf _
                            & "  <Name>n2</Name>" & vbCrLf _
                            & "</IdNameVo>", FileUtil.ReadFile("hoge.xml"))
        Finally
            FileUtil.DeleteFileIfExist("hoge.xml")
        End Try
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString配列()
        Dim vo As New IdNameArrayVo With {.Id = 3, .Name = "n3", .Addresses = New String() {"a1", "a2"}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayVo>" & vbCrLf _
                        & "  <Addresses>a1</Addresses>" & vbCrLf _
                        & "  <Addresses>a2</Addresses>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name>n3</Name>" & vbCrLf _
                        & "</IdNameArrayVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString配列_Nothingなら出力されない()
        Dim vo As New IdNameArrayVo With {.Id = 3, .Name = "n3", .Addresses = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayVo>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name>n3</Name>" & vbCrLf _
                        & "</IdNameArrayVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString配列_Nothingなら出力されない_FlagTrueでも同様()
        Dim vo As New IdNameArrayVo With {.Addresses = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc, True)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayVo>" & vbCrLf _
                        & "  <Id />" & vbCrLf _
                        & "  <Name />" & vbCrLf _
                        & "</IdNameArrayVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString配列_長さ0なら_Empty属性Trueで出力()
        Dim vo As New IdNameArrayVo With {.Addresses = New String() {}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayVo>" & vbCrLf _
                        & "  <Addresses Empty=""True"" />" & vbCrLf _
                        & "</IdNameArrayVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString配列_長さ1で中身Nothingなら_空要素タグで出力()
        Dim vo As New IdNameArrayVo With {.Addresses = New String() {Nothing}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayVo>" & vbCrLf _
                        & "  <Addresses />" & vbCrLf _
                        & "</IdNameArrayVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のString配列_長さ1で中身空文字なら_開始終了タグで出力()
        Dim vo As New IdNameArrayVo With {.Addresses = New String() {""}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayVo>" & vbCrLf _
                        & "  <Addresses>" & vbCrLf _
                        & "  </Addresses>" & vbCrLf _
                        & "</IdNameArrayVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のInt配列()
        Dim vo As New IdNameArrayIntVo With {.Id = 3, .Name = "n3", .EmpNos = New Integer() {3, 4, Nothing}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntVo>" & vbCrLf _
                        & "  <EmpNos>3</EmpNos>" & vbCrLf _
                        & "  <EmpNos>4</EmpNos>" & vbCrLf _
                        & "  <EmpNos>0</EmpNos>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name>n3</Name>" & vbCrLf _
                        & "</IdNameArrayIntVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のInt配列_Nothingなら出力されない()
        Dim vo As New IdNameArrayIntVo With {.Id = 3, .Name = "n3", .EmpNos = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntVo>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name>n3</Name>" & vbCrLf _
                        & "</IdNameArrayIntVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のInt配列_Nothingなら出力されない_FlagTrueでも同様()
        Dim vo As New IdNameArrayIntVo With {.EmpNos = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc, True)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntVo>" & vbCrLf _
                        & "  <Id />" & vbCrLf _
                        & "  <Name />" & vbCrLf _
                        & "</IdNameArrayIntVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のInt配列_長さ0なら_Empty属性Trueで出力()
        Dim vo As New IdNameArrayIntVo With {.EmpNos = New Integer() {}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntVo>" & vbCrLf _
                        & "  <EmpNos Empty=""True"" />" & vbCrLf _
                        & "</IdNameArrayIntVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のInt配列_長さ1で中身Nothingなら_値は既定値になるので値0で出力()
        Dim vo As New IdNameArrayIntVo With {.EmpNos = New Integer() {Nothing}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntVo>" & vbCrLf _
                        & "  <EmpNos>0</EmpNos>" & vbCrLf _
                        & "</IdNameArrayIntVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のNullableInt配列()
        Dim vo As New IdNameArrayIntNullableVo With {.Id = 3, .Name = "n3", .EmpNos = New Integer?() {3, 4}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntNullableVo>" & vbCrLf _
                        & "  <EmpNos>3</EmpNos>" & vbCrLf _
                        & "  <EmpNos>4</EmpNos>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name>n3</Name>" & vbCrLf _
                        & "</IdNameArrayIntNullableVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のNullableInt配列_Nothingなら出力されない()
        Dim vo As New IdNameArrayIntNullableVo With {.Id = 3, .Name = "n3", .EmpNos = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntNullableVo>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name>n3</Name>" & vbCrLf _
                        & "</IdNameArrayIntNullableVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のNullableInt配列_Nothingなら出力されない_FlagTrueでも同様()
        Dim vo As New IdNameArrayIntNullableVo With {.EmpNos = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc, True)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntNullableVo>" & vbCrLf _
                        & "  <Id />" & vbCrLf _
                        & "  <Name />" & vbCrLf _
                        & "</IdNameArrayIntNullableVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のNullableInt配列_長さ0なら_Empty属性Trueで出力()
        Dim vo As New IdNameArrayIntNullableVo With {.EmpNos = New Integer?() {}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntNullableVo>" & vbCrLf _
                        & "  <EmpNos Empty=""True"" />" & vbCrLf _
                        & "</IdNameArrayIntNullableVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のNullableInt配列_長さ1で中身Nothingなら_空要素タグで出力()
        Dim vo As New IdNameArrayIntNullableVo With {.EmpNos = New Integer?() {Nothing}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayIntNullableVo>" & vbCrLf _
                        & "  <EmpNos />" & vbCrLf _
                        & "</IdNameArrayIntNullableVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo配列()
        Dim c1 As New ChildVo With {.Id = 13}
        Dim c2 As New ChildVo With {.name = "Apr"}
        Dim vo As New IdNameArrayChildVo With {.Id = 3, .Name = "n3", .Children = New ChildVo() {c1, c2}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayChildVo>" & vbCrLf _
                        & "  <Children>" & vbCrLf _
                        & "    <Id>13</Id>" & vbCrLf _
                        & "  </Children>" & vbCrLf _
                        & "  <Children>" & vbCrLf _
                        & "    <Name>Apr</Name>" & vbCrLf _
                        & "  </Children>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name>n3</Name>" & vbCrLf _
                        & "</IdNameArrayChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo配列_Nothingなら出力されない()
        Dim vo As New IdNameArrayChildVo With {.Id = 3, .Name = "n3", .Children = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayChildVo>" & vbCrLf _
                        & "  <Id>3</Id>" & vbCrLf _
                        & "  <Name>n3</Name>" & vbCrLf _
                        & "</IdNameArrayChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo配列_Nothingなら出力されない_FlagTrueでも同様()
        Dim vo As New IdNameArrayChildVo With {.Children = Nothing}
        XmlUtil.PopulateVoToDoc(vo, doc, True)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayChildVo>" & vbCrLf _
                        & "  <Id />" & vbCrLf _
                        & "  <Name />" & vbCrLf _
                        & "</IdNameArrayChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo配列_長さ0なら_Empty属性Trueで出力()
        Dim vo As New IdNameArrayChildVo With {.Children = New ChildVo() {}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayChildVo>" & vbCrLf _
                        & "  <Children Empty=""True"" />" & vbCrLf _
                        & "</IdNameArrayChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo直下のVo配列_長さ1で中身Nothingなら_空要素タグで出力()
        Dim vo As New IdNameArrayChildVo With {.Children = New ChildVo() {Nothing}}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameArrayChildVo>" & vbCrLf _
                        & "  <Children />" & vbCrLf _
                        & "</IdNameArrayChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_Vo型Listを持つVo()
        Dim c1 As New ChildVo With {.Id = 11, .Name = "n1"}
        Dim c2 As New ChildVo With {.Id = 12, .Name = "n2"}
        Dim vo As New IdNameListChildVo With {.Id = 4, .Name = "n4", .Children = EzUtil.NewList(c1, c2)}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<IdNameListChildVo>" & vbCrLf _
                        & "  <Children>" & vbCrLf _
                        & "    <Id>11</Id>" & vbCrLf _
                        & "    <Name>n1</Name>" & vbCrLf _
                        & "  </Children>" & vbCrLf _
                        & "  <Children>" & vbCrLf _
                        & "    <Id>12</Id>" & vbCrLf _
                        & "    <Name>n2</Name>" & vbCrLf _
                        & "  </Children>" & vbCrLf _
                        & "  <Id>4</Id>" & vbCrLf _
                        & "  <Name>n4</Name>" & vbCrLf _
                        & "</IdNameListChildVo>", writer.ToString)
    End Sub

    <Test()> Public Sub PopulateVoToNode_日付型はミリ秒まで出力()
        Dim vo As New ImmutableVo With {.AccessDateTime = CDate("2012/03/04 05:06:07.123")}
        XmlUtil.PopulateVoToDoc(vo, doc)

        doc.Save(writer)

        Assert.AreEqual(HEADER_STRINGWRITER & vbCrLf & "<ImmutableVo>" & vbCrLf _
                        & "  <AccessDateTime>2012/03/04 05:06:07.123</AccessDateTime>" & vbCrLf _
                        & "</ImmutableVo>", writer.ToString)
    End Sub
End Class

Public Class PopulateNodeToVoTest : Inherits XmlUtilTest
    Private doc As XmlDocument

    <SetUp()> Public Sub SetUp()
        doc = New XmlDocument
        Dim declaration As XmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", "yes")
        doc.AppendChild(declaration)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt_有効な値ならそのまま取得()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <No>2</No>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New PrimitiveVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.AreEqual(2, vo.No)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt_要素が無ければ_既定値0のまま()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New PrimitiveVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.AreEqual(0, vo.No)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt_空要素タグなら_既定値0のまま()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <No />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New PrimitiveVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.AreEqual(0, vo.No)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt_開始終了タグなら_既定値0のまま()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <No></No>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New PrimitiveVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.AreEqual(0, vo.No)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt_不正な値なら_既定値0のまま()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <No>あいう</No>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New PrimitiveVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.AreEqual(0, vo.No)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt_有効な値ならそのまま取得()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Id>2</Id>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.AreEqual(2, vo.Id)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt_要素が無ければ_Nothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Id)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt_空要素タグなら_Nothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Id />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Id)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt_開始終了タグなら_Nothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Id></Id>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Id)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt_不正な値なら_Nothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Id>あいうえ</Id>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Id)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString_有効な値ならそのまま取得()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Name>あいう</Name>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.AreEqual("あいう", vo.Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString_要素が無ければ_Nothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString_空要素タグなら_Nothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Name />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString_開始終了タグなら_空文字()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Name></Name>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.AreEqual("", vo.Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo_有効な値ならそのまま取得()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Child>" & vbCrLf _
                        & "    <Id>9</Id>" & vbCrLf _
                        & "    <Name>Sep</Name>" & vbCrLf _
                        & "  </Child>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameInChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Child)
        Assert.AreEqual(9, vo.Child.Id)
        Assert.AreEqual("Sep", vo.Child.Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo_要素が無ければ_Nothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameInChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Child)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo_空要素タグなら_Nothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Child />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameInChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Child)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo_開始終了タグなら_空インスタンス()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Child></Child>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameInChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Child)
        Assert.IsNull(vo.Child.Id)
        Assert.IsNull(vo.Child.Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo_不正な値なら_空インスタンス()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Child>あいうえ</Child>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameInChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Child)
        Assert.IsNull(vo.Child.Id)
        Assert.IsNull(vo.Child.Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt配列_有効な値ならそのまま取得()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos>2</EmpNos>" & vbCrLf _
                        & "  <EmpNos>3</EmpNos>" & vbCrLf _
                        & "  <EmpNos>4</EmpNos>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(3, vo.EmpNos.Length)
        Assert.AreEqual(2, vo.EmpNos(0))
        Assert.AreEqual(3, vo.EmpNos(1))
        Assert.AreEqual(4, vo.EmpNos(2))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt配列_要素が無ければ_配列がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.EmpNos)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt配列_空要素タグなら_長さ1で値が既定値()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(1, vo.EmpNos.Length)
        Assert.AreEqual(0, vo.EmpNos(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt配列_開始終了タグなら_長さ1で値が既定値()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos></EmpNos>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(1, vo.EmpNos.Length)
        Assert.AreEqual(0, vo.EmpNos(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt配列_不正な値なら_長さ1で値が既定値()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos>あいう</EmpNos>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(1, vo.EmpNos.Length)
        Assert.AreEqual(0, vo.EmpNos(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のInt配列_空要素でEmpty属性Trueなら_長さ0の配列()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos Empty='True' />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(0, vo.EmpNos.Length)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt配列_有効な値ならそのまま取得()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos>2</EmpNos>" & vbCrLf _
                        & "  <EmpNos>3</EmpNos>" & vbCrLf _
                        & "  <EmpNos>4</EmpNos>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntNullableVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(3, vo.EmpNos.Length)
        Assert.AreEqual(2, vo.EmpNos(0))
        Assert.AreEqual(3, vo.EmpNos(1))
        Assert.AreEqual(4, vo.EmpNos(2))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt配列_要素が無ければ_配列がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntNullableVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.EmpNos)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt配列_空要素タグなら_長さ1で値がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntNullableVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(1, vo.EmpNos.Length)
        Assert.IsNull(vo.EmpNos(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt配列_開始終了タグなら_長さ1で値がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos></EmpNos>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntNullableVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(1, vo.EmpNos.Length)
        Assert.IsNull(vo.EmpNos(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt配列_不正な値なら_長さ1で値がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos>あいう</EmpNos>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntNullableVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(1, vo.EmpNos.Length)
        Assert.IsNull(vo.EmpNos(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のNullableInt配列_空要素でEmpty属性Trueなら_長さ0の配列()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <EmpNos Empty='True'></EmpNos>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayIntNullableVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.EmpNos)
        Assert.AreEqual(0, vo.EmpNos.Length)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString配列_有効な値ならそのまま取得()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Addresses>aa</Addresses>" & vbCrLf _
                        & "  <Addresses />" & vbCrLf _
                        & "  <Addresses></Addresses>" & vbCrLf _
                        & "  <Addresses>dd</Addresses>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Addresses)
        Assert.AreEqual(4, vo.Addresses.Length)
        Assert.AreEqual("aa", vo.Addresses(0))
        Assert.IsNull(vo.Addresses(1))
        Assert.AreEqual("", vo.Addresses(2))
        Assert.AreEqual("dd", vo.Addresses(3))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString配列_要素が無ければ_配列がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Addresses)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString配列_空要素タグなら_長さ1で値がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Addresses />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Addresses)
        Assert.AreEqual(1, vo.Addresses.Length)
        Assert.IsNull(vo.Addresses(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString配列_開始終了タグなら_長さ1で値が空文字()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Addresses></Addresses>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Addresses)
        Assert.AreEqual(1, vo.Addresses.Length)
        Assert.AreEqual("", vo.Addresses(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のString配列_空要素でEmpty属性Trueなら_長さ0の配列()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Addresses Empty='True'></Addresses>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Addresses)
        Assert.AreEqual(0, vo.Addresses.Length)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo配列_有効な値ならそのまま取得()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Children>" & vbCrLf _
                        & "    <Id>2</Id>" & vbCrLf _
                        & "  </Children>" & vbCrLf _
                        & "  <Children>" & vbCrLf _
                        & "    <Name>Feb</Name>" & vbCrLf _
                        & "  </Children>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Children)
        Assert.AreEqual(2, vo.Children.Length)
        Assert.AreEqual(2, vo.Children(0).Id)
        Assert.AreEqual("Feb", vo.Children(1).Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo配列_要素が無ければ_配列がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNull(vo.Children)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo配列_空要素タグなら_長さ1で値がNothing()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Children />" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Children)
        Assert.AreEqual(1, vo.Children.Length)
        Assert.IsNull(vo.Children(0))
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo配列_開始終了タグなら_長さ1で値が空インスタンス()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Children></Children>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Children)
        Assert.AreEqual(1, vo.Children.Length)
        Assert.IsNotNull(vo.Children(0))
        Assert.IsNull(vo.Children(0).Id)
        Assert.IsNull(vo.Children(0).Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo配列_不正な値なら_長さ1で値が空インスタンス()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Children>あいう</Children>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Children)
        Assert.AreEqual(1, vo.Children.Length)
        Assert.IsNotNull(vo.Children(0))
        Assert.IsNull(vo.Children(0).Id)
        Assert.IsNull(vo.Children(0).Name)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo直下のVo配列_空要素でEmpty属性Trueなら_長さ0の配列()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<dummy>" & vbCrLf _
                        & "  <Children Empty='True'></Children>" & vbCrLf _
                        & "</dummy>")
        Dim vo As New IdNameArrayChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Children)
        Assert.AreEqual(0, vo.Children.Length)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_Vo型Listを持つVo()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<Dummy>" & vbCrLf _
                        & "  <Children>" & vbCrLf _
                        & "    <Name>n1</Name>" & vbCrLf _
                        & "  </Children>" & vbCrLf _
                        & "  <Children />" & vbCrLf _
                        & "  <Children></Children>" & vbCrLf _
                        & "  <Children>" & vbCrLf _
                        & "    <Id>13</Id>" & vbCrLf _
                        & "  </Children>" & vbCrLf _
                        & "</Dummy>")
        Dim vo As New IdNameListChildVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.Children)
        Assert.AreEqual(4, vo.Children.Count)
        Assert.IsNotNull(vo.Children(0))
        Assert.AreEqual("n1", vo.Children(0).Name)
        Assert.IsNull(vo.Children(1))
        Assert.IsNotNull(vo.Children(2))
        Assert.IsNull(vo.Children(2).Id)
        Assert.IsNull(vo.Children(2).Name)
        Assert.IsNotNull(vo.Children(3))
        Assert.AreEqual(13, vo.Children(3).Id)
    End Sub

    <Test()> Public Sub PopulateNodeToVo_日付型_ミリ秒もパース()
        doc.LoadXml(HEADER_STRINGWRITER & vbCrLf & "<Dummy>" & vbCrLf _
                        & "  <AccessDateTime>2011/12/01 02:03:04.567</AccessDateTime>" & vbCrLf _
                        & "</Dummy>")
        Dim vo As New ImmutableVo
        XmlUtil.PopulateNodeToVo(doc.LastChild, vo)

        Assert.IsNotNull(vo.AccessDateTime)
        Assert.AreEqual(CDate("2011/12/01 02:03:04.567"), vo.AccessDateTime.Value)
    End Sub

End Class
