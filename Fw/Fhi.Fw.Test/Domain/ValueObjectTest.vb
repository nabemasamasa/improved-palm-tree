Imports NUnit.Framework

Namespace Domain
    Public MustInherit Class ValueObjectTest

#Region "Nested classes.."
        ' 値オブジェクト実装例
        Public Class TestingValueObject : Inherits ValueObject

            Public ReadOnly Name As String
            Public ReadOnly ZipCode As String
            Public ReadOnly Address As String

            Public Sub New(name As String, zipCode As String, address As String)
                Me.Name = name
                Me.ZipCode = zipCode
                Me.Address = address
            End Sub

            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                ' すべての値をreturnすること
                ' ※MSDNでは Yield Return してるけどVB2010にはないので...
                Return New Object() {Name, ZipCode, Address}
            End Function

        End Class
        Private Class NestedValueObject : Inherits ValueObject
            Public ReadOnly StaffName As String
            Public ReadOnly Customer As TestingValueObject
            Public Sub New(staffName As String, customer As TestingValueObject)
                Me.StaffName = staffName
                Me.Customer = customer
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Me.StaffName, Me.Customer}
            End Function
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()

        End Sub

        Public Class DefaultTest : Inherits ValueObjectTest

            <TestCase("n1", "z1", "a1")>
            <TestCase(Nothing, "z1", "a1")>
            <TestCase("n1", Nothing, "a1")>
            <TestCase("n1", "z1", Nothing)>
            <TestCase("n1", Nothing, Nothing)>
            <TestCase(Nothing, "z1", Nothing)>
            <TestCase(Nothing, Nothing, "a1")>
            <TestCase("", "", "")>
            <TestCase("n1", "", "")>
            <TestCase("", "z1", "")>
            <TestCase("", "", "a1")>
            Public Sub Equalsメソッドで_同値比較できる(name As String, zipCode As String, address As String)
                Dim item1 As New TestingValueObject(name, zipCode, address)
                Dim item2 As New TestingValueObject(name, zipCode, address)
                Assert.That(item1.Equals(item2), [Is].True)
                Assert.That(item2.Equals(item1), [Is].True)
            End Sub

            <TestCase("n1", "z1", "a1")>
            <TestCase(Nothing, "z1", "a1")>
            <TestCase("n1", Nothing, "a1")>
            <TestCase("n1", "z1", Nothing)>
            Public Sub 値が違えば_Equalsメソッドは_偽になる(name As String, zipCode As String, address As String)
                Dim item1 As New TestingValueObject(name, If(zipCode, "") & "zipCode", address)
                Dim item2 As New TestingValueObject(name, zipCode, address)
                Assert.That(item1.Equals(item2), [Is].False)
                Assert.That(item2.Equals(item1), [Is].False)
            End Sub

            <TestCase("s1", "n1", "z1", "a1")>
            <TestCase(Nothing, "n1", "z1", "a1")>
            <TestCase("s1", Nothing, "z1", "a1")>
            <TestCase("s1", "n1", Nothing, "a1")>
            <TestCase("s1", "n1", "z1", Nothing)>
            Public Sub Equalsメソッドで_同値比較できる_ValueObjectの入れ子でもok(staffName As String, name As String, zipCode As String, address As String)
                Dim item1 As New NestedValueObject(staffName, New TestingValueObject(name, zipCode, address))
                Dim item2 As New NestedValueObject(staffName, New TestingValueObject(name, zipCode, address))
                Assert.That(item1.Equals(item2), [Is].True)
                Assert.That(item2.Equals(item1), [Is].True)
            End Sub

            <Test()> Public Sub Equalsメソッドで_同値比較できる_ValueObjectの入れ子がnullでもok()
                Dim item1 As New NestedValueObject("s1", Nothing)
                Dim item2 As New NestedValueObject("s1", Nothing)
                Assert.That(item1.Equals(item2), [Is].True)
                Assert.That(item2.Equals(item1), [Is].True)
            End Sub

            <TestCase("s1", "n1", "z1", "a1")>
            <TestCase(Nothing, "n1", "z1", "a1")>
            <TestCase("s1", Nothing, "z1", "a1")>
            <TestCase("s1", "n1", Nothing, "a1")>
            <TestCase("s1", "n1", "z1", Nothing)>
            Public Sub 値が違えば_Equalsメソッドは_偽になる_ValueObjectの入れ子でもok(staffName As String, name As String, zipCode As String, address As String)
                Dim item1 As New NestedValueObject(staffName, New TestingValueObject(name, If(zipCode, "") & "*", address))
                Dim item2 As New NestedValueObject(staffName, New TestingValueObject(name, zipCode, address))
                Assert.That(item1.Equals(item2), [Is].False)
                Assert.That(item2.Equals(item1), [Is].False)
            End Sub

            <Test()> Public Sub インスタンスが別でも_同値なら_Dictionaryの値を取得できる()
                Dim key1 As New TestingValueObject("a", "b", "c")
                Dim key2 As New TestingValueObject("a", "b", "c")
                Dim valueBy As New Dictionary(Of TestingValueObject, String)
                valueBy.Add(key1, "xyz")

                Assert.That(valueBy.ContainsKey(key2), [Is].True)
                Assert.That(valueBy(key2), [Is].EqualTo("xyz"), "インスタンス別でも引き出せる")
            End Sub

        End Class

        Public Class XorだとHashCodeが重複するので回避Test : Inherits ValueObjectTest

            Private Class StrPVO : Inherits PrimitiveValueObject(Of String)
                Public Sub New(ByVal value As String)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class BytePVO : Inherits PrimitiveValueObject(Of Byte?)
                Public Sub New(ByVal value As Byte?)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class TestingVO : Inherits ValueObject
                Public ReadOnly Str As StrPVO
                Public ReadOnly Byt As BytePVO
                Public Sub New(str As String, byt As Byte)
                    Me.Str = New StrPVO(str)
                    Me.Byt = New BytePVO(byt)
                End Sub
                Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                    Return New Object() {Str, Byt}
                End Function
            End Class
            Private Class TestingVO2 : Inherits ValueObject
                Public ReadOnly Vo As TestingVO
                Public ReadOnly Str2 As StrPVO
                Public Sub New(vo As TestingVO, str As StrPVO)
                    Me.Vo = vo
                    Me.Str2 = str
                End Sub
                Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                    Return New Object() {Vo, Str2}
                End Function
            End Class

            <Test()> Public Sub _32bitOSで_HashCodeが重複する問題を解決する()
                ' 32bitOSで Xor演算だと ("13212".GetHashCode Xor 1) = ("13211".GetHashCode Xor 0) が成り立つので回避
                Dim a As New TestingVO("13212", 1)
                Dim b As New TestingVO("13211", 0)
                Assert.That(a.GetHashCode, [Is].Not.EqualTo(b.GetHashCode))
            End Sub

            <Test()> Public Sub ValueObjectが_入れ子のキー値だと_HashCodeが重複する問題を解決する()
                Dim a As New TestingVO2(New TestingVO("13212", 1), New StrPVO("10"))
                Dim b As New TestingVO2(New TestingVO("13211", 0), New StrPVO("10"))
                Assert.That(a.GetHashCode, [Is].Not.EqualTo(b.GetHashCode))
            End Sub

            <Test()> Public Sub 内部値が違うなら_Hashも異なる_Byte()
                Dim a As New BytePVO(10)
                Dim b As New BytePVO(20)
                Assert.That(a.GetHashCode, [Is].Not.EqualTo(b.GetHashCode))
            End Sub

            <Test()> Public Sub 内部値が違うなら_Hashも異なる_String()
                Dim a As New StrPVO("10")
                Dim b As New StrPVO("20")
                Assert.That(a.GetHashCode, [Is].Not.EqualTo(b.GetHashCode))
            End Sub

        End Class

    End Class
End Namespace
