Imports NUnit.Framework

Namespace TestUtil.Test
    Public MustInherit Class NonPublicUtilTest

#Region "Testing Nested classes..."
        Private Class Holder
            Private Class TruePrivateVo
            End Class
            Private Class PrivateVoPrivateCtor
                Private Sub New()
                    Me.New(Nothing)
                End Sub
                Private Sub New(ByVal value As String)
                    _value = value
                End Sub

                Private _value As String
                Public Property Value() As String
                    Get
                        Return _value
                    End Get
                    Set(ByVal value As String)
                        _value = value
                    End Set
                End Property
            End Class
            Private Class PrivateVoPublicCtor
                Public Sub New()
                    Me.New(Nothing)
                End Sub
                Public Sub New(ByVal value As String)
                    _value = value
                End Sub

                Private _value As String
                Public Property Value() As String
                    Get
                        Return _value
                    End Get
                    Set(ByVal value As String)
                        _value = value
                    End Set
                End Property
            End Class
            Private Class PrivateVoField
                Public Shared Sub Initialize()
                    SharedFieldOfPublic = Nothing
                    SharedFieldOfPrivate = Nothing
                End Sub
                Public Shared SharedFieldOfPublic As String
                Private Shared SharedFieldOfPrivate As String
                Public Shared Function GetSharedFieldOfPrivate() As String
                    Return SharedFieldOfPrivate
                End Function
                Public FieldOfPublic As String
                Private fieldOfPrivate As String
                Public Function GetFieldOfPrivate() As String
                    Return fieldOfPrivate
                End Function
            End Class
            Private Class PrivateVoProperty
                Public Shared Sub Initialize()
                    _sharedValueOfPrivate = Nothing
                    _sharedValueOfPrivateSet = Nothing
                    _sharedValueOfPublic = Nothing
                End Sub

                Private _valueOfPublic As String
                Public Property ValueOfPublic() As String
                    Get
                        Return _valueOfPublic
                    End Get
                    Set(ByVal value As String)
                        _valueOfPublic = value
                    End Set
                End Property

                Private _valueOfPrivate As String
                Private Property ValueOfPrivate() As String
                    Get
                        Return _valueOfPrivate
                    End Get
                    Set(ByVal value As String)
                        _valueOfPrivate = value
                    End Set
                End Property

                Public Function GetValueOfPrivate() As String
                    Return _valueOfPrivate
                End Function

                Private Shared _sharedValueOfPrivate As String
                Private Shared Property SharedValueOfPrivate() As String
                    Get
                        Return _sharedValueOfPrivate
                    End Get
                    Set(ByVal value As String)
                        _sharedValueOfPrivate = value
                    End Set
                End Property

                Public Shared Function GetSharedValueOfPrivate() As String
                    Return _sharedValueOfPrivate
                End Function

                Private _valueOfPrivateSet As String
                Public Property ValueOfPrivateSet() As String
                    Get
                        Return _valueOfPrivateSet
                    End Get
                    Private Set(ByVal value As String)
                        _valueOfPrivateSet = value
                    End Set
                End Property

                Private Shared _sharedValueOfPrivateSet As String
                Public Shared Property SharedValueOfPrivateSet() As String
                    Get
                        Return _sharedValueOfPrivateSet
                    End Get
                    Private Set(ByVal value As String)
                        _sharedValueOfPrivateSet = value
                    End Set
                End Property

                Private Shared _sharedValueOfPublic As String
                Public Shared Property SharedValueOfPublic() As String
                    Get
                        Return _sharedValueOfPublic
                    End Get
                    Private Set(ByVal value As String)
                        _sharedValueOfPublic = value
                    End Set
                End Property
            End Class
            Private Class PrivateMethod
                Private ReadOnly _value As String
                Public Sub New()
                    Me.New(Nothing)
                End Sub
                Public Sub New(ByVal value As String)
                    _value = value
                End Sub
                Private Function CombinePrivate(ByVal value As String) As String
                    Return "-" & _value & value
                End Function
                Public Function CombinePublic(ByVal value As String) As String
                    Return "+" & _value & value
                End Function
                Private Shared Function CombineSharedPrivate(ByVal value As String) As String
                    Return "-(s)" & value
                End Function
                Public Shared Function CombineSharedPublic(ByVal value As String) As String
                    Return "+(s)" & value
                End Function
            End Class
            Private Class SubPrivateVoField : Inherits PrivateVoField
                Private fieldOfPrivate2 As Object
            End Class
            Private Class SubPrivateVoProperty : Inherits PrivateVoProperty
                Private _valueOfProperty2 As Object
                Private Property ValueOfProperty2() As Object
                    Get
                        Return _valueOfProperty2
                    End Get
                    Set(ByVal value As Object)
                        _valueOfProperty2 = value
                    End Set
                End Property
            End Class
            Public Shared Sub Initialize()
                PrivateVoField.Initialize()
                PrivateVoProperty.Initialize()
            End Sub
        End Class
        Private Class PrivateVo
        End Class
        Public Class PublicVoPrivateCtor
            Private Sub New()
                Me.New(Nothing)
            End Sub
            Private Sub New(ByVal value As String)
                _value = value
            End Sub

            Private _value As String
            Public Property Value() As String
                Get
                    Return _value
                End Get
                Set(ByVal value As String)
                    _value = value
                End Set
            End Property
        End Class
        Public Class PublicVoPublicCtor
            Public Sub New()
                Me.New(Nothing)
            End Sub
            Public Sub New(ByVal value As String)
                _value = value
            End Sub

            Private _value As String
            Public Property Value() As String
                Get
                    Return _value
                End Get
                Set(ByVal value As String)
                    _value = value
                End Set
            End Property
        End Class
        Public Class PublicVoField
            Public Shared SharedFieldOfPublic As String
            Private Shared SharedFieldOfPrivate As String
            Public Shared Function GetSharedFieldOfPrivate() As String
                Return SharedFieldOfPrivate
            End Function
            Public FieldOfPublic As String
            Private fieldOfPrivate As String
            Public Function GetFieldOfPrivate() As String
                Return fieldOfPrivate
            End Function
            Public Shared Sub Initialize()
                SharedFieldOfPublic = Nothing
                SharedFieldOfPrivate = Nothing
            End Sub
        End Class
        Public Class PublicVoProperty
            Public Shared Sub Initialize()
                _sharedValueOfPrivate = Nothing
                _sharedValueOfPrivateSet = Nothing
                _sharedValueOfPublic = Nothing
            End Sub

            Private _valueOfPublic As String
            Public Property ValueOfPublic() As String
                Get
                    Return _valueOfPublic
                End Get
                Set(ByVal value As String)
                    _valueOfPublic = value
                End Set
            End Property

            Private _valueOfPrivate As String
            Private Property ValueOfPrivate() As String
                Get
                    Return _valueOfPrivate
                End Get
                Set(ByVal value As String)
                    _valueOfPrivate = value
                End Set
            End Property

            Public Function GetValueOfPrivate() As String
                Return _valueOfPrivate
            End Function

            Private Shared _sharedValueOfPrivate As String
            Private Shared Property SharedValueOfPrivate() As String
                Get
                    Return _sharedValueOfPrivate
                End Get
                Set(ByVal value As String)
                    _sharedValueOfPrivate = value
                End Set
            End Property

            Public Shared Function GetSharedValueOfPrivate() As String
                Return _sharedValueOfPrivate
            End Function

            Private _valueOfPrivateSet As String
            Public Property ValueOfPrivateSet() As String
                Get
                    Return _valueOfPrivateSet
                End Get
                Private Set(ByVal value As String)
                    _valueOfPrivateSet = value
                End Set
            End Property

            Private Shared _sharedValueOfPrivateSet As String
            Public Shared Property SharedValueOfPrivateSet() As String
                Get
                    Return _sharedValueOfPrivateSet
                End Get
                Private Set(ByVal value As String)
                    _sharedValueOfPrivateSet = value
                End Set
            End Property

            Private Shared _sharedValueOfPublic As String
            Public Shared Property SharedValueOfPublic() As String
                Get
                    Return _sharedValueOfPublic
                End Get
                Set(ByVal value As String)
                    _sharedValueOfPublic = value
                End Set
            End Property
        End Class
        Public Class PublicMethod
            Private ReadOnly _value As String
            Public Sub New()
                Me.New(Nothing)
            End Sub
            Public Sub New(ByVal value As String)
                _value = value
            End Sub
            Private Function CombinePrivate(ByVal value As String) As String
                Return "-" & _value & value
            End Function
            Public Function CombinePublic(ByVal value As String) As String
                Return "+" & _value & value
            End Function
            Private Shared Function CombineSharedPrivate(ByVal value As String) As String
                Return "-(s)" & value
            End Function
            Public Shared Function CombineSharedPublic(ByVal value As String) As String
                Return "+(s)" & value
            End Function
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()
            Holder.Initialize()
            PublicVoField.Initialize()
            PublicVoProperty.Initialize()
        End Sub

        Public Class GetTypeOfNonPublicTest : Inherits NonPublicUtilTest

            <Test()> Public Sub 非公開のクラス情報を取得できる_内部クラスはプラス記号で結合する()
                Dim actualType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+TruePrivateVo")

                Assert.That(actualType, [Is].Not.Null)
                Assert.That(actualType.FullName, [Is].EqualTo("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+TruePrivateVo"))
            End Sub

            <Test()> Public Sub 非公開のクラス情報を取得できる_内部クラスはプラス記号で結合する2()
                Dim actualType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PrivateVo")

                Assert.That(actualType, [Is].Not.Null)
                Assert.That(actualType.FullName, [Is].EqualTo("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PrivateVo"))
            End Sub

        End Class

        Public Class NewInstanceTest : Inherits NonPublicUtilTest

            <Test()> Public Sub PublicクラスのNonPublicコンストラクタでインスタンスを生成できる_引数ナシ()
                Dim actual As PublicVoPrivateCtor = NonPublicUtil.NewInstance(Of PublicVoPrivateCtor)()

                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.GetType, [Is].EqualTo(GetType(PublicVoPrivateCtor)))
            End Sub

            <Test()> Public Sub PublicクラスのNonPublicコンストラクタでインスタンスを生成できる_引数アリ()
                Dim actual As PublicVoPrivateCtor = NonPublicUtil.NewInstance(Of PublicVoPrivateCtor)("aiueo")

                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.GetType, [Is].EqualTo(GetType(PublicVoPrivateCtor)))
                Assert.That(actual.Value, [Is].EqualTo("aiueo"), "引数付きのコンストラクタで設定された値")
            End Sub

            <Test()> Public Sub NonPublicクラスのNonPublicコンストラクタでインスタンスを生成できる_引数ナシ()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoPrivateCtor")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.GetType, [Is].EqualTo(privateType))
            End Sub

            <Test()> Public Sub NonPublicクラスのNonPublicコンストラクタでインスタンスを生成できる_引数アリ()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoPrivateCtor")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType, "12345")

                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.GetType, [Is].EqualTo(privateType))
            End Sub

            <Test()> Public Sub NonPublicクラスのPublicコンストラクタでインスタンスを生成できる_引数ナシ()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoPublicCtor")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.GetType, [Is].EqualTo(privateType))
            End Sub

            <Test()> Public Sub NonPublicクラスのPublicコンストラクタでインスタンスを生成できる_引数アリ()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoPublicCtor")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType, "12345")

                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.GetType, [Is].EqualTo(privateType))
            End Sub

            <Test()> Public Sub PublicクラスのPublicコンストラクタでインスタンス生成は例外になる()
                Try
                    NonPublicUtil.NewInstance(GetType(PublicVoPublicCtor))
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoPublicCtor' でPublic .ctor() なら正規にインスタンス生成すべき"))
                End Try
            End Sub

            <Test()> Public Sub 存在しないコンストラクタは例外になる()
                Try
                    NonPublicUtil.NewInstance(GetType(PublicVoPublicCtor), 123)
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoPublicCtor に.ctor(System.Int32) は無い"))
                End Try
            End Sub

        End Class

        Public Class SetFieldTest : Inherits NonPublicUtilTest

            <Test()> Public Sub PrivateクラスのPrivateFieldに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoField")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                NonPublicUtil.SetField(actual, "fieldOfPrivate", "abc")

                Assert.That(NonPublicUtil.GetField(actual, "fieldOfPrivate"), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub PrivateクラスのPublicFieldに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoField")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                NonPublicUtil.SetField(actual, "FieldOfPublic", "bbc")

                Assert.That(NonPublicUtil.GetField(actual, "FieldOfPublic"), [Is].EqualTo("bbc"))
            End Sub

            <Test()> Public Sub Privateクラスの親クラスのPrivateFieldに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+SubPrivateVoField")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                NonPublicUtil.SetField(actual, "fieldOfPrivate", "abc")

                Assert.That(NonPublicUtil.GetField(actual, "fieldOfPrivate"), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub PublicクラスのPrivateFieldに値を設定できる()
                Dim vo As New PublicVoField

                NonPublicUtil.SetField(vo, "fieldOfPrivate", "ccc")

                Assert.That(vo.GetFieldOfPrivate, [Is].EqualTo("ccc"))
            End Sub

            <Test()> Public Sub PublicクラスのPublicFieldに値設定は例外になる()
                Dim vo As New PublicVoField
                Try
                    NonPublicUtil.SetField(vo, "FieldOfPublic", "ddd")
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField' でPublicメンバー 'FieldOfPublic' なら正規に値を設定すべき"))
                End Try
            End Sub

            <Test()> Public Sub PublicクラスのPublicFieldの値取得も例外になる()
                Dim vo As New PublicVoField
                Try
                    NonPublicUtil.GetField(vo, "FieldOfPublic")
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField' でPublicメンバー 'FieldOfPublic' なら正規に値を取得すべき"))
                End Try
            End Sub

            <Test()> Public Sub PrivateクラスのPrivateSharedFieldに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoField")
                NonPublicUtil.SetFieldOfShared(privateType, "SharedFieldOfPrivate", "xyz")

                Assert.That(NonPublicUtil.GetFieldOfShared(privateType, "SharedFieldOfPrivate"), [Is].EqualTo("xyz"))
            End Sub

            <Test()> Public Sub PrivateクラスのPublicSharedFieldに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoField")
                NonPublicUtil.SetFieldOfShared(privateType, "SharedFieldOfPublic", "yyy")

                Assert.That(NonPublicUtil.GetFieldOfShared(privateType, "SharedFieldOfPublic"), [Is].EqualTo("yyy"))
            End Sub

            <Test()> Public Sub Privateクラスの親クラスのPrivateSharedFieldに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+SubPrivateVoField")
                NonPublicUtil.SetFieldOfShared(privateType, "SharedFieldOfPrivate", "xyz")

                Assert.That(NonPublicUtil.GetFieldOfShared(privateType, "SharedFieldOfPrivate"), [Is].EqualTo("xyz"))
            End Sub

            <Test()> Public Sub PublicクラスのPrivateSharedFieldに値を設定できる()
                NonPublicUtil.SetFieldOfShared(Of PublicVoField)("SharedFieldOfPrivate", "xxx")

                Assert.That(PublicVoField.GetSharedFieldOfPrivate, [Is].EqualTo("xxx"))
            End Sub

            <Test()> Public Sub PublicクラスのPublicSharedFieldに値設定は例外になる()
                Try
                    NonPublicUtil.SetFieldOfShared(GetType(PublicVoField), "SharedFieldOfPublic", "yyy")
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField' でPublicメンバー 'SharedFieldOfPublic' なら正規に値を設定すべき"))
                End Try
            End Sub

            <Test()> Public Sub PublicクラスのPublicSharedFieldの値取得も例外になる()
                Try
                    NonPublicUtil.GetFieldOfShared(GetType(PublicVoField), "SharedFieldOfPublic")
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField' でPublicメンバー 'SharedFieldOfPublic' なら正規に値を取得すべき"))
                End Try
            End Sub

            <Test()> Public Sub SetField_プロパティが無ければ例外になる()
                Dim vo As New PublicVoField
                Try
                    NonPublicUtil.SetField(vo, "NonExistField", "abc")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField にメンバー 'NonExistField' は無い"))
                End Try
            End Sub

            <Test()> Public Sub SetFieldOfShared_Sharedプロパティが無ければ例外になる()
                Try
                    NonPublicUtil.SetFieldOfShared(Of PublicVoField)("NonExistField", "abc")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField にSharedメンバー 'NonExistField' は無い"))
                End Try
            End Sub

            <Test()> Public Sub プロパティにSetFieldOfSharedは例外になる()
                Try
                    NonPublicUtil.SetFieldOfShared(Of PublicVoField)("FieldOfPublic", "abc")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField にSharedメンバー 'FieldOfPublic' は無い"))
                End Try
            End Sub

            <Test()> Public Sub SharedプロパティにSetFieldは例外になる()
                Dim vo As New PublicVoField
                Try
                    NonPublicUtil.SetField(vo, "SharedFieldOfPublic", "abc")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField にメンバー 'SharedFieldOfPublic' は無い"))
                End Try
            End Sub

            <Test()> Public Sub GetField_プロパティが無ければ例外になる()
                Dim vo As New PublicVoField
                Try
                    NonPublicUtil.GetField(vo, "NonExistField")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField にメンバー 'NonExistField' は無い"))
                End Try
            End Sub

            <Test()> Public Sub GetFieldOfShared_Sharedプロパティが無ければ例外になる()
                Try
                    NonPublicUtil.GetFieldOfShared(Of PublicVoField)("NonExistField")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField にSharedメンバー 'NonExistField' は無い"))
                End Try
            End Sub

            <Test()> Public Sub プロパティにGetFieldOfSharedは例外になる()
                Try
                    NonPublicUtil.GetFieldOfShared(Of PublicVoField)("FieldOfPublic")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField にSharedメンバー 'FieldOfPublic' は無い"))
                End Try
            End Sub

            <Test()> Public Sub SharedプロパティにGetFieldは例外になる()
                Dim vo As New PublicVoField
                Try
                    NonPublicUtil.GetField(vo, "SharedFieldOfPublic")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoField にメンバー 'SharedFieldOfPublic' は無い"))
                End Try
            End Sub

        End Class

        Public Class GetSetPropertyTest : Inherits NonPublicUtilTest

            <Test()> Public Sub PrivateクラスのPrivatePropertyに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoProperty")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                NonPublicUtil.SetProperty(actual, "ValueOfPrivate", "abc")

                Assert.That(NonPublicUtil.GetProperty(actual, "ValueOfPrivate"), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub PrivateクラスのPublicPropertyだけどPrivateSetなPropertyに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoProperty")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                NonPublicUtil.SetProperty(actual, "ValueOfPrivateSet", "bbc")

                Assert.That(NonPublicUtil.GetProperty(actual, "ValueOfPrivateSet"), [Is].EqualTo("bbc"))
            End Sub

            <Test()> Public Sub Privateクラスの親クラスのPrivatePropertyに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+SubPrivateVoProperty")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                NonPublicUtil.SetProperty(actual, "ValueOfPrivate", "abc")

                Assert.That(NonPublicUtil.GetProperty(actual, "ValueOfPrivate"), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub PrivateクラスのPublicPropertyに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoProperty")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType)

                NonPublicUtil.SetProperty(actual, "ValueOfPublic", "qwerty")

                Assert.That(NonPublicUtil.GetProperty(actual, "ValueOfPublic"), [Is].EqualTo("qwerty"))
            End Sub

            <Test()> Public Sub PublicクラスのPrivatePropertyに値を設定できる()
                Dim actual As New PublicVoProperty

                NonPublicUtil.SetProperty(actual, "ValueOfPrivate", "def")

                Assert.That(NonPublicUtil.GetProperty(actual, "ValueOfPrivate"), [Is].EqualTo("def"))
            End Sub

            <Test()> Public Sub PublicクラスのPublicPropertyだけどPrivateSetなPropertyに値を設定できる()
                Dim actual As New PublicVoProperty

                NonPublicUtil.SetProperty(actual, "ValueOfPrivateSet", "ghj")

                Dim valueOfPrivateSet As String = actual.ValueOfPrivateSet
                Assert.That(valueOfPrivateSet, [Is].EqualTo("ghj"))
            End Sub

            <Test()> Public Sub PublicクラスのPublicPropertyに値設定は例外になる()
                Dim actual As New PublicVoProperty
                Try
                    NonPublicUtil.SetProperty(actual, "ValueOfPublic", "qwerty")
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoProperty' でPublicプロパティ 'ValueOfPublic' なら正規に値を設定すべき"))
                End Try
            End Sub

            <Test()> Public Sub PrivateクラスのPrivateSharedなPropertyに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoProperty")

                NonPublicUtil.SetPropertyOfShared(privateType, "SharedValueOfPrivate", "wwww")

                Assert.That(NonPublicUtil.GetPropertyOfShared(privateType, "SharedValueOfPrivate"), [Is].EqualTo("wwww"))
            End Sub

            <Test()> Public Sub PrivateクラスのPublicSharedPropertyだけどPrivateSetなPropertyに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoProperty")

                NonPublicUtil.SetPropertyOfShared(privateType, "SharedValueOfPrivateSet", "333")

                Assert.That(NonPublicUtil.GetPropertyOfShared(privateType, "SharedValueOfPrivateSet"), [Is].EqualTo("333"))
            End Sub

            <Test()> Public Sub Privateクラスの親クラスのPrivateSharedなPropertyに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+SubPrivateVoProperty")

                NonPublicUtil.SetPropertyOfShared(privateType, "SharedValueOfPrivate", "wwww")

                Assert.That(NonPublicUtil.GetPropertyOfShared(privateType, "SharedValueOfPrivate"), [Is].EqualTo("wwww"))
            End Sub

            <Test()> Public Sub PrivateクラスのPublicSharedなPropertyに値を設定できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateVoProperty")

                NonPublicUtil.SetPropertyOfShared(privateType, "SharedValueOfPublic", "vvvv")

                Assert.That(NonPublicUtil.GetPropertyOfShared(privateType, "SharedValueOfPublic"), [Is].EqualTo("vvvv"))
            End Sub

            <Test()> Public Sub PublicクラスのPrivateSharedなPropertyに値を設定できる()
                NonPublicUtil.SetPropertyOfShared(Of PublicVoProperty)("SharedValueOfPrivate", "xxx")

                Assert.That(NonPublicUtil.GetPropertyOfShared(Of PublicVoProperty)("SharedValueOfPrivate"), [Is].EqualTo("xxx"))
            End Sub

            <Test()> Public Sub PublicクラスのPublicSharedPropertyだけどPrivateSetなPropertyに値を設定できる()
                NonPublicUtil.SetPropertyOfShared(GetType(PublicVoProperty), "SharedValueOfPrivateSet", "yy")

                Dim sharedValueOfPrivateSet As String = PublicVoProperty.SharedValueOfPrivateSet
                Assert.That(sharedValueOfPrivateSet, [Is].EqualTo("yy"))
            End Sub

            <Test()> Public Sub PublicクラスのPublicSharedなPropertyに値設定は例外になる()
                Try
                    NonPublicUtil.SetPropertyOfShared(Of PublicVoProperty)("SharedValueOfPublic", "yy")
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoProperty' でPublicプロパティ 'SharedValueOfPublic' なら正規に値を設定すべき"))
                End Try
            End Sub

            <Test()> Public Sub SetProperty_プロパティが無ければ例外になる()
                Dim vo As New PublicVoProperty
                Try
                    NonPublicUtil.SetProperty(vo, "NonExistProperty", "abc")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoProperty にプロパティ 'NonExistProperty' は無い"))
                End Try
            End Sub

            <Test()> Public Sub SetProperty_Sharedプロパティが無ければ例外になる()
                Try
                    NonPublicUtil.SetPropertyOfShared(Of PublicVoProperty)("NonExistProperty", "abc")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoProperty にSharedプロパティ 'NonExistProperty' は無い"))
                End Try
            End Sub

            <Test()> Public Sub プロパティにSetPropertyOfSharedは例外になる()
                Try
                    NonPublicUtil.SetPropertyOfShared(Of PublicVoProperty)("ValueOfPrivate", "abc")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoProperty にSharedプロパティ 'ValueOfPrivate' は無い"))
                End Try
            End Sub

            <Test()> Public Sub SharedプロパティにSetPropertyは例外になる()
                Dim vo As New PrivateVo
                Try
                    NonPublicUtil.SetProperty(vo, "SharedValueOfPrivate", "abc")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PrivateVo にプロパティ 'SharedValueOfPrivate' は無い"))
                End Try
            End Sub

            <Test()> Public Sub GetProperty_プロパティが無ければ例外になる()
                Dim vo As New PublicVoProperty
                Try
                    NonPublicUtil.GetProperty(vo, "NonExistProperty")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoProperty にプロパティ 'NonExistProperty' は無い"))
                End Try
            End Sub

            <Test()> Public Sub GetProperty_Sharedプロパティが無ければ例外になる()
                Try
                    NonPublicUtil.GetPropertyOfShared(Of PublicVoProperty)("NonExistProperty")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoProperty にSharedプロパティ 'NonExistProperty' は無い"))
                End Try
            End Sub

            <Test()> Public Sub プロパティにGetPropertyOfSharedは例外になる()
                Try
                    NonPublicUtil.GetPropertyOfShared(Of PublicVoProperty)("ValueOfPrivate")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicVoProperty にSharedプロパティ 'ValueOfPrivate' は無い"))
                End Try
            End Sub

            <Test()> Public Sub SharedプロパティにGetPropertyは例外になる()
                Dim vo As New PrivateVo
                Try
                    NonPublicUtil.GetProperty(vo, "SharedValueOfPrivate")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PrivateVo にプロパティ 'SharedValueOfPrivate' は無い"))
                End Try
            End Sub

        End Class

        Public Class InvokeMethodTest : Inherits NonPublicUtilTest

            <Test()> Public Sub PrivateクラスのPrivateMethodを実行できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateMethod")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType, "abc")

                Dim result As Object = NonPublicUtil.InvokeMethod(privateType, actual, "CombinePrivate", "xyz")
                Assert.That(result, [Is].EqualTo("-abcxyz"))
            End Sub

            <Test()> Public Sub PrivateクラスのPublicMethodを実行できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateMethod")
                Dim actual As Object = NonPublicUtil.NewInstance(privateType, "ddd")

                Dim result As Object = NonPublicUtil.InvokeMethod(privateType, actual, "CombinePublic", "eee")
                Assert.That(result, [Is].EqualTo("+dddeee"))
            End Sub

            <Test()> Public Sub PublicクラスのPrivateMethodを実行できる()
                Dim actual As New PublicMethod("qqq")

                Dim result As Object = NonPublicUtil.InvokeMethod(Of PublicMethod)(actual, "CombinePrivate", "www")
                Assert.That(result, [Is].EqualTo("-qqqwww"))
            End Sub

            <Test()> Public Sub PublicクラスのPublicMethod実行は例外になる()
                Dim actual As New PublicMethod("qqq")
                Try
                    NonPublicUtil.InvokeMethod(Of PublicMethod)(actual, "CombinePublic", "www")
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicMethod' でPublicメソッド 'CombinePublic' なら正規に値を設定すべき"))
                End Try
            End Sub

            <Test()> Public Sub PrivateクラスのPrivateSharedMethodを実行できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateMethod")

                Dim result As Object = NonPublicUtil.InvokeMethodOfShared(privateType, "CombineSharedPrivate", "xyz")
                Assert.That(result, [Is].EqualTo("-(s)xyz"))
            End Sub

            <Test()> Public Sub PrivateクラスのPublicSharedMethodを実行できる()
                Dim privateType As Type = NonPublicUtil.GetTypeOfNonPublic(Of NonPublicUtilTest)("Fhi.Fw.TestUtil.Test.NonPublicUtilTest+Holder+PrivateMethod")

                Dim result As Object = NonPublicUtil.InvokeMethodOfShared(privateType, "CombineSharedPublic", "eee")
                Assert.That(result, [Is].EqualTo("+(s)eee"))
            End Sub

            <Test()> Public Sub PublicクラスのPrivateSharedMethodを実行できる()
                Dim result As Object = NonPublicUtil.InvokeMethodOfShared(Of PublicMethod)("CombineSharedPrivate", "www")
                Assert.That(result, [Is].EqualTo("-(s)www"))
            End Sub

            <Test()> Public Sub PublicクラスのPublicSharedMethod実行は例外になる()
                Try
                    NonPublicUtil.InvokeMethodOfShared(Of PublicMethod)("CombineSharedPublic", "www")
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Publicクラス'Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicMethod' でPublicSharedメソッド 'CombineSharedPublic' なら正規に値を設定すべき"))
                End Try
            End Sub

            <Test()> Public Sub InvokeMethod_Methodが無ければ例外になる()
                Dim instance As New PublicMethod
                Try
                    NonPublicUtil.InvokeMethod(instance, "NotExistMethod", "www")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicMethod のメソッド 'NotExistMethod' はない"))
                End Try
            End Sub

            <Test()> Public Sub InvokeMethodOfShared_SharedMethodが無ければ例外になる()
                Try
                    NonPublicUtil.InvokeMethodOfShared(Of PublicMethod)("NotExistMethod", "www")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicMethod のSharedメソッド 'NotExistMethod' はない"))
                End Try
            End Sub

            <Test()> Public Sub InstanceMethodにInvokeMethodOfSharedは例外になる()
                Try
                    NonPublicUtil.InvokeMethodOfShared(Of PublicMethod)("CombinePrivate", "www")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicMethod のSharedメソッド 'CombinePrivate' はない"))
                End Try
            End Sub

            <Test()> Public Sub SharedMethodにInvokeMethodは例外になる()
                Dim instance As New PublicMethod
                Try
                    NonPublicUtil.InvokeMethod(instance, "CombineSharedPrivate", "www")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("型 Fhi.Fw.TestUtil.Test.NonPublicUtilTest+PublicMethod のメソッド 'CombineSharedPrivate' はない"))
                End Try
            End Sub

        End Class

    End Class
End Namespace
