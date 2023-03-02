Imports System.Reflection
Imports Fhi.Fw.Domain

Namespace Util.Fixed.Extend

    ''' <summary>
    ''' 固定長文字列と固定長Voとの相互変換（連携？）を担うクラス
    ''' </summary>
    ''' <typeparam name="T">固定長Voの型</typeparam>
    ''' <remarks></remarks>
    Public Class FixedVoBroker(Of T)

#Region "nested class"
        Private Class MyFixedDefine : Inherits AbstractFixedDefine

            Public ReadOnly Root As FixedGroup

            Public Sub New(ByVal root As FixedGroup)
                Me.Root = root
            End Sub

            ''' <summary>
            ''' 固定長Root情報を取得する
            ''' </summary>
            ''' <returns>固定長Root情報</returns>
            ''' <remarks></remarks>
            Protected Overrides Function GetRootEntryImpl() As FixedGroup
                Return Root
            End Function
        End Class

#End Region

#Region "Public properties..."
        ''' <summary>固定長列設定したVo</summary>
        Public ReadOnly Property Vo() As T
            Get
                Return _vo
            End Get
        End Property

        ''' <summary>固定長文字列の全体</summary>
        Public Property FixedString() As String
            Get
                Return aDefine.FixedString
            End Get
            Set(ByVal value As String)
                aDefine.FixedString = value
            End Set
        End Property

        ''' <summary>数値型がNullのときzero埋めするか？</summary>
        Public Property IsZeroPaddingIfNull() As Boolean
            Get
                Return aDefine.IsZeroPaddingIfNull
            End Get
            Set(ByVal value As Boolean)
                aDefine.IsZeroPaddingIfNull = value
                aDefine.InitializeFixedString()
                ApplyFixedStringToVo()
            End Set
        End Property
#End Region

        Private ReadOnly builder As FixedRuleBuilder(Of T)
        Private ReadOnly aDefine As MyFixedDefine
        Private _vo As T

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">固定長列設定ルール（ラムダ式）</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As FixedRule(Of T).RuleConfigure)
            Me.New(New FixedRule(Of T)(rule))
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">固定長列設定ルール</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As IFixedRule(Of T))

            builder = New FixedRuleBuilder(Of T)(rule)

            aDefine = New MyFixedDefine(builder.MakeRootGroup)

            NewVoInstanceAndInitialize()
        End Sub

        ''' <summary>
        ''' Voのインスタンスを再生成し、初期化する（固定長列設定に従い、Voと固定長文字列の初期化を行う）
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub NewVoInstanceAndInitialize()
            _vo = VoUtil.NewInstance(Of T)()
            Initialize()
        End Sub

        ''' <summary>
        ''' 初期化する（固定長列設定に従い、Voと固定長文字列の初期化を行う）
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Initialize()

            PerformInitializeVo()
            ApplyVoToFixedString()
        End Sub

        ''' <summary>
        ''' 固定長列設定に従いVoを初期化（繰り返しやインスタンスの生成）
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub PerformInitializeVo()

            RecurInitializeVo(_vo, builder.ResultProperties)
        End Sub

        Private Sub RecurInitializeVo(ByVal vo As Object, ByVal properties As ICollection(Of FixedRuleProperty))
            For Each aProperty As FixedRuleProperty In properties
                Dim initializeValues As New List(Of Object)
                If aProperty.IsGroup Then
                    If TypeUtil.IsTypeCollection(aProperty.Info.PropertyType) Then
                        For i As Integer = 0 To aProperty.Repeat - 1
                            Dim originValue As Object = VoUtil.NewElementInstanceFromCollectionType(aProperty.Info.PropertyType)
                            RecurInitializeVo(originValue, aProperty.Builder.ResultProperties)
                            initializeValues.Add(originValue)
                        Next
                    Else
                        Dim originValue As Object = Activator.CreateInstance(aProperty.Info.PropertyType)
                        RecurInitializeVo(originValue, aProperty.Builder.ResultProperties)
                        initializeValues.Add(originValue)
                    End If

                ElseIf TypeUtil.IsTypeCollection(aProperty.Info.PropertyType) _
                        AndAlso TypeUtil.DetectElementType(aProperty.Info.PropertyType).Equals(GetType(String)) Then
                    For i As Integer = 0 To aProperty.Repeat - 1
                        initializeValues.Add(String.Empty)
                    Next

                ElseIf aProperty.Info.PropertyType.Equals(GetType(String)) Then
                    initializeValues.Add(String.Empty)
                End If

                If TypeUtil.IsTypeCollection(aProperty.Info.PropertyType) Then
                    Dim newValues As Object = VoUtil.NewCollectionInstanceWithInitialElement(aProperty.Info.PropertyType, aProperty.Repeat, initializeValues)
                    aProperty.Info.SetValue(vo, newValues, Nothing)

                ElseIf 0 < initializeValues.Count Then
                    aProperty.Info.SetValue(vo, initializeValues(0), Nothing)
                End If
            Next

        End Sub

        ''' <summary>
        ''' Voの値で固定長文字列を変更する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ApplyVoToFixedString()

            RecurApplyVoToFixedString(_vo, "", builder.ResultProperties, aDefine)
        End Sub

        Private Sub RecurApplyVoToFixedString(ByVal vo As Object, ByVal prefixPath As String, ByVal properties As ICollection(Of FixedRuleProperty), ByVal aDefine As MyFixedDefine)

            For Each aProperty As FixedRuleProperty In properties
                If TypeUtil.IsTypeCollection(aProperty.Info.PropertyType) Then
                    If aProperty.IsGroup Then
                        Dim values As ICollection(Of Object) = VoUtil.ConvObjectToCollection(aProperty.Info.GetValue(vo, Nothing))
                        Dim index As Integer = 0
                        For Each value As Object In values
                            Dim prefix As String = String.Format("{0}{1}[{2}].", prefixPath, aProperty.Name, index)
                            RecurApplyVoToFixedString(value, prefix, aProperty.Builder.ResultProperties, aDefine)
                            index += 1
                        Next
                    Else
                        Dim values As ICollection(Of Object) = VoUtil.ConvObjectToCollection(aProperty.Info.GetValue(vo, Nothing))
                        Dim index As Integer = 0
                        For Each value As Object In values
                            Dim name As String = String.Format("{0}{1}[{2}]", prefixPath, aProperty.Name, index)
                            aDefine.SetValue(name, value)
                            index += 1
                        Next
                    End If
                ElseIf aProperty.IsGroup Then
                    Dim value As Object = aProperty.Info.GetValue(vo, Nothing)
                    Dim prefix As String = String.Format("{0}{1}.", prefixPath, aProperty.Name)
                    RecurApplyVoToFixedString(value, prefix, aProperty.Builder.ResultProperties, aDefine)
                Else
                    Dim value As Object = aProperty.Info.GetValue(vo, Nothing)
                    aDefine.SetValue(prefixPath & aProperty.Name, value)
                End If
            Next
        End Sub

        ''' <summary>
        ''' 固定長文字列の値でVoを変更する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ApplyFixedStringToVo()

            RecurApplyFixedStringToVo(_vo, "", builder.ResultProperties, aDefine)
        End Sub

        Private Sub RecurApplyFixedStringToVo(ByVal vo As Object, ByVal prefixPath As String, ByVal properties As ICollection(Of FixedRuleProperty), ByVal aDefine As MyFixedDefine)

            For Each aProperty As FixedRuleProperty In properties
                Dim value As Object = aProperty.Info.GetValue(vo, Nothing)
                If value Is Nothing AndAlso (aProperty.IsGroup OrElse TypeUtil.IsTypeCollection(aProperty.Info.PropertyType)) Then
                    RecurInitializeVo(vo, New FixedRuleProperty() {aProperty})
                    value = aProperty.Info.GetValue(vo, Nothing)
                End If

                If aProperty.IsGroup Then
                    If TypeUtil.IsTypeCollection(aProperty.Info.PropertyType) Then
                        Dim index As Integer = 0
                        For Each collectionValue As Object In VoUtil.ConvObjectToCollection(value)
                            Dim prefix As String = String.Format("{0}{1}[{2}].", prefixPath, aProperty.Name, index)
                            RecurApplyFixedStringToVo(collectionValue, prefix, aProperty.Builder.ResultProperties, aDefine)
                            index += 1
                        Next
                    Else
                        Dim prefix As String = String.Format("{0}{1}.", prefixPath, aProperty.Name)
                        RecurApplyFixedStringToVo(value, prefix, aProperty.Builder.ResultProperties, aDefine)
                    End If

                ElseIf aProperty.Info.PropertyType.IsArray Then
                    Dim anArray As Array = DirectCast(value, Array)
                    If anArray.Length <> aProperty.Repeat Then
                        ' サイズが違うなら別インスタンスで作り直す
                        anArray = DirectCast(VoUtil.NewCollectionInstance(aProperty.Info.PropertyType, aProperty.Repeat), Array)
                        aProperty.Info.SetValue(vo, anArray, Nothing)
                    End If
                    For i As Integer = 0 To anArray.Length - 1
                        Dim arrayValue As Object = aDefine.GetValue(String.Format("{0}{1}[{2}]", prefixPath, aProperty.Name, i))
                        anArray.SetValue(arrayValue, i)
                    Next

                ElseIf GetType(IList).IsAssignableFrom(aProperty.Info.PropertyType) Then
                    DirectCast(value, IList).Clear()
                    Dim addMethod As MethodInfo = value.GetType.GetMethod("Add")
                    For i As Integer = 0 To aProperty.Repeat - 1
                        Dim collectionValue As Object = aDefine.GetValue(String.Format("{0}{1}[{2}]", prefixPath, aProperty.Name, i))
                        addMethod.Invoke(value, New Object() {If(collectionValue, VoUtil.NewElementInstanceFromCollectionType(aProperty.Info.PropertyType))})
                    Next

                ElseIf GetType(PrimitiveValueObject).IsAssignableFrom(aProperty.Info.PropertyType) Then
                    Dim propertyValue As Object = ConvertInnerValueType(aProperty.Info.PropertyType, aDefine.GetValue(prefixPath & aProperty.Name))
                    Dim pvo As Object = Activator.CreateInstance(aProperty.Info.PropertyType, propertyValue)
                    aProperty.Info.SetValue(vo, pvo, Nothing)

                Else
                    aProperty.Info.SetValue(vo, aDefine.GetValue(prefixPath & aProperty.Name), Nothing)
                End If
            Next
        End Sub

        Private Function ConvertInnerValueType(propertyType As Type, value As Object) As Object
            Dim constructor As ConstructorInfo = ValueObject.DetectConstructor(propertyType)
            Dim parameterInfos As ParameterInfo() = constructor.GetParameters
            If parameterInfos.Length <> 1 Then
                Throw New InvalidOperationException("引数のないPrimitiveValueObjectは対応できない")
            End If

            Dim innerValueType As Type = parameterInfos(0).ParameterType
            If innerValueType Is GetType(Int32) Then
                Return Convert.ToInt32(value)
            ElseIf innerValueType Is GetType(Int64) Then
                Return Convert.ToInt64(value)
            ElseIf innerValueType Is GetType(Int16) Then
                Return Convert.ToInt16(value)
            ElseIf innerValueType Is GetType(Decimal) Then
                Return Convert.ToDecimal(value)
            End If
            Return value
        End Function

    End Class
End Namespace