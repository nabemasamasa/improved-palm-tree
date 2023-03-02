Imports Fhi.Fw.Domain
Imports Fhi.Fw.Util
Imports System.Linq.Expressions
Imports System.Reflection

Namespace TestUtil.DebugString
    ''' <summary>
    ''' 検証用文字列作成ルールの作成を担うクラス
    ''' </summary>
    ''' <typeparam name="T">検証対象の型T</typeparam>
    ''' <remarks></remarks>
    Public Class DebugStringRuleBuilder(Of T)

#Region "Nested classes..."
        Private Class BinderImpl : Implements IDebugStringRuleBinder

            Private ReadOnly builder As DebugStringRuleBuilder(Of T)

            Public Sub New(ByVal builder As DebugStringRuleBuilder(Of T))
                Me.builder = builder
            End Sub

            Public Function Bind(ByVal ParamArray fields As Object()) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.Bind
                For Each aField As Object In fields
                    builder.Message(aField)
                Next
                Return Me
            End Function

            Public Function BindWithTitle(ByVal field As Object, ByVal title As String) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.BindWithTitle
                builder.Message(field, title)
                Return Me
            End Function

            Public Function BindFuncWithTitle(Of T1)(ByVal fieldLambda As Expression(Of Func(Of T1)), ByVal title As String) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.BindFuncWithTitle
                Dim lambda As LambdaExpression = TryCast(fieldLambda, LambdaExpression)
                If lambda Is Nothing Then
                    Throw New ArgumentException("Lambda式であるべき", "fieldLambda")
                End If
                builder.Message(lambda, title)
                Return Me
            End Function

            Public Function JoinDetails(Of T1)(ByVal field As IEnumerable(Of T1), ByVal aConfigure As DebugStringRuleBuilder(Of T1).Configure) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.JoinDetails
                Dim maker As New DebugStringMaker(Of T1)(aConfigure)
                builder.MessageDetail(field, maker)
                Return Me
            End Function

            Public Function JoinDetails(Of T1)(ByVal field As CollectionObject(Of T1), ByVal aConfigure As DebugStringRuleBuilder(Of T1).Configure) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.JoinDetails
                Dim maker As New DebugStringMaker(Of T1)(aConfigure)
                builder.MessageDetail(field, maker)
                Return Me
            End Function

            Public Function JoinFromSideBySide(Of T1)(ByVal field As IEnumerable(Of T1)) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.JoinFromSideBySide
                Return JoinFromSideBySide(field, Function(defineBy As IDebugStringRuleBinder, tValue As T1) defineBy.Bind(tValue))
            End Function

            Public Function JoinFromSideBySide(Of T1)(ByVal field As IEnumerable(Of T1), ByVal sideBySideConfigure As DebugStringRuleBuilder(Of T1).Configure) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.JoinFromSideBySide
                Return JoinFromSideBySide(field, Nothing, sideBySideConfigure)
            End Function

            Public Function JoinFromSideBySide(Of T1)(ByVal field As IEnumerable(Of T1), ByVal fixedRepeatCount As Integer?, ByVal sideBySideConfigure As DebugStringRuleBuilder(Of T1).Configure) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.JoinFromSideBySide
                Dim maker As New DebugStringMaker(Of T1)(sideBySideConfigure)
                builder.MessageSideBySide(field, fixedRepeatCount, maker)
                Return Me
            End Function

            Public Function JoinFromSideBySide(Of T1)(ByVal field As CollectionObject(Of T1)) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.JoinFromSideBySide
                Return JoinFromSideBySide(field, Function(defineBy As IDebugStringRuleBinder, tValue As T1) defineBy.Bind(tValue))
            End Function

            Public Function JoinFromSideBySide(Of T1)(ByVal field As CollectionObject(Of T1), ByVal sideBySideConfigure As DebugStringRuleBuilder(Of T1).Configure) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.JoinFromSideBySide
                Return JoinFromSideBySide(field, Nothing, sideBySideConfigure)
            End Function

            Public Function JoinFromSideBySide(Of T1)(ByVal field As CollectionObject(Of T1), ByVal fixedRepeatCount As Integer?, ByVal sideBySideConfigure As DebugStringRuleBuilder(Of T1).Configure) As IDebugStringRuleBinder Implements IDebugStringRuleBinder.JoinFromSideBySide
                Dim maker As New DebugStringMaker(Of T1)(sideBySideConfigure)
                builder.MessageSideBySide(field, fixedRepeatCount, maker)
                Return Me
            End Function
        End Class
        Private Interface IValue
            ''' <summary>明細数</summary>
            ReadOnly Property DetailsCount As Integer
            ''' <summary>
            ''' 値情報を展開する
            ''' </summary>
            ''' <param name="index">行index</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Expand(index As Integer?) As String()
        End Interface
        Private Class NormalValue : Implements IValue
            Private ReadOnly value As Object
            Public Sub New(value As Object)
                Me.value = value
            End Sub
            Public ReadOnly Property DetailsCount() As Integer Implements IValue.DetailsCount
                Get
                    Return 1
                End Get
            End Property
            Public Function Expand(index As Integer?) As String() Implements IValue.Expand
                Dim debugValue As String = DebugStringMaker.ConvDebugValue(value)
                If IsNumeric(debugValue) Then
                    debugValue = StringUtil.TrimEndDecimalZero(debugValue)
                End If
                Return {debugValue}
            End Function
        End Class
        Private Class ValueOfDetails : Implements IValue
            Private ReadOnly makeDataCallback As Func(Of String()())
            Public Sub New(makeDataCallback As Func(Of String()()))
                Me.makeDataCallback = makeDataCallback
            End Sub
            Public ReadOnly Property DetailsCount() As Integer Implements IValue.DetailsCount
                Get
                    Return makeDataCallback.Invoke.Length
                End Get
            End Property
            Public Function Expand(index As Integer?) As String() Implements IValue.Expand
                Dim data As String()() = makeDataCallback.Invoke
                If data.Length <= If(index, 0) Then
                    Return Enumerable.Range(0, data(0).Length).Select(Function(i) DebugStringMaker.ConvDebugValue(Nothing)).ToArray
                End If
                Return data(If(index, 0))
            End Function
        End Class
        Private Class ValueOfSideBySide : Implements IValue
            Private ReadOnly makeDataCallback As Func(Of String()())
            Private ReadOnly getMaxRepeatCallback As Func(Of Integer)
            Public Sub New(makeDataCallback As Func(Of String()()), getMaxRepeatCallback As Func(Of Integer))
                Me.makeDataCallback = makeDataCallback
                Me.getMaxRepeatCallback = getMaxRepeatCallback
            End Sub
            Public ReadOnly Property DetailsCount() As Integer Implements IValue.DetailsCount
                Get
                    Return 1
                End Get
            End Property
            Public Function Resolve(index As Integer?) As String() Implements IValue.Expand
                Dim maxRepeat As Integer = getMaxRepeatCallback.Invoke
                Dim data As String()() = makeDataCallback.Invoke
                Return Enumerable.Range(0, maxRepeat).SelectMany(
                    Function(i) If(data.Length <= i,
                                   Enumerable.Range(0, data(0).Length).Select(Function(j) DebugStringMaker.ConvDebugValue(Nothing)).ToArray,
                                   data(i))).ToArray
            End Function
        End Class
        Private Class SideBySideSet
            ''' <summary>検証文字列作成</summary>
            Public ReadOnly Maker As DebugStringMakerWildcard
            ''' <summary>固定表示繰り返し数</summary>
            Private ReadOnly FixedRepeatCount As Integer?
            ''' <summary>実最大繰り返し数</summary>
            Public MaxRepeatCount As Integer
            Public Sub New(maker As DebugStringMakerWildcard, fixedRepeatCount As Integer?)
                Me.Maker = maker
                Me.FixedRepeatCount = fixedRepeatCount
            End Sub
            ''' <summary>
            ''' 繰り返し数を解決する
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function ResolveCountOfRepeat() As Integer
                Return If(FixedRepeatCount, MaxRepeatCount)
            End Function
        End Class
#End Region

        Private ReadOnly voMarker As New VoPropertyMarker
        Private ReadOnly ordinalProperties As New List(Of DebugStringRuleProperty)

        ''' <summary>親列名を抑止するか？</summary>
        Friend Property SuppressesParentColumnName As Boolean

        ''' <summary>
        ''' 検証用文字列作成の定義を行うDelegate
        ''' </summary>
        ''' <param name="defineBy">項目定義の設定先</param>
        ''' <param name="vo">項目定義用の列情報</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Delegate Function Configure(ByVal defineBy As IDebugStringRuleBinder, ByVal vo As T) As IDebugStringRuleBinder

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">区切り列順設定のルール</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As Configure)
            Dim vo As T = voMarker.CreateMarkedVo(Of T)()
            rule.Invoke(New BinderImpl(Me), vo)
        End Sub

        ''' <summary>Voアクセス用の区切り列順設定情報[]</summary>
        Friend ReadOnly Property ResultProperties() As DebugStringRuleProperty()
            Get
                Return ordinalProperties.ToArray
            End Get
        End Property

        ''' <summary>
        ''' 項目定義のメッセージ
        ''' </summary>
        ''' <param name="fieldValue">項目定義値の列値</param>
        ''' <remarks></remarks>
        Private Sub Message(ByVal fieldValue As Object)
            Message(fieldValue, Nothing)
        End Sub

        ''' <summary>
        ''' 項目定義と出力タイトル設定のメッセージ
        ''' </summary>
        ''' <param name="fieldValue">項目定義値の列値</param>
        ''' <param name="title">出力タイトル</param>
        ''' <remarks></remarks>
        Private Sub Message(ByVal fieldValue As Object, ByVal title As String)
            voMarker.Assert3OrMoreBoolean(fieldValue)
            Dim debugStringRuleProperty As New DebugStringRuleProperty(voMarker.GetMemberName(fieldValue), Function(vo) voMarker.GetValue(fieldValue, vo), title)
            ordinalProperties.Add(debugStringRuleProperty)
        End Sub
        '
        ''' <summary>
        ''' 項目定義と出力タイトル設定のメッセージ
        ''' </summary>
        ''' <param name="lambda">項目定義値の列値取得Lambda</param>
        ''' <param name="title">出力タイトル</param>
        ''' <remarks></remarks>
        Private Sub Message(ByVal lambda As LambdaExpression, ByVal title As String)
            ordinalProperties.Add(New DebugStringRuleProperty(Nothing, lambda, title))
        End Sub

        Private ReadOnly detailMakerByMarker As New Dictionary(Of DebugStringRuleProperty, DebugStringMakerWildcard)
        ''' <summary>
        ''' 項目定義と明細項目の出力設定をするメッセージ
        ''' </summary>
        ''' <param name="fieldValue">項目定義値の列値</param>
        ''' <param name="maker">明細項目の出力設定</param>
        ''' <remarks></remarks>
        Private Sub MessageDetail(fieldValue As Object, maker As DebugStringMakerWildcard)
            Me.Message(fieldValue)
            detailMakerByMarker.Add(ordinalProperties.Last, maker)
        End Sub

        Private ReadOnly sideBySideMakerByMarker As New Dictionary(Of DebugStringRuleProperty, SideBySideSet)
        ''' <summary>
        ''' 項目定義と明細項目の出力設定をするメッセージ
        ''' </summary>
        ''' <param name="fieldValue">項目定義値の列値</param>
        ''' <param name="fixedRepeatCount">固定表示繰り返し数</param>
        ''' <param name="maker">明細項目の出力設定</param>
        ''' <remarks></remarks>
        Private Sub MessageSideBySide(ByVal fieldValue As Object, ByVal fixedRepeatCount As Integer?, ByVal maker As DebugStringMakerWildcard)
            Me.Message(fieldValue)
            sideBySideMakerByMarker.Add(ordinalProperties.Last, New SideBySideSet(maker:=maker, fixedRepeatCount:=fixedRepeatCount))
        End Sub

        ''' <summary>
        ''' 検証用値情報のタイトル一覧を作成する
        ''' </summary>
        ''' <param name="parentTitle">親タイトル</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Function MakeTitles(Optional parentTitle As String = Nothing) As String()
            Return ordinalProperties.SelectMany(
                Function(rule)
                    If detailMakerByMarker.ContainsKey(rule) Then
                        Return detailMakerByMarker(rule).MakeTitles(rule.MakeTitleIfNecessary(parentTitle))
                    ElseIf sideBySideMakerByMarker.ContainsKey(rule) Then
                        Dim repeatCount As Integer = sideBySideMakerByMarker(rule).ResolveCountOfRepeat
                        Dim baseTitles As String() = sideBySideMakerByMarker(rule).Maker.MakeTitles(
                            If(SuppressesParentColumnName, Nothing, rule.MakeTitleIfNecessary(parentTitle)))
                        Return Enumerable.Range(0, repeatCount).SelectMany(
                            Function(i) baseTitles.Select(Function(b) String.Format("{0}#{1}", b, i))).ToArray
                    End If
                    Return New String() {rule.MakeTitleIfNecessary(parentTitle)}
                End Function).ToArray
        End Function

        ''' <summary>
        ''' Empty値の値情報作成Callbackを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Function GetCallbackThatMakeEmptyValues() As Func(Of String())
            Return Function() MakeTitles.Select(Function(t) DebugStringMaker.ConvDebugValue(Nothing)).ToArray
        End Function

        Private ReadOnly beforeBuildData As New List(Of IValue())
        ''' <summary>
        ''' 検証用値情報を作成して格納する
        ''' </summary>
        ''' <param name="record"></param>
        ''' <remarks></remarks>
        Friend Sub StoreAfterMaking(record As T)
            Dim values As New List(Of IValue)
            For Each ruleProperty As DebugStringRuleProperty In ordinalProperties
                Dim value As Object
                If ruleProperty.Lambda IsNot Nothing Then
                    value = VoUtil.InvokeExpressionBy(record, ruleProperty.Lambda)
                Else
                    value = ruleProperty.GetValueCallback.Invoke(record)
                End If
                If detailMakerByMarker.ContainsKey(ruleProperty) Then
                    values.Add(New ValueOfDetails(BuildByWildcardMaker(value, detailMakerByMarker(ruleProperty))))
                ElseIf sideBySideMakerByMarker.ContainsKey(ruleProperty) Then
                    Dim sideBySideSet As DebugStringRuleBuilder(Of T).SideBySideSet = sideBySideMakerByMarker(ruleProperty)
                    Dim buildCallback As Func(Of String()()) = BuildByWildcardMaker(value, sideBySideSet.Maker)
                    Dim sideBySideMatrix As String()() = buildCallback.Invoke
                    If Not IsEmptyValue(value) Then
                        sideBySideSet.MaxRepeatCount = Math.Max(sideBySideMatrix.Length, sideBySideSet.MaxRepeatCount)
                    End If
                    values.Add(New ValueOfSideBySide(buildCallback, Function() sideBySideSet.ResolveCountOfRepeat))
                Else
                    values.Add(New NormalValue(value))
                End If
            Next
            beforeBuildData.Add(values.ToArray)
        End Sub

        ''' <summary>
        ''' 格納情報から値情報だけ構築する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Function BuildValuesWithStored() As String()()
            Return GetCallbackThatBuildValuesAndClearStored.Invoke
        End Function

        ''' <summary>
        ''' 格納情報から値情報を構築して格納情報をクリアするCallbackを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Function GetCallbackThatBuildValuesAndClearStored() As Func(Of String()())
            Dim anotherBuildData As New List(Of IValue())(beforeBuildData)
            FlushBuildData()
            Return Function() anotherBuildData.SelectMany(Function(values) BuildValuesBy(values)).ToArray
        End Function

        Private Function BuildValuesBy(ByVal values As IValue()) As IEnumerable(Of String())
            Dim matrix As New List(Of String())
            Dim maxCountOfValues As Integer = values.Max(Function(v) v.DetailsCount)
            For row As Integer = 0 To maxCountOfValues - 1
                Dim aRow As Integer = row
                matrix.Add(values.SelectMany(Function(value) value.Expand(aRow)).ToArray)
            Next
            Return matrix
        End Function

        Private Sub FlushBuildData()
            beforeBuildData.Clear()
        End Sub

        Private Function IsEmptyValue(ByVal value As Object) As Boolean
            Return (Not TypeOf value Is ICollectionObject AndAlso CollectionUtil.IsEmpty(DirectCast(value, ICollection))) _
                   OrElse (TypeOf value Is ICollectionObject AndAlso DirectCast(value, ICollectionObject).Count = 0)
        End Function

        Private Function BuildByWildcardMaker(ByVal value As Object, ByVal wildcardMaker As DebugStringMakerWildcard) As Func(Of String()())
            If IsEmptyValue(value) Then
                Return Function() {wildcardMaker.GetCallbackThatMakeEmptyValues.Invoke}
            ElseIf TypeOf value Is ICollection Then
                Dim valueCollection As ICollection = DirectCast(value, ICollection)
                For Each detail As Object In valueCollection
                    wildcardMaker.StoreAfterMaking(detail)
                Next
                Return wildcardMaker.GetCallbackThatBuildValuesAndClearStored()
            ElseIf TypeOf value Is ICollectionObject Then
                Dim valueCollection As ICollectionObject = DirectCast(value, ICollectionObject)
                For i As Integer = 0 To valueCollection.Count - 1
                    wildcardMaker.StoreAfterMaking(valueCollection(i))
                Next
                Return wildcardMaker.GetCallbackThatBuildValuesAndClearStored()
            Else
                Throw New InvalidOperationException("未対応の型を#BindDetail()している." & value.GetType.FullName)
            End If
        End Function

    End Class
End Namespace