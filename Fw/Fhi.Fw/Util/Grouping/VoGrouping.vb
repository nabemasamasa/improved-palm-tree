Imports System.Reflection

Namespace Util.Grouping

    ''' <summary>
    ''' グループ化を行う
    ''' </summary>
    ''' <typeparam name="T">グループ化するVOの型</typeparam>
    ''' <remarks></remarks>
    Public Class VoGrouping(Of T)
        ''' <summary>
        ''' LocatorImplの管理を担うクラス
        ''' </summary>
        ''' <remarks></remarks>
        Private Class LocatorOwner
            Private ReadOnly rule As IVoGroupingRule(Of T)
            Private ReadOnly voMarker As New VoPropertyMarker
            Private ReadOnly groupingInfos As New List(Of PropertyInfo)
            Private ReadOnly maxInfos As New List(Of PropertyInfo)
            Private ReadOnly topInfos As New List(Of PropertyInfo)
            ''' <summary>
            ''' コンストラクタ
            ''' </summary>
            ''' <param name="rule">グループ化のルール</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal rule As IVoGroupingRule(Of T))
                Me.rule = rule

                Dim aType As Type = GetType(T)
                Dim vo As T = CType(Activator.CreateInstance(aType), T)
                voMarker.MarkVo(vo)

                rule.Configure(New LocatorImpl(Me), vo)
            End Sub

            ''' <summary>
            ''' LocatorImplのGroup項目メッセージを貰う
            ''' </summary>
            ''' <param name="value">値</param>
            ''' <remarks></remarks>
            Public Sub MessageGroup(ByVal value As Object)
                groupingInfos.Add(voMarker.GetPropertyInfo(value))
            End Sub

            ''' <summary>
            ''' LocatorImplのMax項目メッセージを貰う
            ''' </summary>
            ''' <param name="value">値</param>
            ''' <remarks></remarks>
            Public Sub MessageMax(ByVal value As Object)
                maxInfos.Add(voMarker.GetPropertyInfo(value))
            End Sub

            ''' <summary>
            ''' LocatorImplのTop項目メッセージを貰う
            ''' </summary>
            ''' <param name="value">値</param>
            ''' <remarks></remarks>
            Public Sub MessageTop(ByVal value As Object)
                topInfos.Add(voMarker.GetPropertyInfo(value))
            End Sub

            ''' <summary>
            ''' グループ化項目の値から一意のキーを作成する
            ''' </summary>
            ''' <param name="vo">値object</param>
            ''' <returns>一意のキー</returns>
            ''' <remarks></remarks>
            Public Function MakeKey(ByVal vo As T) As String
                Dim results As New List(Of String)
                For Each info As PropertyInfo In groupingInfos
                    Dim value As Object = info.GetValue(vo, Nothing)
                    results.Add(If(value Is Nothing, "null", value.ToString))
                Next
                Return Join(results.ToArray, "_:_")
            End Function

            ''' <summary>
            ''' グループ化項目の値で値objectを作成する
            ''' </summary>
            ''' <param name="vo">元となる値object</param>
            ''' <returns>グループ化項目の値で作成した値object</returns>
            ''' <remarks></remarks>
            Public Function MakeValue(ByVal vo As T) As T
                Dim aType As Type = GetType(T)
                Dim result As T = CType(Activator.CreateInstance(aType), T)
                For Each info As PropertyInfo In groupingInfos
                    info.SetValue(result, info.GetValue(vo, Nothing), Nothing)
                Next
                Return result
            End Function

            Private Function DetectEffectiveValues(ByVal info As PropertyInfo, ByVal breakdownVos As IEnumerable(Of T)) As List(Of Object)

                Dim values As New List(Of Object)
                For Each bdVo As T In breakdownVos
                    Dim value As Object = info.GetValue(bdVo, Nothing)
                    If value IsNot Nothing Then
                        values.Add(value)
                    End If
                Next
                Return values
            End Function

            Public Sub SetMaxValueTo(ByVal vo As T, ByVal breakdownVos As IEnumerable(Of T))
                For Each info As PropertyInfo In maxInfos

                    Dim values As List(Of Object) = DetectEffectiveValues(info, breakdownVos)
                    If values.Count = 0 Then
                        Continue For
                    End If

                    info.SetValue(vo, GetDetector(info).DetectMaxValue(values), Nothing)
                Next
            End Sub

            Public Sub SetTopValueTo(ByVal vo As T, ByVal breakdownVos As IEnumerable(Of T))
                For Each info As PropertyInfo In topInfos

                    Dim values As List(Of Object) = DetectEffectiveValues(info, breakdownVos)
                    If values.Count = 0 Then
                        Continue For
                    End If

                    info.SetValue(vo, values(0), Nothing)
                Next
            End Sub

            Private Function GetDetector(ByVal info As PropertyInfo) As Detector
                Dim propertyType As Type = TypeUtil.GetTypeIfNullable(info.PropertyType)
                If propertyType Is GetType(String) Then
                    Return New StringDetector
                ElseIf propertyType Is GetType(Int32) Then
                    Return New Int32Detector
                ElseIf propertyType Is GetType(Int64) Then
                    Return New Int64Detector
                End If
                Throw New NotSupportedException(propertyType.Name & " には未対応")
            End Function
        End Class

        Private Interface Detector
            Function DetectMaxValue(ByVal values As IEnumerable(Of Object)) As Object
        End Interface

        Private Class StringDetector : Implements Detector

            Public Function DetectMaxValue(ByVal effectiveValues As IEnumerable(Of Object)) As Object Implements Detector.DetectMaxValue
                Dim maxValue As String = String.Empty
                For Each value As Object In effectiveValues
                    If maxValue.CompareTo(value.ToString) < 0 Then
                        maxValue = value.ToString
                    End If
                Next
                Return maxValue
            End Function
        End Class

        Private Class Int32Detector : Implements Detector

            Public Function DetectMaxValue(ByVal effectiveValues As IEnumerable(Of Object)) As Object Implements Detector.DetectMaxValue
                Dim maxValue As Int32 = Int32.MinValue
                For Each value As Object In effectiveValues
                    maxValue = Math.Max(maxValue, Convert.ToInt32(value))
                Next
                Return maxValue
            End Function
        End Class

        Private Class Int64Detector : Implements Detector

            Public Function DetectMaxValue(ByVal effectiveValues As IEnumerable(Of Object)) As Object Implements Detector.DetectMaxValue
                Dim maxValue As Int64 = Int64.MinValue
                For Each value As Object In effectiveValues
                    maxValue = Math.Max(maxValue, Convert.ToInt64(value))
                Next
                Return maxValue
            End Function
        End Class

        ''' <summary>
        ''' グループ化する項目を指定させるクラス
        ''' </summary>
        ''' <remarks></remarks>
        Private Class LocatorImpl : Implements IVoGroupingLocator

            Private ReadOnly owner As LocatorOwner
            Public Sub New(ByVal owner As LocatorOwner)
                Me.owner = owner
            End Sub

            ''' <summary>
            ''' グループ化する項目を指定する
            ''' </summary>
            ''' <param name="groupingFields">グループ化する項目[]</param>
            ''' <returns>項目のLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function By(ByVal ParamArray groupingFields As Object()) As IVoGroupingLocator Implements IVoGroupingLocator.By
                For Each groupingField As Object In groupingFields
                    owner.MessageGroup(groupingField)
                Next
                Return Me
            End Function

            ''' <summary>
            ''' 最大値を取得する項目を指定する
            ''' </summary>
            ''' <param name="maxValueField">最大値を取得する項目</param>
            ''' <returns>項目のLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function Max(ByVal maxValueField As Object) As IVoGroupingLocator Implements IVoGroupingLocator.Max
                owner.MessageMax(maxValueField)
                Return Me
            End Function

            ''' <summary>
            ''' 先頭値を取得する項目を指定する
            ''' </summary>
            ''' <param name="maxValueField">先頭値を取得する項目</param>
            ''' <returns>項目のLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function Top(ByVal maxValueField As Object) As IVoGroupingLocator Implements IVoGroupingLocator.Top
                owner.MessageTop(maxValueField)
                Return Me
            End Function
        End Class

        Private owner As LocatorOwner

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">グループ化のルール（ラムダ式）</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As VoGroupingRule(Of T).RuleConfigure)
            Me.New(New VoGroupingRule(Of T)(rule))
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">グループ化のルール</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As IVoGroupingRule(Of T))
            owner = New LocatorOwner(rule)
        End Sub

        Private _Result As New List(Of T)
        ''' <summary>グループ化した結果List</summary>
        Public ReadOnly Property Result() As List(Of T)
            Get
                Return _Result
            End Get
        End Property

        Private _ResultBreakdown As New List(Of List(Of T))
        ''' <summary>グループ化した内容の内訳</summary>
        Public ReadOnly Property ResultBreakdown() As List(Of T)()
            Get
                Return _ResultBreakdown.ToArray
            End Get
        End Property

        ''' <summary>
        ''' グループ化して返す
        ''' </summary>
        ''' <param name="dupulicates">重複した要素[]</param>
        ''' <returns>グループ化したList</returns>
        ''' <remarks></remarks>
        Public Function Group(ByVal ParamArray dupulicates As T()) As List(Of T)
            Return Me.Group(DirectCast(dupulicates, IEnumerable(Of T)))
        End Function

        ''' <summary>
        ''' グループ化して返す
        ''' </summary>
        ''' <param name="dupulicateList">重複したList</param>
        ''' <returns>グループ化したList</returns>
        ''' <remarks></remarks>
        Public Function Group(ByVal dupulicateList As IEnumerable(Of T)) As List(Of T)

            _Result.Clear()
            _ResultBreakdown.Clear()

            Dim valueByGroupKey As New Dictionary(Of String, T)
            Dim breakdownsByGroupKey As New Dictionary(Of String, List(Of T))
            For Each vo As T In dupulicateList
                Dim groupKey As String = owner.MakeKey(vo)
                If Not valueByGroupKey.ContainsKey(groupKey) Then
                    valueByGroupKey.Add(groupKey, owner.MakeValue(vo))
                    _Result.Add(valueByGroupKey(groupKey))

                    breakdownsByGroupKey.Add(groupKey, New List(Of T))
                    _ResultBreakdown.Add(breakdownsByGroupKey(groupKey))
                End If
                breakdownsByGroupKey(groupKey).Add(vo)
            Next

            For i As Integer = 0 To _Result.Count - 1
                owner.SetMaxValueTo(_Result(i), _ResultBreakdown(i))
                owner.SetTopValueTo(_Result(i), _ResultBreakdown(i))
            Next

            Return _Result
        End Function
    End Class
End Namespace