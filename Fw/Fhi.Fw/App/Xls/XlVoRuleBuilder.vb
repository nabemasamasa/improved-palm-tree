Imports Fhi.Fw.Util

Namespace App.Xls
    ''' <summary>
    ''' Excel⇔Vo相互変換設定の構築用クラス
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <remarks></remarks>
    Public Class XlVoRuleBuilder(Of T)
        Private Class XlVoSelectorImpl : Implements XlVoSelector
            Private ReadOnly _Builder As XlVoRuleBuilder(Of T)

            Public Sub New(builder As XlVoRuleBuilder(Of T))
                _Builder = builder
            End Sub

            Public Function Use(ByVal field As Object) As XlVoSelector Implements XlVoSelector.Use
                _Builder.AddRule(field, Nothing, Nothing)
                Return Me
            End Function

            Public Function Use(ByVal ParamArray fields As Object()) As XlVoSelector Implements XlVoSelector.Use
                If fields Is Nothing Then
                    _Builder.AddEmptyRule()
                    Return Me
                End If
                For Each field As Object In fields
                    _Builder.AddRule(field)
                Next
                Return Me
            End Function

            Public Function UseWithFunc(ByVal field As Object, Optional ByVal toVoDecorator As Func(Of Object, Object) = Nothing, Optional ByVal toXlsDecorator As Func(Of Object, Object) = Nothing) As XlVoSelector Implements XlVoSelector.UseWithFunc
                If toVoDecorator Is Nothing AndAlso toXlsDecorator Is Nothing Then
                    Throw New ArgumentException("toVoDecoratorもしくはtoXlsDecoratorのいずれかを設定する必要があります")
                End If
                _Builder.AddRule(field, toVoDecorator, toXlsDecorator)
                Return Me
            End Function
        End Class

        Private ReadOnly _Marker As New VoPropertyMarker
        Private ReadOnly _Rules As New List(Of XlVoPropertyRule)

        ''' <summary>相互変換設定用のDelegate</summary>
        ''' <param name="define">ルール設定用のインターフェース</param>
        ''' <param name="vo">Vo</param>
        ''' <returns>ルール設定用のインターフェース</returns>
        ''' <remarks></remarks>
        Public Delegate Function Configure(define As XlVoSelector, vo As T) As XlVoSelector

        ''' <summary>Xls⇔Vo相互変換ルール</summary>
        Friend ReadOnly Property Rules As XlVoPropertyRule()
            Get
                Return _Rules.ToArray
            End Get
        End Property

        Friend Sub New(rule As Configure)
            Dim vo As T = _Marker.CreateMarkedVo(Of T)()
            rule.Invoke(New XlVoSelectorImpl(Me), vo)
        End Sub

        Private Sub AddEmptyRule()
            AddRule(Nothing)
        End Sub

        Private Sub AddRule(ByVal field As Object)
            If field Is Nothing Then
                _Rules.Add(Nothing)
                Return
            End If
            _Rules.Add(New XlVoPropertyRule(_Marker.GetMemberName(field), Nothing, Nothing))
        End Sub

        Private Sub AddRule(ByVal field As Object, ByVal getValueDecolator As Func(Of Object, Object), ByVal setValueDecolator As Func(Of Object, Object))
            If field Is Nothing Then
                _Rules.Add(Nothing)
                Return
            End If
            _Rules.Add(New XlVoPropertyRule(_Marker.GetMemberName(field), getValueDecolator, setValueDecolator))
        End Sub

    End Class
End Namespace