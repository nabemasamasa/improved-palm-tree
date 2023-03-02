Namespace Util
    Public Class NameGetter

        Public Interface Locator
            ''' <summary>
            ''' 対象となるプロパティを指定する
            ''' </summary>
            ''' <param name="values">プロパティ値[]</param>
            ''' <returns>自身</returns>
            ''' <remarks></remarks>
            Function By(ByVal ParamArray values As Object()) As Locator
        End Interface

#Region "Locator impl"
        Private Class LocatorImpl : Implements Locator
            Private ReadOnly owner As NameGetter

            Public Sub New(ByVal owner As NameGetter)
                Me.owner = owner
            End Sub

            Public Function By(ByVal ParamArray values As Object()) As Locator Implements Locator.By
                EzUtil.AssertParameterIsNotEmpty(values, "values")
                For Each val As Object In values
                    owner.ordinalValues.Add(val)
                Next
                Return Me
            End Function
        End Class
#End Region

        Private ReadOnly voMarker As New VoPropertyMarker
        Private ordinalValues As List(Of Object)

        Public Function [Is](ByVal vo As Object) As Locator
            voMarker.Clear()
            voMarker.MarkVo(vo)
            ordinalValues = New List(Of Object)
            _Result = Nothing
            Return New LocatorImpl(Me)
        End Function

        Private _Result As String()
        Public ReadOnly Property Result() As String()
            Get
                If _Result Is Nothing Then
                    Dim aResult As New List(Of String)
                    For Each value As Object In ordinalValues
                        If Not voMarker.Contains(value) Then
                            Throw New InvalidOperationException("voのプロパティを指定すべき")
                        End If
                        aResult.Add(StringUtil.Decamelize(voMarker.GetPropertyInfo(value).Name))
                    Next
                    _Result = aResult.ToArray
                End If
                Return _Result
            End Get
        End Property
    End Class
End Namespace