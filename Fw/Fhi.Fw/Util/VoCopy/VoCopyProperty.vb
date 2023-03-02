Imports System.Reflection

Namespace Util.VoCopy

    ''' <summary>
    ''' 異なるプロパティ間でデータのコピーを行うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class VoCopyProperty(Of Tx, Ty)

        ''' <summary>
        ''' 定義する
        ''' </summary>
        ''' <param name="defineBy">コピーするプロパティの定義</param>
        ''' <param name="x">定義対象その１</param>
        ''' <param name="y">定義対象その２</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Delegate Function Configure(ByVal defineBy As IVoCopyPropertyDefine, ByVal x As Tx, ByVal y As Ty) As IVoCopyPropertyDefine
#Region "Nested classes..."

        Private Class VoCopyPropertyDefine : Implements IVoCopyPropertyDefine
            Private ReadOnly voXMarker As New VoPropertyMarker
            Private ReadOnly voYMarker As New VoPropertyMarker
            Public ReadOnly propertyRelations As New Dictionary(Of PropertyInfo, PropertyInfo)
            Public Sub New(ByVal rule As Configure)
                Dim xType As Type = GetType(Tx)
                Dim yType As Type = GetType(Ty)
                Dim voX As Tx = DirectCast(Activator.CreateInstance(xType), Tx)
                Dim voY As Ty = DirectCast(Activator.CreateInstance(yType), Ty)
                voXMarker.MarkVo(voX)
                voYMarker.MarkVo(voY)

                rule.Invoke(New BinderImpl(Me), voX, voY)
            End Sub

            ''' <summary>
            ''' コピーするプロパティを設定する
            ''' </summary>
            ''' <param name="x">関連づけるプロパティ値</param>
            ''' <param name="y">関連づけるプロパティ値</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Bind(ByVal x As Object, ByVal y As Object) As IVoCopyPropertyDefine Implements IVoCopyPropertyDefine.Bind
                Dim xPropety As PropertyInfo
                Dim yPropety As PropertyInfo
                If voXMarker.Contains(x) AndAlso voYMarker.Contains(y) Then
                    xPropety = voXMarker.GetPropertyInfo(x)
                    yPropety = voYMarker.GetPropertyInfo(y)
                Else
                    xPropety = voXMarker.GetPropertyInfo(y)
                    yPropety = voYMarker.GetPropertyInfo(x)
                End If

                propertyRelations.Add(xPropety, yPropety)
                propertyRelations.Add(yPropety, xPropety)
                Return Me
            End Function

        End Class

        Private Class BinderImpl : Implements IVoCopyPropertyDefine

            Private ReadOnly builder As VoCopyPropertyDefine

            Public Sub New(ByVal builder As VoCopyPropertyDefine)
                Me.builder = builder
            End Sub

            Public Function Bind(ByVal x As Object, ByVal y As Object) As IVoCopyPropertyDefine Implements IVoCopyPropertyDefine.Bind
                builder.Bind(x, y)
                Return Me
            End Function
        End Class
#End Region

        Private ReadOnly builder As VoCopyPropertyDefine

        Public Sub New(ByVal aConfigure As Configure)
            builder = New VoCopyPropertyDefine(aConfigure)
        End Sub

        ''' <summary>
        ''' 定義したプロパティ同士の値をコピーする
        ''' </summary>
        ''' <param name="src">コピー元</param>
        ''' <param name="dest">コピー先</param>
        ''' <remarks></remarks>
        Public Sub CopyProperties(ByVal src As Tx, ByVal dest As Ty)
            Dim xType As Type = GetType(Tx)
            For Each info As PropertyInfo In xType.GetProperties
                If Not builder.propertyRelations.ContainsKey(info) Then
                    Continue For
                End If
                Dim value As Object = info.GetValue(src, Nothing)
                builder.propertyRelations(info).SetValue(dest, value, Nothing)
            Next
        End Sub

        ''' <summary>
        ''' 定義したプロパティ同士の値をコピーする
        ''' </summary>
        ''' <param name="src">コピー元</param>
        ''' <param name="dest">コピー先</param>
        ''' <remarks></remarks>
        Public Sub CopyProperties(ByVal src As Ty, ByVal dest As Tx)
            Dim yType As Type = GetType(Ty)
            For Each info As PropertyInfo In yType.GetProperties
                If Not builder.propertyRelations.ContainsKey(info) Then
                    Continue For
                End If
                Dim value As Object = info.GetValue(src, Nothing)
                builder.propertyRelations(info).SetValue(dest, value, Nothing)
            Next

        End Sub
    End Class
End Namespace