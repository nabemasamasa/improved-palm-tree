Imports System.Reflection

Namespace Util

    ''' <summary>
    ''' デキャメライズしたプロパティ名で、プロパティアクセスを実現させるクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DecamelizedPropertyAccessor

        Private Enum PropertyType
            GETTER_AND_SETTER
            GETTER
            SETTER
        End Enum

        Private propertyInfoByDecamelizeNameByType As New Dictionary(Of Type, Dictionary(Of String, PropertyInfo))

        ''' <summary>
        ''' デキャメライズしたプロパティ名でPropertyInfo(Getter＆Setter)を返す
        ''' </summary>
        ''' <param name="recordVo"></param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPropertyInfo(recordVo As Object, decamelizeName As String) As PropertyInfo
            Return GetPropertyInfoByDecamelizeName(recordVo, PropertyType.GETTER_AND_SETTER).Item(decamelizeName)
        End Function

        ''' <summary>
        ''' デキャメライズしたプロパティ名でPropertyInfo(Getter)を返す
        ''' </summary>
        ''' <param name="recordVo"></param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetGetPropertyInfo(recordVo As Object, decamelizeName As String) As PropertyInfo
            Return GetPropertyInfoByDecamelizeName(recordVo, PropertyType.GETTER).Item(decamelizeName)
        End Function

        ''' <summary>
        ''' デキャメライズしたプロパティ名でPropertyInfo(Setter)を返す
        ''' </summary>
        ''' <param name="recordVo"></param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSetPropertyInfo(recordVo As Object, decamelizeName As String) As PropertyInfo
            Return GetPropertyInfoByDecamelizeName(recordVo, PropertyType.SETTER).Item(decamelizeName)
        End Function

        ''' <summary>
        ''' プロパティ名と同名のプロパティ(Getter＆Setter)を持つかを返す
        ''' </summary>
        ''' <param name="recordVo">値</param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function Contains(recordVo As Object, decamelizeName As String) As Boolean
            Return GetPropertyInfoByDecamelizeName(recordVo, PropertyType.GETTER_AND_SETTER).ContainsKey(decamelizeName)
        End Function

        ''' <summary>
        ''' プロパティ名と同名のGetプロパティを持つかを返す
        ''' </summary>
        ''' <param name="recordVo">値</param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function ContainsGetter(recordVo As Object, decamelizeName As String) As Boolean
            Return GetPropertyInfoByDecamelizeName(recordVo, PropertyType.GETTER).ContainsKey(decamelizeName)
        End Function

        ''' <summary>
        ''' プロパティ名と同名のSetプロパティを持つかを返す
        ''' </summary>
        ''' <param name="recordVo">値</param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function ContainsSetter(recordVo As Object, decamelizeName As String) As Boolean
            Return GetPropertyInfoByDecamelizeName(recordVo, PropertyType.SETTER).ContainsKey(decamelizeName)
        End Function

        ''' <summary>
        ''' 値を取得する
        ''' </summary>
        ''' <param name="vo">ValueObject</param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <param name="index">index</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function GetValue(vo As Object, decamelizeName As String, ParamArray index As Object()) As Object
            Return GetGetPropertyInfo(vo, decamelizeName).GetValue(vo, index)
        End Function

        ''' <summary>
        ''' 値を設定する
        ''' </summary>
        ''' <param name="vo">ValueObject</param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <param name="value">設定値</param>
        ''' <remarks></remarks>
        Public Sub SetValue(vo As Object, decamelizeName As String, value As Object)
            SetValue(vo, decamelizeName, value, Nothing)
        End Sub

        ''' <summary>
        ''' 値を設定する
        ''' </summary>
        ''' <param name="vo">ValueObject</param>
        ''' <param name="decamelizeName">デキャメライズしたプロパティ名</param>
        ''' <param name="value">設定値</param>
        ''' <param name="index">index</param>
        ''' <remarks></remarks>
        Public Sub SetValue(vo As Object, decamelizeName As String, value As Object, index As Object())
            Dim aProperty As PropertyInfo = GetSetPropertyInfo(vo, decamelizeName)
            aProperty.SetValue(vo, VoUtil.ResolveValue(value, aProperty.PropertyType, True), index)
        End Sub

        Private Function GetPropertyInfoByDecamelizeName(vo As Object, propertyType As PropertyType) As Dictionary(Of String, PropertyInfo)
            EzUtil.AssertParameterIsNotNull(vo, "vo")
            Dim aType As Type = vo.GetType
            If Not propertyInfoByDecamelizeNameByType.ContainsKey(aType) Then
                Dim propertyInfos As Dictionary(Of String, PropertyInfo) = New Dictionary(Of String, PropertyInfo)
                propertyInfoByDecamelizeNameByType.Add(aType, propertyInfos)
                For Each info As PropertyInfo In aType.GetProperties
                    Select Case propertyType
                        Case PropertyType.GETTER_AND_SETTER
                            If info.GetGetMethod Is Nothing OrElse info.GetSetMethod Is Nothing Then
                                Continue For
                            End If
                        Case PropertyType.GETTER
                            If info.GetGetMethod Is Nothing Then
                                Continue For
                            End If
                        Case PropertyType.SETTER
                            If info.GetSetMethod Is Nothing Then
                                Continue For
                            End If
                    End Select
                    propertyInfos.Add(StringUtil.Decamelize(info.Name), info)
                Next
            End If
            Return propertyInfoByDecamelizeNameByType(aType)
        End Function

        ''' <summary>
        ''' デキャメライズしたプロパティ名(Getter＆Setter)一覧を返す
        ''' </summary>
        ''' <param name="vo">Vo</param>
        ''' <returns>デキャメライズしたプロパティ名[]</returns>
        ''' <remarks></remarks>
        Public Function GetDecamelizePropertyNames(vo As Object) As String()
            Dim result As New List(Of String)(GetPropertyInfoByDecamelizeName(vo, PropertyType.GETTER_AND_SETTER).Keys)
            Return result.ToArray
        End Function

        ''' <summary>
        ''' デキャメライズしたプロパティ名(Getter)一覧を返す
        ''' </summary>
        ''' <param name="vo">Vo</param>
        ''' <returns>デキャメライズしたプロパティ名[]</returns>
        ''' <remarks></remarks>
        Public Function GetDecamelizeGetPropertyNames(vo As Object) As String()
            Dim result As New List(Of String)(GetPropertyInfoByDecamelizeName(vo, PropertyType.GETTER).Keys)
            Return result.ToArray
        End Function

        ''' <summary>
        ''' デキャメライズしたプロパティ名(Setter)一覧を返す
        ''' </summary>
        ''' <param name="vo">Vo</param>
        ''' <returns>デキャメライズしたプロパティ名[]</returns>
        ''' <remarks></remarks>
        Public Function GetDecamelizeSetPropertyNames(vo As Object) As String()
            Dim result As New List(Of String)(GetPropertyInfoByDecamelizeName(vo, PropertyType.SETTER).Keys)
            Return result.ToArray
        End Function

    End Class
End Namespace