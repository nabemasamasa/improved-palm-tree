Imports System.Xml
Imports System.Reflection

''' <summary>
''' Xml操作用のユーティリティ
''' </summary>
''' <remarks>
''' 【プロパティ値⇔XML要素 の相互関係】<![CDATA[
'''                          値型                 Nullableの値型      Null許容型          Null許容型
''' 要素                     Integer              Integer?            String              Object(String以外)
''' なし                     既定値=0             Nothing             Nothing             Nothing
''' <Xyz />                  既定値=0             Nothing             Nothing             Nothing
''' <Xyz></Xyz>              既定値=0             Nothing             空文字              空インスタンス
''' <Xyz>123</Xyz>           123                  123                 "123"               空インスタンス(不正値)
''' <Xyz>あい</Xyz>          既定値=0             Nothing(不正値)     "あい"              空インスタンス(不正値)
''' 
'''                          コレクション型       コレクション型      コレクション型      コレクション型
''' なし                     コレクションそのものがNothing
''' <Xyz />                  1要素で中身 既定値=0 1要素で中身 Nothing 1要素で中身 Nothing 1要素で中身 Nothing
''' <Xyz></Xyz>              1要素で中身 既定値=0 1要素で中身 Nothing 1要素で中身 空文字  1要素で中身 空インスタンス
''' <Xyz>123</Xyz>           1要素で中身 123      1要素で中身 123     1要素で中身 "123"   1要素で中身 空インスタンス(不正値)
''' <Xyz Empty="True" />     長さ0のコレクション  長さ0のコレクション 長さ0のコレクション 長さ0のコレクション
''' <Xyz Empty="True"></Xyz> 長さ0のコレクション  長さ0のコレクション 長さ0のコレクション 長さ0のコレクション
''' ]]>
''' </remarks>
Public Class XmlUtil

    Public Class Attribute
        ''' <summary>長さ0配列を表す属性</summary>
        Public Const ARRAY_EMPTY As String = "Empty"
    End Class

    ''' <summary>
    ''' ノードの内容をVo値に埋め込む
    ''' </summary>
    ''' <param name="node">ノード</param>
    ''' <param name="vo">Vo値</param>
    ''' <remarks></remarks>
    Public Shared Sub PopulateNodeToVo(ByVal node As XmlNode, ByVal vo As Object)
        PerformPopulateNodeToVo(node, vo)
    End Sub

    ''' <summary>
    ''' ノードの内容をVo値に埋め込む
    ''' </summary>
    ''' <param name="node">ノード</param>
    ''' <param name="vo">Vo値</param>
    ''' <remarks></remarks>
    Private Shared Sub PerformPopulateNodeToVo(ByVal node As XmlNode, ByVal vo As Object)

        If vo Is Nothing Then
            Return
        End If

        Dim infoByName As New Dictionary(Of String, PropertyInfo)
        For Each info As PropertyInfo In vo.GetType.GetProperties
            infoByName.Add(info.Name, info)
        Next

        Dim collectionByName As New Dictionary(Of String, List(Of Object))
        For Each child As XmlNode In node.ChildNodes
            If Not infoByName.ContainsKey(child.Name) Then
                Continue For
            End If

            Dim info As PropertyInfo = infoByName(child.Name)
            If info.GetSetMethod Is Nothing Then
                Continue For
            End If

            If TypeUtil.IsTypeCollection(info.PropertyType) Then
                Dim elementType As Type = TypeUtil.DetectElementType(info.PropertyType)
                If TypeUtil.IsTypeCollection(elementType) Then
                    Throw New NotSupportedException("コレクションの入れ子は未対応")
                End If

                If Not collectionByName.ContainsKey(child.Name) Then
                    collectionByName.Add(child.Name, New List(Of Object))
                End If

                Dim attrLength As XmlAttribute = GetAttributeIgnoreCase(child, Attribute.ARRAY_EMPTY)
                If attrLength IsNot Nothing AndAlso EzUtil.IsTrue(attrLength.Value) Then
                    Continue For
                End If

                If TypeUtil.IsTypeImmutable(elementType) Then
                    collectionByName(child.Name).Add(VoUtil.ResolveValue(GetValue(child), elementType))

                Else
                    Dim newInstance As Object = Nothing
                    If Not IsEmptyElementTag(child) Then
                        newInstance = VoUtil.NewInstance(elementType)
                        PerformPopulateNodeToVo(child, newInstance)
                    End If
                    collectionByName(child.Name).Add(newInstance)
                End If

            ElseIf TypeUtil.IsTypeImmutable(TypeUtil.GetTypeIfNullable(info.PropertyType)) Then

                info.SetValue(vo, VoUtil.ResolveValue(GetValue(child), info.PropertyType), Nothing)
            Else

                If Not IsEmptyElementTag(child) Then
                    Dim newInstance As Object = VoUtil.NewInstance(info.PropertyType)
                    info.SetValue(vo, newInstance, Nothing)
                    PerformPopulateNodeToVo(child, newInstance)
                End If
            End If
        Next

        For Each pair As KeyValuePair(Of String, List(Of Object)) In collectionByName
            Dim name As String = pair.Key
            Dim collection As List(Of Object) = pair.Value
            Dim info As PropertyInfo = infoByName(name)
            Dim resolvedCollection As Object = VoUtil.NewCollectionInstanceWithInitialElement(info.PropertyType, collection.Count, collection)
            info.SetValue(vo, resolvedCollection, Nothing)
        Next
    End Sub

    ''' <summary>
    ''' 属性を返す（大/小文字同一視）
    ''' </summary>
    ''' <param name="node">ノード</param>
    ''' <param name="attributeName">属性名</param>
    ''' <returns>属性</returns>
    ''' <remarks></remarks>
    Public Shared Function GetAttributeIgnoreCase(ByVal node As XmlNode, ByVal attributeName As String) As XmlAttribute
        For Each attr As XmlAttribute In node.Attributes
            If attr.Name.Equals(attributeName, StringComparison.OrdinalIgnoreCase) Then
                Return attr
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' ノードの「内容」を返す
    ''' </summary>
    ''' <param name="node">ノード</param>
    ''' <returns>値</returns>
    ''' <remarks></remarks>
    Public Shared Function GetValue(ByVal node As XmlNode) As String
        If IsEmptyElementTag(node) Then
            Return Nothing
        End If
        Return node.InnerText
    End Function

    ''' <summary>
    ''' ノードが空エレメントタグ（空要素タグ）かを返す
    ''' </summary>
    ''' <param name="node">ノード</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmptyElementTag(ByVal node As XmlNode) As Boolean

        Return node.OuterXml.EndsWith("/>")
    End Function

    ''' <summary>
    ''' Xmlドキュメント直下に要素を作成しVo値を埋め込む
    ''' </summary>
    ''' <param name="vo">Vo値</param>
    ''' <param name="doc">Xmlドキュメント</param>
    ''' <param name="isEmptyElementIfNull">Null値を空要素で出力する場合、true. 出力しない場合、false</param>
    ''' <remarks></remarks>
    Public Shared Sub PopulateVoToDoc(ByVal vo As Object, ByVal doc As XmlDocument, Optional ByVal isEmptyElementIfNull As Boolean = False)
        If vo Is Nothing Then
            Return
        End If
        Dim element As XmlElement = doc.CreateElement(vo.GetType.Name)
        doc.AppendChild(element)
        PopulateVoToDoc(vo, element, isEmptyElementIfNull)
    End Sub

    ''' <summary>
    ''' XmlノードにVo値を埋め込む
    ''' </summary>
    ''' <param name="vo">Vo値</param>
    ''' <param name="parentNode">Xmlドキュメント</param>
    ''' <param name="isEmptyElementIfNull">Null値を空要素で出力する場合、true. 出力しない場合、false</param>
    ''' <remarks></remarks>
    Public Shared Sub PopulateVoToDoc(ByVal vo As Object, ByVal parentNode As XmlNode, Optional ByVal isEmptyElementIfNull As Boolean = False)
        If vo Is Nothing Then
            Return
        End If
        PerformPopulateVoToNode(vo, parentNode.OwnerDocument, parentNode, isEmptyElementIfNull)
    End Sub

    ''' <summary>
    ''' Nodeに値を埋め込む
    ''' </summary>
    ''' <param name="value">値obj</param>
    ''' <param name="doc"></param>
    ''' <param name="node">値objの属性をsetするNode</param>
    ''' <param name="isEmptyElementIfNull">Null値を空要素で出力する場合、true. 出力しない場合、false</param>
    ''' <remarks></remarks>
    Private Shared Sub PerformPopulateVoToNode(ByVal value As Object, ByVal doc As XmlDocument, ByVal node As XmlNode, ByVal isEmptyElementIfNull As Boolean)

        If value Is Nothing Then
            Return
        End If
        If TypeUtil.IsTypeImmutable(value.GetType) Then
            If TypeOf value Is DateTime Then
                node.InnerText = DirectCast(value, DateTime).ToString("yyyy/MM/dd HH:mm:ss.fff")
            Else
                node.InnerText = StringUtil.ToString(value)
            End If
            Return
        End If

        For Each info As PropertyInfo In value.GetType.GetProperties

            If info.GetGetMethod Is Nothing Then
                Continue For
            End If

            Dim propertyValue As Object = info.GetValue(value, Nothing)
            Dim propertyType As Type = TypeUtil.GetTypeIfNullable(info.PropertyType)

            Dim isTypeCollection As Boolean = TypeUtil.IsTypeCollection(info.PropertyType)
            Dim values As Object()
            If isTypeCollection Then
                If propertyValue Is Nothing Then
                    Continue For
                    ' コレクション値がNullなら、その属性は出力しない
                End If
                values = VoUtil.ConvObjectToArray(propertyValue)

                If values.Length = 0 Then
                    Dim element As XmlElement = doc.CreateElement(info.Name)
                    element.Attributes.Append(MakeAttribute(doc, Attribute.ARRAY_EMPTY, Boolean.TrueString))
                    node.AppendChild(element)
                    Continue For
                End If
            Else
                values = New Object() {propertyValue}
            End If

            For Each elementVo As Object In values
                If elementVo Is Nothing AndAlso Not isEmptyElementIfNull AndAlso Not isTypeCollection Then
                    Continue For
                End If
                Dim elementNode As XmlElement = doc.CreateElement(info.Name)
                node.AppendChild(elementNode)
                PerformPopulateVoToNode(elementVo, doc, elementNode, isEmptyElementIfNull)
            Next
        Next

        ' valueが Vo で中身無ければ <Vo></Vo> にする
        If node.ChildNodes.Count = 0 Then
            node.InnerText = ""
        End If
    End Sub

    ''' <summary>
    ''' 属性を作成する
    ''' </summary>
    ''' <param name="doc">XMLドキュメント</param>
    ''' <param name="name">属性名</param>
    ''' <param name="value">値</param>
    ''' <returns>属性</returns>
    ''' <remarks></remarks>
    Private Shared Function MakeAttribute(ByVal doc As XmlDocument, ByVal name As String, ByVal value As String) As XmlAttribute

        Dim attr As XmlAttribute = doc.CreateAttribute(name)
        attr.Value = value
        Return attr
    End Function
End Class
