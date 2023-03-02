Namespace Util.Csv
    Public Enum CsvSeparator
        COMMA = 0
        TAB
    End Enum

    ''' <summary>
    ''' カンマ区切り・タブ区切りのテキスト文字列とVoとの相互変換（連携？）を担うクラス
    ''' </summary>
    ''' <typeparam name="T">Voの型</typeparam>
    ''' <remarks></remarks>
    Public Class CsvVoBroker(Of T)

        Private Const DQ As Char = """"c
        Private Shared ReadOnly SEPARATORS As Char() = {","c, ControlChars.Tab}

        Private ReadOnly builder As CsvRuleBuilder(Of T)
        Private _csvStrings As String()
        Private _vos As T()
        Private _separator As Char
        Private _csvSeparator As CsvSeparator

        ''' <summary>列タイトルを出力する場合、true</summary>
        Public IsOutputTitle As Boolean = False

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">区切り列順設定ルール（ラムダ式）</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As CsvRule(Of T).RuleConfigure)
            Me.New(New CsvRule(Of T)(rule))
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">区切り列順設定ルール（ラムダ式）</param>
        ''' <param name="separator">区切り文字</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As CsvRule(Of T).RuleConfigure, ByVal separator As CsvSeparator)

            Me.New(New CsvRule(Of T)(rule), separator)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">区切り列順設定ルール</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As ICsvRule(Of T))
            Me.New(rule, CsvSeparator.COMMA)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">区切り列順設定ルール</param>
        ''' <param name="separator">区切り文字</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As ICsvRule(Of T), ByVal separator As CsvSeparator)

            builder = New CsvRuleBuilder(Of T)(rule)
            Me.Separator = separator
        End Sub

        ''' <summary>CSVファイル設定した各行のVo</summary>
        Public Property Vos() As T()
            Get
                Return _vos
            End Get
            Set(ByVal value As T())
                _vos = value
                ApplyVosToCsvStrings()
            End Set
        End Property

        ''' <summary>CSVファイルの各行</summary>
        Public Property CsvStrings() As String()
            Get
                Return _csvStrings
            End Get
            Set(ByVal value As String())
                _csvStrings = value
                ApplyCsvStringsToVos()
            End Set
        End Property

        ''' <summary>区切り文字</summary>
        Public Property Separator() As CsvSeparator
            Get
                Return _csvSeparator
            End Get
            Set(ByVal value As CsvSeparator)
                _csvSeparator = value
                _separator = SEPARATORS(value)
            End Set
        End Property

        ''' <summary>
        ''' Voの値でCSV文字列を変更する
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ApplyVosToCsvStrings()
            Dim result As New List(Of String)

            Dim properties As CsvRuleProperty() = builder.ResultProperties
            If IsOutputTitle Then
                Dim lineValues As String() = MakeTitleValues(properties)
                result.Add(FileUtil.OutputCsvRow(lineValues, _separator))
            End If

            For Each vo As T In _vos
                Dim csvValues As List(Of String) = PerformApplyVoToCsvString(vo, properties)
                result.Add(FileUtil.OutputCsvRow(csvValues, _separator))
            Next

            _csvStrings = result.ToArray
        End Sub

        ''' <summary>
        ''' 列名の一覧を作成する
        ''' </summary>
        ''' <param name="properties">列定義[]</param>
        ''' <returns>列名[]</returns>
        ''' <remarks></remarks>
        Private Function MakeTitleValues(ByVal properties As CsvRuleProperty()) As String()
            Return RecurMakeTitleValues(properties, Nothing)
        End Function

        Private Function RecurMakeTitleValues(ByVal properties As CsvRuleProperty(), ByVal prefix As String) As String()

            Dim lineValues As New List(Of String)
            For Each aProperty As CsvRuleProperty In properties
                If aProperty.IsTitleOnly Then
                    lineValues.Add(aProperty.Title)
                    Continue For
                End If
                Dim isTypeCollection As Boolean = TypeUtil.IsTypeCollection(aProperty.Info.PropertyType)
                If aProperty.IsGroup Then
                    Dim csvRuleProperties As CsvRuleProperty() = aProperty.Builder.ResultProperties
                    If isTypeCollection Then
                        For i As Integer = 0 To aProperty.Repeat - 1
                            lineValues.AddRange(RecurMakeTitleValues(csvRuleProperties, MakeTitleName(prefix, aProperty, i)))
                        Next
                    Else
                        lineValues.AddRange(RecurMakeTitleValues(csvRuleProperties, MakeTitleName(prefix, aProperty)))
                    End If
                ElseIf isTypeCollection Then
                    For i As Integer = 0 To aProperty.Repeat - 1
                        lineValues.Add(MakeTitleName(prefix, aProperty, i))
                    Next
                Else
                    lineValues.Add(MakeTitleName(prefix, aProperty))
                End If
            Next
            Return lineValues.ToArray
        End Function

        ''' <summary>
        ''' 列設定情報の列名を作成する
        ''' </summary>
        ''' <param name="prefix">接頭文字</param>
        ''' <param name="aProperty">CSV列設定情報</param>
        ''' <returns>列名</returns>
        ''' <remarks></remarks>
        Private Function MakeTitleName(ByVal prefix As String, ByVal aProperty As CsvRuleProperty) As String
            If StringUtil.IsEmpty(prefix) Then
                Return aProperty.Info.Name
            End If
            Return String.Format("{0}.{1}", prefix, aProperty.Info.Name)
        End Function

        ''' <summary>
        ''' 列設定情報の列名を作成する
        ''' </summary>
        ''' <param name="prefix">接頭文字</param>
        ''' <param name="aProperty">CSV列設定情報</param>
        ''' <param name="i">添え字</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function MakeTitleName(ByVal prefix As String, ByVal aProperty As CsvRuleProperty, ByVal i As Integer) As String
            Return String.Format("{0}({1})", MakeTitleName(prefix, aProperty), i)
        End Function

        Private Shared Function PerformApplyVoToCsvString(ByVal vo As Object, ByVal properties As ICollection(Of CsvRuleProperty)) As List(Of String)
            Dim line As New List(Of String)
            For Each aProperty As CsvRuleProperty In properties
                If aProperty.IsTitleOnly Then
                    line.Add(String.Empty)
                    Continue For
                End If
                If TypeUtil.IsTypeCollection(aProperty.Info.PropertyType) Then
                    Dim value As Object = aProperty.Info.GetValue(vo, Nothing)
                    Dim valueList As New List(Of Object)
                    If value IsNot Nothing Then
                        valueList.AddRange(VoUtil.ConvObjectToCollection(value))
                    End If
                    Dim emptyValue As Object = If(valueList.Count < aProperty.Repeat, VoUtil.NewElementInstanceFromCollectionType(aProperty.Info.PropertyType), Nothing)

                    If aProperty.IsGroup Then
                        Dim csvRuleProperties As CsvRuleProperty() = aProperty.Builder.ResultProperties
                        For i As Integer = 0 To aProperty.Repeat - 1
                            line.AddRange(PerformApplyVoToCsvString(If(i < valueList.Count, valueList(i), emptyValue), csvRuleProperties))
                        Next
                    Else
                        For i As Integer = 0 To aProperty.Repeat - 1
                            line.Add(StringUtil.ToString(If(i < valueList.Count, valueList(i), emptyValue)))
                        Next
                    End If

                ElseIf aProperty.IsGroup Then
                    Dim csvRuleProperties As CsvRuleProperty() = aProperty.Builder.ResultProperties
                    line.AddRange(PerformApplyVoToCsvString(aProperty.Info.GetValue(vo, Nothing), csvRuleProperties))

                Else
                    Dim value As Object = aProperty.Info.GetValue(vo, Nothing)
                    If aProperty.ToCsvDecorator IsNot Nothing Then
                        line.Add(aProperty.ToCsvDecorator(value))
                    Else
                        line.Add(StringUtil.ToString(value))
                    End If

                End If
            Next
            Return line
        End Function

        ''' <summary>
        ''' CSV文字列の値でVoを変更する
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ApplyCsvStringsToVos()
            Dim result As New List(Of T)
            Dim properties As CsvRuleProperty() = builder.ResultProperties
            For Each csvString As String In _csvStrings
                Dim vo As T = VoUtil.NewInstance(Of T)()
                result.Add(vo)
                Dim values As String() = StringUtil.SplitForEnclosedDQ(csvString, _separator)
                PerformApplyCsvStringToVo(vo, values, 0, properties)
            Next
            _vos = result.ToArray
        End Sub

        Private Shared Sub PerformApplyCsvStringToVo(ByVal vo As Object, ByVal csvValues As String(), ByVal csvStartIndex As Integer, ByVal properties As CsvRuleProperty())
            Dim csvIndex As Integer = csvStartIndex
            For propertyIndex As Integer = 0 To properties.Length - 1

                If csvValues.Length <= csvIndex Then
                    Return
                End If

                Dim aProperty As CsvRuleProperty = properties(propertyIndex)
                If aProperty.IsGroup Then

                    Dim originValue As Object
                    If TypeUtil.IsTypeCollection(aProperty.Info.PropertyType) Then
                        originValue = VoUtil.NewCollectionInstance(aProperty.Info.PropertyType, aProperty.Repeat)
                    Else
                        originValue = VoUtil.NewInstance(aProperty.Info.PropertyType)
                    End If
                    aProperty.Info.SetValue(vo, originValue, Nothing)

                    If Not TypeUtil.IsArrayOrCollection(originValue) Then
                        PerformApplyCsvStringToVo(originValue, csvValues, csvIndex, aProperty.Builder.ResultProperties)
                        csvIndex += aProperty.Builder.ResultProperties.Length
                        Continue For
                    End If

                    Dim valueCollection As ICollection(Of Object) = VoUtil.ConvObjectToCollection(originValue)

                    Dim index As Integer = 0
                    For Each value As Object In valueCollection
                        PerformApplyCsvStringToVo(value, csvValues, csvIndex + index * aProperty.Builder.ResultProperties.Length, aProperty.Builder.ResultProperties)
                        index += 1
                    Next
                    csvIndex += index * aProperty.Builder.ResultProperties.Length
                ElseIf aProperty.IsTitleOnly Then
                    csvIndex += 1
                ElseIf TypeUtil.IsTypeCollection(aProperty.Info.PropertyType) Then
                    Dim elementType As Type = TypeUtil.DetectElementType(aProperty.Info.PropertyType)
                    Dim resolvedValues As New List(Of Object)
                    For j As Integer = 0 To aProperty.Repeat - 1
                        If csvValues.Length <= csvIndex + j Then
                            Continue For
                        End If
                        resolvedValues.Add(ResolveValue(csvValues(csvIndex + j), elementType))
                    Next

                    Dim collectionValue As Object = VoUtil.NewCollectionInstanceWithInitialElement(aProperty.Info.PropertyType, aProperty.Repeat, resolvedValues)
                    aProperty.Info.SetValue(vo, collectionValue, Nothing)
                    csvIndex += aProperty.Repeat

                Else
                    Dim propertyType As Type = TypeUtil.GetTypeIfNullable(aProperty.Info.PropertyType)
                    Dim value As Object
                    If aProperty.ToVoDecorator IsNot Nothing Then
                        value = aProperty.ToVoDecorator(csvValues(csvIndex))
                    Else
                        value = ResolveValue(csvValues(csvIndex), propertyType)
                    End If

                    aProperty.Info.SetValue(vo, value, Nothing)
                    csvIndex += 1
                End If
            Next
        End Sub

        ''' <summary>
        ''' 値を型に合わせて変換する
        ''' </summary>
        ''' <param name="value">値</param>
        ''' <param name="propertyType">型Type</param>
        ''' <returns>変換後の値</returns>
        ''' <remarks></remarks>
        Private Shared Function ResolveValue(ByVal value As String, ByVal propertyType As Type) As Object

            If value Is Nothing Then
                Return Nothing
            End If

            Return VoUtil.ResolveValue(value, propertyType)
        End Function

    End Class
End Namespace