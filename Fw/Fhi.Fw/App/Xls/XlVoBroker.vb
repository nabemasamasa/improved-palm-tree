Imports Fhi.Fw.Domain

Namespace App.Xls
    ''' <summary>
    ''' Excel⇔Voの相互変換を行う為のクラス
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <remarks></remarks>
    Friend Class XlVoBroker(Of T As New)
        Private ReadOnly _rules As XlVoPropertyRule()

        Friend Sub New(rules As IEnumerable(Of XlVoPropertyRule))
            _rules = rules.ToArray
        End Sub

        Friend Sub New(define As XlVoRuleBuilder(Of T).Configure)
            _rules = New XlVoRuleBuilder(Of T)(define).Rules
        End Sub

        ''' <summary>
        ''' 行情報の配列からひとつのVoへ変換する
        ''' </summary>
        ''' <param name="rowData">行情報[]</param>
        ''' <returns>Vo</returns>
        ''' <remarks></remarks>
        Public Function CreateVo(rowData As Object()) As T
            Dim vo As New T
            For i As Integer = 0 To _rules.Count - 1
                If rowData.Count <= i Then
                    Return vo
                End If
                If _rules(i) Is Nothing Then
                    Continue For
                End If
                Dim rule As XlVoPropertyRule = _rules(i)
                Dim value As Object = VoUtil.ResolveValueContainsPVO(rowData(i), DetectType(rule.PropertyName))
                If rule.ToVoDecolator IsNot Nothing Then
                    value = rule.ToVoDecolator(value)
                End If
                VoUtil.SetValue(vo, rule.PropertyName, value)
            Next
            Return vo
        End Function

        Private Function DetectType(propertyName As String) As Type
            Return GetType(T).GetProperties.FirstOrDefault(Function(info) info.Name.Equals(propertyName)).PropertyType
        End Function

        ''' <summary>
        ''' Vo[]をExcel出力用の二次元配列に変換する
        ''' </summary>
        ''' <param name="vos">Vo[]</param>
        ''' <returns>二次元配列</returns>
        ''' <remarks></remarks>
        Public Function ConvertToValues(vos As IEnumerable(Of T)) As Array
            Dim data As Array = Array.CreateInstance(GetType(Object), vos.Count, _rules.Count)
            For row As Integer = 0 To vos.Count - 1
                For col As Integer = 0 To _rules.Count - 1
                    If _rules(col) Is Nothing Then
                        Continue For
                    End If
                    Dim rule As XlVoPropertyRule = _rules(col)
                    Dim value As Object = VoUtil.GetValue(vos(row), _rules(col).PropertyName)
                    If value IsNot Nothing AndAlso TypeUtil.IsTypeValueObjectOrSubClass(value.GetType) Then
                        value = DirectCast(value, PrimitiveValueObject).Value
                    End If
                    If rule.ToXlsDecolator IsNot Nothing Then
                        value = rule.ToXlsDecolator(value)
                    End If
                    data.SetValue(value, row, col)
                Next
            Next
            Return data
        End Function

    End Class
End Namespace