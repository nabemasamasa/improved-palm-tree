Namespace Util.Csv
    ''' <summary>
    ''' Csvの書き込みを担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CsvWriter
        Private Const DQ As String = """"
        Private Const DQDQ As String = DQ & DQ

        Private ReadOnly fileName As String
        Private ReadOnly csvLines As New List(Of String)
        Private ReadOnly currentColumnValues As New List(Of String)

        ' FIXME カンマをデフォルトにすべき
        ''' <summary>区切り文字 ※初期値はTAB</summary>
        Public Separator As String = vbTab

        Public Sub New(ByVal fileName As String)
            Me.fileName = fileName
        End Sub

        ''' <summary>
        ''' インスタンスに対象文字列を追加する
        ''' </summary>
        ''' <param name="value">対象文字列</param>
        ''' <remarks></remarks>
        Public Sub Append(ByVal value As String)
            If value Is Nothing Then
                currentColumnValues.Add(value)
            Else
                currentColumnValues.Add(DQ & value.Replace(DQ, DQDQ) & DQ)
            End If
        End Sub

        ''' <summary>
        ''' インスタンスにデータを1行追加する
        ''' </summary>
        ''' <param name="values">1行のデータ</param>
        ''' <remarks>改行を含む</remarks>
        Public Sub AppendLine(ByVal ParamArray values As String())
            For Each value As String In values
                Append(value)
            Next
            [Next]()
        End Sub

        ''' <summary>
        ''' 次の行へ移動する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub [Next]()
            csvLines.Add(Join(currentColumnValues.ToArray, Separator))
            currentColumnValues.Clear()
        End Sub

        ''' <summary>
        ''' Csvファイルにインスタンスの内容を書き込む
        ''' </summary>
        ''' <param name="appends">追記する場合、true</param>
        ''' <remarks></remarks>
        Public Sub Write(Optional ByVal appends As Boolean = False)
            Dim value As String = Join(csvLines.ToArray, vbCrLf)
            If appends Then
                FileUtil.AppendFile(fileName, value)
            Else
                FileUtil.WriteFile(fileName, value)
            End If
        End Sub

    End Class
End Namespace