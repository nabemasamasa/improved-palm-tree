Namespace Util.Fixed
    Public Class FixedGroupHelper

        Private baseGroup As FixedGroup

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="baseGroup">Groupインスタンス</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal baseGroup As FixedGroup)
            Me.baseGroup = baseGroup
        End Sub

        Public ReadOnly Property Group() As FixedGroup
            Get
                Return baseGroup
            End Get
        End Property

        ''' <summary>
        ''' Group先頭からのOffset位置を返す
        ''' </summary>
        ''' <param name="name">カンマ区切りの名前  ex. "GRP1[1].SUB[2].NAME"</param>
        ''' <returns>Offset位置</returns>
        ''' <remarks></remarks>
        Public Function GetEntry(ByVal name As String) As IFixedEntry
            Dim parser As New NameParser(name)
            parser.Parse()
            Dim fixedEntry As IFixedEntry = baseGroup
            For Each nameRepeat As NameRepeat In parser.Result
                fixedEntry = fixedEntry.GetChlid(nameRepeat.Name)
                If fixedEntry Is Nothing Then
                    Throw New ArgumentException(String.Format("{0} のうち、{1} は見つかりません.", name, nameRepeat.Name), "name")
                End If
            Next
            Return fixedEntry
        End Function

        ''' <summary>
        ''' Group先頭からのOffset位置を返す
        ''' </summary>
        ''' <param name="name">カンマ区切りの名前  ex. "GRP1[1].SUB[2].NAME"</param>
        ''' <returns>Offset位置</returns>
        ''' <remarks></remarks>
        Public Function GetOffset(ByVal name As String) As Integer
            Dim result As Integer = 0
            Dim parser As New NameParser(name)
            parser.Parse()

            Dim fixedEntry As IFixedEntry = baseGroup
            For Each nameRepeat As NameRepeat In parser.Result
                fixedEntry = fixedEntry.GetChlid(nameRepeat.Name)
                If fixedEntry Is Nothing Then
                    Throw New ArgumentException(String.Format("{0} のうち、{1} は見つかりません.", name, nameRepeat.Name), "name")
                End If
                result += fixedEntry.Offset + fixedEntry.Length * nameRepeat.RepeatIndex
            Next
            Return result
        End Function

        Private Class NameParser

            Private name As String
            Private _result As New List(Of NameRepeat)

            Public Sub New(ByVal name As String)
                Me.name = name
            End Sub

            Public ReadOnly Property Result() As NameRepeat()
                Get
                    Return _result.ToArray
                End Get
            End Property

            Public Sub Parse()
                _result.Clear()
                For Each s As String In name.Split("."c)
                    Dim indexOfStart As Integer = s.IndexOf("["c)
                    Dim indexOfEnd As Integer = s.IndexOf("]"c)
                    If 0 <= indexOfStart AndAlso 0 <= indexOfEnd AndAlso indexOfStart < indexOfEnd AndAlso IsNumeric(s.Substring(indexOfStart + 1, indexOfEnd - indexOfStart - 1)) Then
                        _result.Add(New NameRepeat(s.Substring(0, indexOfStart), CInt(s.Substring(indexOfStart + 1, indexOfEnd - indexOfStart - 1))))
                    Else
                        _result.Add(New NameRepeat(s, 0))
                    End If
                Next
            End Sub
        End Class

        Private Class NameRepeat
            Public Name As String
            Public RepeatIndex As Integer
            Public Sub New(ByVal name As String, ByVal repeatIndex As Integer)
                Me.Name = name
                Me.RepeatIndex = repeatIndex
            End Sub
        End Class
    End Class
End Namespace