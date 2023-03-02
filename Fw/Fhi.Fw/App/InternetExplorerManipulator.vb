Option Strict Off

Namespace App
    ''' <summary>
    ''' IE操作を担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class InternetExplorerManipulator

        ''' <summary>
        ''' hrefリンクの文言を取得する
        ''' </summary>
        ''' <param name="href">リンクURL</param>
        ''' <returns>文言(InnerHtml)</returns>
        ''' <remarks></remarks>
        Public Function DetectInnerHtmlByHref(ByVal href As String) As String
            Const IE8 As String = "HTMLDocumentClass"
            Const IE11 As String = "HTMLDocument"
            Dim shellWindowsType As Type = GetSHDocVwShellWindowsType()
            Dim shellWindows As Object = Activator.CreateInstance(shellWindowsType)
            Try
                For Each shellWindow As Object In DirectCast(shellWindows, IEnumerable)
                    Try
                        Dim document As Object = shellWindow.Document
                        If Not IE8.Equals(TypeName(document)) AndAlso Not IE11.Equals(TypeName(document)) Then
                            'IE以外(エクスプローラ)は除外する
                            Continue For
                        End If
                        Dim result As String = RecurDetectInnerHtmlByHref(document, href)
                        If result IsNot Nothing Then
                            Return result
                        End If
                    Finally
                        ReleaseComObject(shellWindow)
                    End Try
                Next
            Finally
                ReleaseComObject(shellWindows)
            End Try
            Return Nothing
        End Function

        Private Function RecurDetectInnerHtmlByHref(ByVal ieDocument As Object, ByVal href As String) As String
            Dim frames As Object = ieDocument.All.Tags("frame")
            If 0 < frames.Length Then
                For i As Integer = 0 To frames.Length - 1
                    ' frames(i).Document だと無限ループになる. 意味わかんない
                    Dim result As String = RecurDetectInnerHtmlByHref(ieDocument.Frames.Item(i).Document, href)
                    If result IsNot Nothing Then
                        Return result
                    End If
                Next
            End If
            Dim anchors As Object = ieDocument.All.Tags("A")
            For i As Integer = 0 To anchors.Length - 1
                If anchors(i).href IsNot Nothing AndAlso anchors(i).href.Equals(href) Then
                    Return anchors(i).InnerHtml
                End If
            Next
            Return Nothing
        End Function

        Private Shared Sub ReleaseComObject(Of T As Class)(ByRef comObject As T, Optional ByVal force As Boolean = False)
            If comObject Is Nothing Then
                Return
            End If
            If System.Runtime.InteropServices.Marshal.IsComObject(comObject) Then
                If force Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comObject)
                Else
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject)
                End If
            End If
        End Sub

        ''' <summary>
        ''' SHDocVw.ShellWindows型を返す
        ''' </summary>
        ''' <returns>SHDocVw.ShellWindows型</returns>
        ''' <remarks></remarks>
        Private Shared Function GetSHDocVwShellWindowsType() As Type
            Return Type.GetTypeFromCLSID(New Guid("9BA05972-F6A8-11CF-A442-00A0C90A8F39"))
        End Function

        ''' <summary>
        ''' （内包する）タグを除去する
        ''' </summary>
        ''' <param name="html">HTMLテキスト</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RemoveTags(ByVal html As String) As String
            If html Is Nothing Then
                Return Nothing
            End If
            Dim removingHtml As String = html
            Dim index As Integer = removingHtml.IndexOf("<")
            While 0 <= index
                Dim endIndex As Integer = removingHtml.IndexOf(">", index)
                If endIndex < 0 Then
                    Exit While
                End If
                removingHtml = removingHtml.Substring(0, index) & removingHtml.Substring(endIndex + 1)
                index = removingHtml.IndexOf("<")
            End While
            Return removingHtml
        End Function

    End Class
End Namespace
