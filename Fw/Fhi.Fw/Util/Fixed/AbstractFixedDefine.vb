Imports System.Text

Namespace Util.Fixed
    ''' <summary>
    ''' 固定長Root情報に従い、固定長文字列を取得・設定することを担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class AbstractFixedDefine

        ''' <summary>「残りすべて」を意味する長さ</summary>
        Friend Const LENGTH_AS_OTHERS As Integer = -1

        Private _rootHelper As FixedGroupHelper

        ''' <summary>固定長文字列の全体</summary>
        Private _fixedString As String

        ''' <summary>数値型がNullのときzero埋めするか？</summary>
        Private _isZeroPaddingIfNull As Boolean

        ''' <summary>固定長文字列の全体</summary>
        Public Property FixedString() As String
            Get
                Return _fixedString
            End Get
            Set(ByVal value As String)
                _fixedString = value
            End Set
        End Property

        ''' <summary>数値型がNullのときzero埋めするか？</summary>
        Public Property IsZeroPaddingIfNull() As Boolean
            Get
                Return _isZeroPaddingIfNull
            End Get
            Set(ByVal value As Boolean)
                _isZeroPaddingIfNull = value
                For Each number As FixedNumber In GetRootHelper.Group.DetectEntries(Of FixedNumber)()
                    number.IsZeroPaddingIfNull = _isZeroPaddingIfNull
                Next
            End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub InitializeIfNull()
            If _fixedString IsNot Nothing Then
                Return
            End If
            InitializeFixedString()
        End Sub

        ''' <summary>
        ''' 固定長文字列を初期化する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub InitializeFixedString()

            _fixedString = GetRootHelper.Group.Format(Nothing)
        End Sub

        Private Function GetRootHelper() As FixedGroupHelper
            If _rootHelper Is Nothing Then
                _rootHelper = New FixedGroupHelper(GetRootEntryImpl())
                IsZeroPaddingIfNull = _isZeroPaddingIfNull
            End If
            Return _rootHelper
        End Function

        ''' <summary>
        ''' 指定位置の値を取得する
        ''' </summary>
        ''' <param name="name">位置を表わすパス名（グループ区切りは"."）  ex. "HOGE[1].NAME"</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetValue(ByVal name As String) As Object
            Dim offset As Integer = GetRootHelper.GetOffset(name)
            Dim entry As IFixedEntry = GetRootHelper.GetEntry(name)

            Dim requireLength As Integer = offset + If(entry.Length = LENGTH_AS_OTHERS, 0, entry.Length)
            If _fixedString.Length < requireLength Then
                Throw New InvalidOperationException
            End If
            If entry.Length = LENGTH_AS_OTHERS Then
                Return entry.Parse(_fixedString.Substring(offset))
            End If
            Return entry.Parse(_fixedString.Substring(offset, entry.Length))
        End Function

        ''' <summary>
        ''' 指定位置に値を設定する
        ''' </summary>
        ''' <param name="name">位置を表わすパス名（グループ区切りは"."）  ex. "HOGE[1].NAME"</param>
        ''' <param name="value"></param>
        ''' <remarks></remarks>
        Public Sub SetValue(ByVal name As String, ByVal value As Object)
            InitializeIfNull()

            Dim offset As Integer = GetRootHelper.GetOffset(name)
            Dim entry As IFixedEntry = GetRootHelper.GetEntry(name)
            If entry.Length <> LENGTH_AS_OTHERS AndAlso _fixedString.Length < offset + entry.Length Then
                Throw New InvalidOperationException(String.Format("固定長サイズを超えた箇所に値を設定できない. 繰返し数は適切か？ name='{0}'", name))
            End If

            Dim result As New StringBuilder
            If 0 < offset Then
                result.Append(_fixedString.Substring(0, offset))
            End If
            Dim repeat As Integer = 0
            If TypeUtil.IsArrayOrCollection(value) Then
                If entry.Length = LENGTH_AS_OTHERS Then
                    Throw New InvalidOperationException("配列の場合、文字列の長さを明確に指定してください")
                End If
                For Each obj As Object In VoUtil.ConvObjectToArray(value)
                    result.Append(entry.Format(obj))
                    repeat += 1
                Next
            Else
                result.Append(entry.Format(value))
                repeat += 1
            End If

            If entry.Length <> LENGTH_AS_OTHERS AndAlso offset + entry.Length * repeat < _fixedString.Length Then
                result.Append(_fixedString.Substring(offset + entry.Length * repeat))
            End If

            _fixedString = result.ToString
        End Sub

        ''' <summary>
        ''' 先頭（Root）からのOffset位置を返す
        ''' </summary>
        ''' <param name="name">Ex. "GRP1[1].SUB[2].NAME"</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetOffset(ByVal name As String) As Integer
            Return GetRootHelper.GetOffset(name)
        End Function

        ''' <summary>
        ''' 固定長Root情報を取得する
        ''' </summary>
        ''' <returns>固定長Root情報</returns>
        ''' <remarks></remarks>
        Protected MustOverride Function GetRootEntryImpl() As FixedGroup

    End Class
End Namespace