Imports System.Runtime.InteropServices
Imports System.Text

Namespace Sys

    ''' <summary>
    ''' Windows用INIファイルのアクセスを提供します.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PrivateProfile

        ''' <summary>値の最大読込サイズの初期値</summary>
        Public Const DEFAULT_MAX_VALUE_LENGTH As Integer = 4096

        ''' <summary>INIファイル(フルパス)</summary>
        Public IniFile As String
        ''' <summary>INIファイルの値の最大読込サイズ</summary>
        Public MaxValueLength As Integer

        ''' <summary>読取専用のiniファイルなら、trueに</summary>
        Public IsReadonly As Boolean

#Region "DllImports..."
        ''' <summary>
        ''' INIファイルに値を設定します.
        ''' </summary>
        ''' <param name="lpAppName">アプリケーション名(セクション)</param>
        ''' <param name="lpKeyName">キー名</param>
        ''' <param name="lpString">設定する値</param>
        ''' <param name="lpFileName">INIファイル(フルパス)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DllImport("KERNEL32.DLL")> _
        Protected Shared Function WritePrivateProfileString( _
            ByVal lpAppName As String, _
            ByVal lpKeyName As String, _
            ByVal lpString As String, _
            ByVal lpFileName As String) As Integer
        End Function

        ''' <summary>
        ''' INIファイルから値を取得します.
        ''' </summary>
        ''' <param name="lpAppName">アプリケーション名(セクション)</param>
        ''' <param name="lpKeyName">キー名</param>
        ''' <param name="lpDefault">デフォルト値</param>
        ''' <param name="lpReturnedString">戻り値</param>
        ''' <param name="nSize">戻り値の項目長</param>
        ''' <param name="lpFileName">INIファイル(フルパス)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)> _
        Protected Shared Function GetPrivateProfileString( _
            ByVal lpAppName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As StringBuilder, _
            ByVal nSize As Integer, _
            ByVal lpFileName As String) As Integer
        End Function

        ''' <summary>
        ''' INIファイルから値を取得します.
        ''' </summary>
        ''' <param name="lpAppName">アプリケーション名(セクション)</param>
        ''' <param name="lpKeyName">キー名</param>
        ''' <param name="lpDefault">デフォルト値</param>
        ''' <param name="lpReturnedString">戻り値</param>
        ''' <param name="nSize">戻り値の項目長</param>
        ''' <param name="lpFileName">INIファイル(フルパス)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DllImport("KERNEL32.DLL", EntryPoint:="GetPrivateProfileStringA")> _
        Protected Shared Function GetPrivateProfileStringByByteArray( _
            ByVal lpAppName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As Byte(), _
            ByVal nSize As Integer, _
            ByVal lpFileName As String) As Integer
        End Function
#End Region


        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="iniFile">INIファイル(フルパス)</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFile As String)
            Me.New(iniFile, DEFAULT_MAX_VALUE_LENGTH)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="iniFile">INIファイル(フルパス)</param>
        ''' <param name="maxValueLength">最大値読込サイズ</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFile As String, ByVal maxValueLength As Integer)
            Me.IniFile = iniFile
            Me.MaxValueLength = maxValueLength
        End Sub

        ''' <summary>
        ''' INIファイルに値を設定します.
        ''' </summary>
        ''' <param name="iniDiv">アプリケーション名(セクション)</param>
        ''' <param name="iniKey">キー名</param>
        ''' <param name="value">設定する値</param>
        ''' <remarks></remarks>
        Public Sub Write(ByVal iniDiv As String, ByVal iniKey As String, ByVal value As String)
            If IsReadonly Then
                Throw New InvalidOperationException("読取専用マークがされたiniファイル " & IniFile)
            End If
            WritePrivateProfileString(iniDiv, iniKey, value, IniFile)
        End Sub

        ''' <summary>
        ''' INIファイルから値を取得します.
        ''' </summary>
        ''' <param name="iniDiv">アプリケーション名(セクション)</param>
        ''' <param name="iniKey">キー名</param>
        ''' <returns>INIファイルに設定されている値</returns>
        ''' <remarks></remarks>
        Public Function Read(ByVal iniDiv As String, ByVal iniKey As String) As String

            ' インスタンスは毎回作ること。
            Dim ret As New StringBuilder(MaxValueLength)

            GetPrivateProfileString(iniDiv, iniKey, "", ret, MaxValueLength, IniFile)

            Return ret.ToString
        End Function

        ''' <summary>
        ''' セクションとキーのペアを削除する
        ''' </summary>
        ''' <param name="section">セクション名</param>
        ''' <param name="key">キー名</param>
        ''' <remarks></remarks>
        Public Sub Remove(ByVal section As String, ByVal key As String)
            Write(section, key, Nothing)
        End Sub

        ''' <summary>
        ''' セクション名の一覧を取得する
        ''' </summary>
        ''' <returns>セクション名の一覧</returns>
        ''' <remarks></remarks>
        Public Function GetSections() As String()
            Dim ar As Byte() = New Byte(1023) {}
            Dim resultSize As Integer = GetPrivateProfileStringByByteArray(Nothing, Nothing, String.Empty, ar, ar.Length, IniFile)
            If resultSize = 0 Then
                Return New String() {}
            End If
            Dim result As String = Encoding.[Default].GetString(ar, 0, resultSize - 1)
            Return result.Split(ControlChars.NullChar)
        End Function

        ''' <summary>
        ''' キー名の一覧を取得する
        ''' </summary>
        ''' <param name="section">セクション名</param>
        ''' <returns>キー名の一覧</returns>
        ''' <remarks></remarks>
        Public Function GetKeys(ByVal section As String) As String()
            Dim ar As Byte() = New Byte(1023) {}
            Dim resultSize As Integer = GetPrivateProfileStringByByteArray(section, Nothing, String.Empty, ar, ar.Length, IniFile)
            If resultSize = 0 Then
                Return New String() {}
            End If
            Dim result As String = Encoding.[Default].GetString(ar, 0, resultSize - 1)
            Return result.Split(ControlChars.NullChar)
        End Function

    End Class
End Namespace
