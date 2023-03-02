Imports System.Text

Namespace Util.Fixed
    ''' <summary>
    ''' 固定長列の（繰り返し単位などの）グループ化を表わすクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FixedGroup : Implements IFixedEntry

        Private _Name As String
        Private _Repeat As Integer

        ''' <summary>Folder内で先頭からのoffset位置</summary>
        Private _offset As Integer

        Private lengthContainsChildren As Integer = -1
        Private ReadOnly entryByName As Dictionary(Of String, IFixedEntry)

        Private ReadOnly children As IFixedEntry()

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="children">内包する要素</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ParamArray children As IFixedEntry())
            Me.New(Nothing, 1, children)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">グループ名</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <param name="children">内包する要素</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal repeat As Integer, ByVal ParamArray children As IFixedEntry())

            If name IsNot Nothing Then
                Me._Name = name.ToUpper
            End If
            Me._Repeat = repeat
            Me.children = children

            Me.entryByName = New Dictionary(Of String, IFixedEntry)
            For Each info As IFixedEntry In children
                If info.Name Is Nothing Then
                    Throw New InvalidOperationException("（Name無しの）root groupは、他のGroupに内包出来ません.")
                End If
                entryByName.Add(info.Name, info)
            Next

            If repeat <= 0 Then
                Throw New ArgumentException("繰り返し数は1以上を指定すべき.", "repeat")
            End If

            InitializeOffset(0)
        End Sub

        ''' <summary>
        ''' 指定の型のEntryを抽出する
        ''' </summary>
        ''' <typeparam name="T">抽出する型</typeparam>
        ''' <returns>抽出したEntry[]</returns>
        ''' <remarks></remarks>
        Public Function DetectEntries(Of T As IFixedEntry)() As T()
            Dim result As New List(Of T)
            For Each entry As IFixedEntry In children
                If TypeOf entry Is T Then
                    result.Add(DirectCast(entry, T))
                End If
                If TypeOf entry Is FixedGroup Then
                    result.AddRange(DirectCast(entry, FixedGroup).DetectEntries(Of T))
                End If
            Next
            Return result.ToArray
        End Function

        Public ReadOnly Property Offset() As Integer Implements IFixedEntry.Offset
            Get
                Return _offset
            End Get
        End Property

        Public ReadOnly Property Length() As Integer Implements IFixedEntry.Length
            Get
                If lengthContainsChildren < 0 Then
                    lengthContainsChildren = 0
                    For Each info As IFixedEntry In children
                        lengthContainsChildren += info.Length * info.Repeat
                    Next
                End If
                Return lengthContainsChildren
            End Get
        End Property

        Public ReadOnly Property Name() As String Implements IFixedEntry.Name
            Get
                Return _Name
            End Get
        End Property

        Public ReadOnly Property Repeat() As Integer Implements IFixedEntry.Repeat
            Get
                Return _Repeat
            End Get
        End Property

        Public Function ContainsName(ByVal childName As String) As Boolean Implements IFixedEntry.ContainsName

            If childName Is Nothing Then
                Return False
            End If
            Return entryByName.ContainsKey(childName.ToUpper)
        End Function

        Public Function GetChlid(ByVal childName As String) As IFixedEntry Implements IFixedEntry.GetChlid

            If childName Is Nothing Then
                Return Nothing
            End If
            If Not entryByName.ContainsKey(childName.ToUpper) Then
                Return Nothing
            End If
            Return entryByName(childName.ToUpper)
        End Function

        Public Sub InitializeOffset(ByVal offset As Integer) Implements IFixedEntry.InitializeOffset
            Me._offset = offset

            Dim offsetInFolder As Integer = 0
            For Each info As IFixedEntry In children
                info.InitializeOffset(offsetInFolder)
                offsetInFolder += info.Length * info.Repeat
            Next
        End Sub

        ''' <summary>
        ''' 値を固定長文字列にする
        ''' </summary>
        ''' <param name="value">値</param>
        ''' <returns>固定長文字列</returns>
        ''' <remarks></remarks>
        Public Function Format(ByVal value As Object) As String Implements IFixedEntry.Format
            Dim result As New StringBuilder
            For Each info As IFixedEntry In children
                Dim innerString As String = info.Format(value)
                For aRepeat As Integer = 0 To info.Repeat - 1
                    result.Append(innerString)
                Next
            Next
            Return result.ToString
        End Function

        ''' <summary>
        ''' 固定長文字列を値にする
        ''' </summary>
        ''' <param name="fixedString">固定長文字列</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function Parse(ByVal fixedString As String) As Object Implements IFixedEntry.Parse
            Return Nothing
        End Function
    End Class
End Namespace