Imports System.IO
Imports System.Reflection

''' <summary>
''' アセンブリに関するユーティリティクラス
''' </summary>
''' <remarks></remarks>
Public Class AssemblyUtil

    ''' <summary>
    ''' アセンブリパスを返す
    ''' </summary>
    ''' <returns>アセンブリパス</returns>
    ''' <remarks></remarks>
    Public Shared Function GetPath() As String
        Dim aUriBuilder As UriBuilder = New UriBuilder(Assembly.GetExecutingAssembly.CodeBase)
        ' UNCパス`\\tpc-hoge123\path\to`の場合、aUriBuilder.Path だと`\path\to` になる
        Return Path.GetDirectoryName(Uri.UnescapeDataString(aUriBuilder.Uri.LocalPath))
    End Function

    ''' <summary>
    ''' クラスが属するプロジェクトのアセンブリ名を返す
    ''' </summary>
    ''' <param name="aType">アセンブリ名を取得するクラスタイプ</param>
    ''' <returns>アセンブリ名</returns>
    ''' <remarks></remarks>
    Public Shared Function GetAssemblyNameOfProject(ByVal aType As Type) As String
        Return aType.Assembly.GetName.Name
    End Function

    ''' <summary>
    ''' アプリケーションの（スタートアッププロジェクトの）アセンブリ名を返す
    ''' </summary>
    ''' <returns>アセンブリ名</returns>
    ''' <remarks></remarks>
    Public Shared Function GetAssemblyNameOfSolution() As String
        Return My.Application.Info.AssemblyName
    End Function

    ''' <summary>
    ''' 名前空間内の型一覧を取得する
    ''' </summary>
    ''' <param name="aNamespace">（ルート名前空間含む）名前空間</param>
    ''' <param name="aType">対象アセンブリが実行できる型</param>
    ''' <returns>名前空間内の型一覧</returns>
    ''' <remarks></remarks>
    Public Shared Function GetTypesInNamespace(ByVal aNamespace As String, ByVal aType As Type) As Type()
        Return aType.Assembly.GetTypes().Where(Function(t) t.IsClass AndAlso t.FullName.StartsWith(aNamespace)).ToArray()
    End Function

    ''' <summary>
    ''' コンパイル日時を返す
    ''' </summary>
    ''' <param name="aType">走査対象のプロジェクトが所有する型情報</param>
    ''' <returns>コンパイル日時</returns>
    ''' <remarks></remarks>
    Public Shared Function GetCompiledDate(aType As Type) As DateTime
        Dim ver As Version = aType.Assembly.GetName.Version
        Dim result As DateTime = Convert.ToDateTime("2000/01/01")
        result = result.AddDays(ver.Build)
        result = result.AddSeconds(ver.Revision * 2)
        Return result
    End Function

    ''' <summary>
    ''' 共通言語ランタイム(CLR)が正しい組み合わせであることを検証する
    ''' </summary>
    ''' <typeparam name="Product">製品Projectの１Type</typeparam>
    Public Shared Sub AssertCorrectClrVersion(Of Product)()
        AssertCorrectClrVersion(GetType(Product))
    End Sub
    ''' <summary>
    ''' 共通言語ランタイム(CLR)が正しい組み合わせであることを検証する
    ''' </summary>
    ''' <param name="productType">製品Projectの１Type</param>
    Public Shared Sub AssertCorrectClrVersion(productType As Type)
        Dim fhiFwClrVersion As String = GetFhiFwClrVersion()
        Dim productClrVersion As String = GetAssemblyClrVersion(productType)
        If Not fhiFwClrVersion.Equals(productClrVersion) Then
            Throw New InvalidOperationException(Join({"CLRの組み合わせが正しくない", "Fhi.FwのCLRが" & fhiFwClrVersion, productType.FullName & "のCLRは" & productClrVersion}, vbCrLf))
        End If
    End Sub

    Private Shared Function GetFhiFwClrVersion() As String
        Return GetAssemblyClrVersion(GetType(AssemblyUtil))
    End Function

    Private Shared Function GetAssemblyClrVersion(aType As Type) As String
        Return aType.Assembly.ImageRuntimeVersion
    End Function

End Class