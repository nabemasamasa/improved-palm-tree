Public Class DotNetUtil

    ''' <summary>
    ''' .NET Framework 3.5 以降がインストール済みかを返す
    ''' </summary>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function HasInstalledDotNet35() As Boolean
        Try
            ' System.Threading.ReaderWriterLockSlim は .NET3.5で追加されたクラス
            ' .NET3.5がローカルPCにインストールされていなければ、エラーになる
            Dim hoge As New System.Threading.ReaderWriterLockSlim
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' クラスが属するプロジェクトのコンパイル日時を返す
    ''' </summary>
    ''' <param name="aType">コンパイル日時を取得するプロジェクトのクラスタイプ</param>
    ''' <returns>コンパイル日時</returns>
    ''' <remarks></remarks>
    Public Shared Function GetCompiledDateOfProject(ByVal aType As Type) As DateTime
        Return ConvCompiledDate(aType.Assembly.GetName.Version)
    End Function

    ''' <summary>
    ''' アプリケーションの（スタートアッププロジェクトの）コンパイル日時を返す
    ''' </summary>
    ''' <returns>コンパイル日時</returns>
    ''' <remarks></remarks>
    Public Shared Function GetCompiledDateOfSolution() As DateTime
        Return ConvCompiledDate(My.Application.Info.Version)
    End Function

    ''' <summary>
    ''' コンパイル日時にして返す
    ''' </summary>
    ''' <param name="ver">（AssemblyVersionのBuild番号以下を'*'にしている）バージョン情報</param>
    ''' <returns>コンパイル日時</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvCompiledDate(ByVal ver As Version) As Date
        Dim result As DateTime = DateUtil.CCDate("2000/01/01")
        result = result.AddDays(ver.Build)
        result = result.AddSeconds(ver.Revision * 2)
        Return result
    End Function
End Class
