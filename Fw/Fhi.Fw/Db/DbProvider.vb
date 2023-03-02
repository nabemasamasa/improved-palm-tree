Imports System.Text

Namespace Db
    ''' <summary>
    ''' 各DBごとの設定を担う（ギャップを埋める）クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DbProvider

        ''' <summary>SqlServer接続用のProvider定数</summary>
        Public Shared ReadOnly SqlServer As New DbProvider("System.Data.SqlClient", New SqlServerBehavior)

        ''' <summary>Linkサーバー（SqlServer内）接続用のProvider定数</summary>
        Public Shared ReadOnly LinkServerOnSqlServer As New DbProvider("System.Data.SqlClient", New LinkServerOnSqlServerBehavior)

        ''' <summary>SqlServerCE接続用のProvider定数</summary>
        Public Shared ReadOnly SqlServerCe As New DbProvider("System.Data.SqlServerCe.3.5", New SqlServerCeBehavior)

        ''' <summary>MDB接続用のProvider定数</summary>
        Public Shared ReadOnly MsAccess As New DbProvider("System.Data.OleDb", New MsAccessBehavior)

        ''' <summary>SQLite接続用のProvider定数</summary>
        Public Shared ReadOnly SQLite As New DbProvider("System.Data.SQLite", New SQLiteBehavior)
        ' ↑SQLiteを使うには、以下2点が必要
        ' 1. SQLite.DLL への参照設定
        ' 2. app.configに以下を設定
        ' <configuration>
        '   <system.data>
        '     <DbProviderFactories>
        '       <remove invariant="System.Data.SQLite" />
        '       <add name="SQLite Data Provider"
        '          invariant="System.Data.SQLite"
        '          description=".Net Framework Data Provider for SQLite"
        '          type="System.Data.SQLite.SQLiteFactory, System.Data.SQLite" />
        '     </DbProviderFactories>
        '   </system.data>
        ' </configuration>

        ''' <summary>MySQL接続用のProvider定数</summary>
        Public Shared ReadOnly MySql As New DbProvider("MySql.Data.MySqlClient", New MySqlBehavior)

        ''' <summary>すべてのProvider定数</summary>
        Public Shared ReadOnly AllDbProviders As DbProvider() = {SqlServer, SqlServerCe, MsAccess, SQLite, MySql}

        ''' <summary>
        ''' プロバイダを取得する
        ''' </summary>
        ''' <param name="providerName">プロバイダ名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Detect(ByVal providerName As String) As DbProvider
            For Each provider As DbProvider In AllDbProviders
                If provider.ProviderString.Equals(providerName) Then
                    Return provider
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>Provider名</summary>
        Public ReadOnly ProviderString As String

        Private ReadOnly behavior As IBehavior

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="providerString">DBプロバイダ名</param>
        ''' <param name="behavior">各DBプロバイダ用の振る舞い</param>
        ''' <remarks></remarks>
        Private Sub New(ByVal providerString As String, ByVal behavior As IBehavior)
            Me.ProviderString = providerString
            Me.behavior = behavior
        End Sub

        ''' <summary>
        ''' 接続文字列を解決する
        ''' </summary>
        ''' <param name="connectionString">接続文字列</param>
        ''' <returns>解決した接続文字列</returns>
        ''' <remarks></remarks>
        Public Function ResolveConnectionString(ByVal connectionString As String) As String
            Return behavior.ResolveConnectionString(connectionString)
        End Function

        ''' <summary>SQLタイムアウト(s)（初期値）</summary>
        Public ReadOnly Property DefaultCommandTimeout() As Integer
            Get
                Return behavior.DefaultCommandTimeout
            End Get
        End Property

        ''' <summary>文字列のDbType</summary>
        Public ReadOnly Property DbTypeOfString() As DbType
            Get
                Return behavior.DbTypeOfString
            End Get
        End Property

        ''' <summary>ロックタイムアウト(ms)（初期値）</summary>
        Public ReadOnly Property DefaultLockTimeoutMillis() As Long
            Get
                Return behavior.DefaultLockTimeoutMillis()
            End Get
        End Property

        ''' <summary>
        ''' ロックタイムアウト(ms)を設定する
        ''' </summary>
        ''' <param name="db">DB接続オブジェクト</param>
        ''' <param name="millis">ロックタイムアウト(ms)</param>
        ''' <remarks></remarks>
        Public Sub SetLockTimeoutMillis(ByVal db As DbAccess, ByVal millis As Long)
            behavior.SetLockTimeoutMillis(db, millis)
        End Sub

        ''' <summary>
        ''' 行ロック参照SQLを作成する
        ''' </summary>
        ''' <param name="select">Select句</param>
        ''' <param name="from">From句</param>
        ''' <param name="where">Where句</param>
        ''' <returns>行ロック参照SQL</returns>
        ''' <remarks></remarks>
        Public Function MakeSelectForUpdate(ByVal [select] As String, ByVal from As String, ByVal where As String) As String
            Return behavior.MakeSelectForUpdate([select], from, where)
        End Function

        ''' <summary>
        ''' DBプロバイダ名からインスタンスを取得する
        ''' </summary>
        ''' <param name="providerString">DBプロバイダ名</param>
        ''' <returns>プロバイダのインスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function [Get](ByVal providerString As String) As DbProvider
            For Each aProvider As DbProvider In AllDbProviders
                If aProvider.ToString.Equals(providerString) Then
                    Return aProvider
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' 複数行一括Insertできることを保証する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub AssertCanMultiValuesInsert()
            If Not behavior.CanMultiValuesInsert() Then
                Throw New InvalidOperationException("複数行一括insertできるDBではない")
            End If
        End Sub

        Private Interface IBehavior

            ''' <summary>
            ''' 行ロック参照SQLを作成する
            ''' </summary>
            ''' <param name="select">Select句</param>
            ''' <param name="from">From句</param>
            ''' <param name="where">Where句</param>
            ''' <returns>行ロック参照SQL</returns>
            ''' <remarks></remarks>
            Function MakeSelectForUpdate(ByVal [select] As String, ByVal from As String, ByVal where As String) As String

            ''' <summary>
            ''' 接続文字列を解決する
            ''' </summary>
            ''' <param name="connectionString">接続文字列</param>
            ''' <returns>解決した接続文字列</returns>
            ''' <remarks></remarks>
            Function ResolveConnectionString(ByVal connectionString As String) As String

            ''' <summary>
            ''' ロックタイムアウト(ms)を設定する
            ''' </summary>
            ''' <param name="db">DB接続オブジェクト</param>
            ''' <param name="millis">ロックタイムアウト(ms)</param>
            ''' <remarks></remarks>
            Sub SetLockTimeoutMillis(ByVal db As DbAccess, ByVal millis As Long)

            ''' <summary>SQLタイムアウト(s)（初期値）</summary>
            ReadOnly Property DefaultCommandTimeout() As Integer

            ''' <summary>ロックタイムアウト(ms)（初期値）を返す</summary>
            ReadOnly Property DefaultLockTimeoutMillis() As Long

            ''' <summary>文字列のDbType</summary>
            ReadOnly Property DbTypeOfString() As DbType

            ''' <summary>複数行一括insertができるか？</summary>
            ReadOnly Property CanMultiValuesInsert() As Boolean

        End Interface

        Private Class SqlServerBehavior : Implements IBehavior

            Public ReadOnly Property DefaultCommandTimeout() As Integer Implements IBehavior.DefaultCommandTimeout
                Get
                    Return 300
                End Get
            End Property

            Public Overridable ReadOnly Property DbTypeOfString() As DbType Implements IBehavior.DbTypeOfString
                Get
                    Return DbType.String ' Stringだと、LinkServer接続に対してSQL結果が返って来ないという現象が発生する
                End Get
            End Property

            Public ReadOnly Property DefaultLockTimeoutMillis() As Long Implements IBehavior.DefaultLockTimeoutMillis
                Get
                    Return 10000
                End Get
            End Property

            Public Overridable ReadOnly Property CanMultiValuesInsert As Boolean Implements IBehavior.CanMultiValuesInsert
                Get
                    Return True
                End Get
            End Property

            Public Sub SetLockTimeoutMillis(ByVal db As DbAccess, ByVal millis As Long) Implements IBehavior.SetLockTimeoutMillis
                If millis < 0 Then
                    Return
                End If
                Dim helper As New DbAccessHelper(db)
                helper.ExecuteSetting(String.Format("SET LOCK_TIMEOUT {0}", millis))
            End Sub

            Public Function MakeSelectForUpdate(ByVal [select] As String, ByVal from As String, ByVal where As String) As String Implements IBehavior.MakeSelectForUpdate
                Dim result As New StringBuilder
                Return result.Append([select]).Append(" ").Append(from).Append(" WITH (UPDLOCK,NOWAIT) ").Append(where).ToString()
            End Function

            Public Function ResolveConnectionString(ByVal connectionString As String) As String Implements IBehavior.ResolveConnectionString
                If connectionString Is Nothing Then
                    Return Nothing
                End If

                Const SEPARATOR As Char = ";"c
                '1次的な対応? ADO.NETではDRIVER={SQL Server};はいらない.
                'あるとエラーになるため取り除く.
                Dim result As New StringBuilder
                For Each token As String In connectionString.Split(SEPARATOR)
                    If Not token.ToUpper.StartsWith("DRIVER=") Then
                        result.Append(token).Append(SEPARATOR)
                    End If
                Next

                Return result.ToString
            End Function
        End Class

        Private Class MySqlBehavior : Implements IBehavior

            Public ReadOnly Property DefaultCommandTimeout() As Integer Implements IBehavior.DefaultCommandTimeout
                Get
                    Return 300
                End Get
            End Property

            Public Overridable ReadOnly Property DbTypeOfString() As DbType Implements IBehavior.DbTypeOfString
                Get
                    Return DbType.String
                End Get
            End Property

            Public ReadOnly Property DefaultLockTimeoutMillis() As Long Implements IBehavior.DefaultLockTimeoutMillis
                Get
                    Return 10000
                End Get
            End Property

            Public ReadOnly Property CanMultiValuesInsert As Boolean Implements IBehavior.CanMultiValuesInsert
                Get
                    Return False
                End Get
            End Property

            Public Sub SetLockTimeoutMillis(ByVal db As DbAccess, ByVal millis As Long) Implements IBehavior.SetLockTimeoutMillis
                If millis < 0 Then
                    Return
                End If
                Dim helper As New DbAccessHelper(db)
                helper.ExecuteSetting(String.Format("SET innodb_lock_wait_timeout={0};", millis))
            End Sub

            Public Function MakeSelectForUpdate(ByVal [select] As String, ByVal from As String, ByVal where As String) As String Implements IBehavior.MakeSelectForUpdate
                Dim result As New StringBuilder
                Return result.Append([select]).Append(" ").Append(from).Append(" ").Append(where).Append(" FOR UPDATE;").ToString()
            End Function

            Public Function ResolveConnectionString(ByVal connectionString As String) As String Implements IBehavior.ResolveConnectionString
                If connectionString Is Nothing Then
                    Return Nothing
                End If

                ' DRIVER={MySql};はいらないので、ある場合は取り除く
                Const SEPARATOR As Char = ";"c
                Dim result As New StringBuilder
                For Each token As String In connectionString.Split(SEPARATOR)
                    If Not token.ToUpper.StartsWith("DRIVER=") Then
                        result.Append(token).Append(SEPARATOR)
                    End If
                Next

                Return result.ToString
            End Function
        End Class

        Private Class LinkServerOnSqlServerBehavior : Inherits SqlServerBehavior

            Public Overrides ReadOnly Property DbTypeOfString() As DbType
                Get
                    Return DbType.AnsiString ' Stringだと、LinkServer接続に対してSQL結果が返って来ないという現象が発生する
                End Get
            End Property

            Public Overrides ReadOnly Property CanMultiValuesInsert As Boolean
                Get
                    Return True
                End Get
            End Property

        End Class

        Private Class SqlServerCeBehavior : Implements IBehavior

            Public ReadOnly Property DefaultCommandTimeout() As Integer Implements IBehavior.DefaultCommandTimeout
                Get
                    Return 0 ' 0じゃないとエラーになる
                End Get
            End Property

            Public ReadOnly Property DbTypeOfString() As DbType Implements IBehavior.DbTypeOfString
                Get
                    Return DbType.String ' AnsiStringだとエラーになる
                End Get
            End Property

            Public ReadOnly Property DefaultLockTimeoutMillis() As Long Implements IBehavior.DefaultLockTimeoutMillis
                Get
                    Return 10000
                End Get
            End Property

            Public ReadOnly Property CanMultiValuesInsert As Boolean Implements IBehavior.CanMultiValuesInsert
                Get
                    Return False
                End Get
            End Property

            Public Sub SetLockTimeoutMillis(ByVal db As DbAccess, ByVal millis As Long) Implements IBehavior.SetLockTimeoutMillis
                If millis < 0 Then
                    Return
                End If
                Dim helper As New DbAccessHelper(db)
                helper.ExecuteSetting(String.Format("SET LOCK_TIMEOUT {0}", millis))
            End Sub

            Public Function MakeSelectForUpdate(ByVal [select] As String, ByVal from As String, ByVal where As String) As String Implements IBehavior.MakeSelectForUpdate
                Dim result As New StringBuilder
                Return result.Append([select]).Append(" ").Append(from).Append(" WITH (UPDLOCK) ").Append(where).ToString()
            End Function

            Public Function ResolveConnectionString(ByVal connectionString As String) As String Implements IBehavior.ResolveConnectionString

                Return connectionString
            End Function
        End Class

        Private Class MsAccessBehavior : Implements IBehavior

            Public ReadOnly Property DefaultCommandTimeout() As Integer Implements IBehavior.DefaultCommandTimeout
                Get
                    Return 300
                End Get
            End Property

            Public ReadOnly Property DbTypeOfString() As DbType Implements IBehavior.DbTypeOfString
                Get
                    Return DbType.String
                End Get
            End Property

            Public ReadOnly Property DefaultLockTimeoutMillis() As Long Implements IBehavior.DefaultLockTimeoutMillis
                Get
                    Return -1
                End Get
            End Property

            Public ReadOnly Property CanMultiValuesInsert As Boolean Implements IBehavior.CanMultiValuesInsert
                Get
                    Return False
                End Get
            End Property

            Public Sub SetLockTimeoutMillis(ByVal db As DbAccess, ByVal millis As Long) Implements IBehavior.SetLockTimeoutMillis
                ' nop
            End Sub

            Public Function MakeSelectForUpdate(ByVal [select] As String, ByVal from As String, ByVal where As String) As String Implements IBehavior.MakeSelectForUpdate
                Dim result As New StringBuilder
                Return result.Append([select]).Append(from).Append(where).ToString()
            End Function

            Public Function ResolveConnectionString(ByVal connectionString As String) As String Implements IBehavior.ResolveConnectionString
                Return connectionString
            End Function
        End Class

        Private Class SQLiteBehavior : Implements IBehavior

            Public ReadOnly Property DefaultCommandTimeout() As Integer Implements IBehavior.DefaultCommandTimeout
                Get
                    Return 300
                End Get
            End Property

            Public ReadOnly Property DbTypeOfString() As DbType Implements IBehavior.DbTypeOfString
                Get
                    Return DbType.String
                End Get
            End Property

            Public ReadOnly Property DefaultLockTimeoutMillis() As Long Implements IBehavior.DefaultLockTimeoutMillis
                Get
                    Return -1
                End Get
            End Property

            Public Sub SetLockTimeoutMillis(ByVal db As DbAccess, ByVal millis As Long) Implements IBehavior.SetLockTimeoutMillis
                If millis < 1000 Then
                    db.CommandTimeOut = 1
                Else
                    db.CommandTimeOut = CInt(millis / 1000)
                End If
            End Sub

            Public ReadOnly Property CanMultiValuesInsert As Boolean Implements IBehavior.CanMultiValuesInsert
                Get
                    Return False
                End Get
            End Property

            Public Function MakeSelectForUpdate(ByVal [select] As String, ByVal from As String, ByVal where As String) As String Implements IBehavior.MakeSelectForUpdate
                Dim result As New StringBuilder
                Return result.Append([select]).Append(from).Append(where).ToString()
            End Function

            Public Function ResolveConnectionString(ByVal connectionString As String) As String Implements IBehavior.ResolveConnectionString

                Return connectionString
            End Function
        End Class

    End Class
End Namespace