Imports System.Data
Imports System.Data.Common
Imports System.Text

Namespace Db
    ''' <summary>データベースアクセス抽象クラス</summary>
    ''' <remarks>
    ''' データベースへのアクセスを提供する抽象クラスです.
    ''' 
    ''' 他のプロジェクトで使用する場合,このクラスを派生させ,
    ''' 派生先のコンストラクタにて初期化を行ってください.
    ''' 
    ''' ＜例＞
    ''' <code>
    '''     Public NotInheritable Class DbAccess
    '''         Inherits DbAccess
    ''' 
    '''         Public Sub New()
    '''             Me.DbProvider = Settings.Provider              'DBプロバイダ
    '''             Me.ConnectionStr = Settings.ConnectionStr      '接続文字列
    '''             Me.CommandTimeOut = Settings.CommandTimeOut    'SQLコマンドタイムアウト値(秒)
    '''             Me.Initialize()                                '初期化処理
    '''         End Sub
    '''     End Class
    ''' </code>
    ''' 
    ''' プロバイダに依存しないようにするため DbProviderFactory クラスを使用しています.
    ''' このクラス内で例外処理はしていません.例外処理は各処理にて作成してください.
    ''' </remarks>
    Public Class DbAccess
        Implements IDisposable

        ''' <summary>データベースプロバイダ名</summary>
        Protected m_dbProvider As String

        ''' <summary>プロバイダーファクトリーオブジェクト</summary>
        Private ReadOnly m_dbFactory As DbProviderFactory

        ''' <summary>接続文字列</summary>
        Private m_connectionString As String

        ''' <summary>コネクションオブジェクト</summary>
        Private ReadOnly m_cn As DbConnection

        ''' <summary>前回使用したコマンドオブジェクト</summary>
        Private m_cmd As DbCommand

        ''' <summary>トランザクションオブジェクト</summary>
        Private m_trans As DbTransaction

        ''' <summary>パラメーターリスト</summary>
        Private m_dbParams As New List(Of Common.DbParameter)

        ''' <summary>コマンドタイムアウト値(秒)</summary>
        Private m_cmdTimeOut As Integer

        ''' <summary>文字列のDbType</summary>
        Friend ReadOnly DbTypeOfString As DbType

        Private ReadOnly dbProvider As DbProvider

        ''' <summary>コンストラクタ</summary>
        Public Sub New(ByVal aProvider As DbProvider)
            Me.New(aProvider, Nothing)
        End Sub

        Public Sub New(ByVal aProvider As DbProvider, ByVal connectionString As String)

            Me.dbProvider = aProvider
            Me.m_dbFactory = DbProviderFactories.GetFactory(aProvider.ProviderString)
            Me.m_cn = m_dbFactory.CreateConnection()
            Me.m_cmdTimeOut = aProvider.DefaultCommandTimeout
            Me.DbTypeOfString = aProvider.DbTypeOfString

            Me.ConnectionString = aProvider.ResolveConnectionString(connectionString)
        End Sub

        ''' <summary>デストラクタ</summary>
        Protected Overrides Sub Finalize()
            'オブジェクトの破棄
            Me.Dispose(False)

            '基底クラスのデストラクタを呼び出す
            MyBase.Finalize()
        End Sub

        ''' <summary>データベースを開きます</summary>
        Public Sub Open()
            m_cn.Open()
        End Sub

        ''' <summary>データベースへの接続を閉じます</summary>
        Public Sub Close()
            'トランザクションオブジェクトが有効(トランザクション中)の場合,ロールバックを行う.
            If HasTransaction Then
                m_trans.Rollback()
            End If

            'コネクションオブジェクトの状態が閉じていない場合のみ,接続を閉じる.
            If Not m_cn.State = ConnectionState.Closed Then
                m_cn.Close()
            End If
        End Sub

        ''' <summary>トランザクションを開始します</summary>
        Public Overloads Sub BeginTransaction()
            'トランザクション中は処理を抜ける
            If HasTransaction Then
                Return
            End If

            'トランザクション開始
            m_trans = m_cn.BeginTransaction()
        End Sub

        ''' <summary>トランザクションを開始します</summary>
        ''' <param name="isolationLevel">トランザクションの分離レベルを設定します</param>
        Public Overloads Sub BeginTransaction(ByVal isolationLevel As IsolationLevel)
            'トランザクション中は処理を抜ける
            If HasTransaction Then
                Return
            End If

            'トランザクション開始
            m_trans = m_cn.BeginTransaction(isolationLevel)
        End Sub

        ''' <summary>トランザクションをコミットします</summary>
        Public Sub Commit()
            If Not HasTransaction Then
                Return
            End If

            m_trans.Commit()
            m_trans.Dispose()
            m_trans = Nothing
        End Sub

        ''' <summary>トランザクションをロールバックします</summary>
        Public Sub Rollback()
            If Not HasTransaction Then
                Return
            End If

            m_trans.Rollback()
            m_trans.Dispose()
            m_trans = Nothing
        End Sub

        ''' <summary>
        ''' プロシージャを実行する
        ''' </summary>
        ''' <param name="procedureName">プロシージャ名</param>
        ''' <returns>実行結果[]</returns>
        ''' <remarks></remarks>
        Public Overloads Function ExecuteProcedure(ByVal procedureName As String) As List(Of Object)

            '前回使用したコマンドオブジェクトが存在する場合開放を行う.
            If Not m_cmd Is Nothing Then
                m_cmd.Dispose()
            End If

            'コマンドオブジェクトの作成と実行
            m_cmd = m_dbFactory.CreateCommand()
            m_cmd.CommandType = CommandType.StoredProcedure
            Me.SetCommand(m_cmd, procedureName)

            m_cmd.ExecuteNonQuery()

            Dim result As New List(Of Object)
            For Each dbParam As DbParameter In m_dbParams
                If dbParam.Direction = ParameterDirection.InputOutput OrElse dbParam.Direction = ParameterDirection.Output _
                        OrElse dbParam.Direction = ParameterDirection.ReturnValue Then
                    result.Add(dbParam.Value)
                End If
            Next

            '実行後パラメーターをクリアする
            Me.ClearParameters()

            Return result
        End Function

        ''' <summary>SQLステートメントを実行します</summary>
        ''' <param name="sql">実行するSQLステートメント</param>
        ''' <returns>影響を受けた行数</returns>
        Public Overloads Function ExecuteNonQuery(ByVal sql As String) As Integer
            Dim ret As Integer = 0      '戻り値(影響を受けた行数)

            '前回使用したコマンドオブジェクトが存在する場合開放を行う.
            If Not m_cmd Is Nothing Then
                m_cmd.Dispose()
            End If

            'コマンドオブジェクトの作成と実行
            m_cmd = m_dbFactory.CreateCommand()
            Me.SetCommand(m_cmd, sql)
            ret = m_cmd.ExecuteNonQuery()

            '実行後パラメーターをクリアする
            Me.ClearParameters()

            Return ret
        End Function

        ''' <summary>前回使用したコマンドを使用してSQLステートメントを実行します</summary>
        ''' <returns>影響を受けた行数</returns>
        Public Overloads Function ExecuteNonQuery() As Integer
            Dim ret As Integer = 0      '戻り値(影響を受けた行数)

            '前回使用したコマンドオブジェクトを使用して実行
            ret = m_cmd.ExecuteNonQuery()

            '実行後パラメーターをクリアする
            Me.ClearParameters()

            Return ret
        End Function

        ''' <summary>SQLステートメントを実行し, DbDataReader を取得します</summary>
        ''' <param name="sql">実行するSQLステートメント</param>
        ''' <returns>結果</returns>
        Public Overloads Function ExecuteReader(ByVal sql As String) As DbDataReader

            '前回使用したコマンドオブジェクトが存在する場合開放を行う.
            If Not m_cmd Is Nothing Then
                m_cmd.Dispose()
            End If

            'コマンドオブジェクトの作成と実行
            m_cmd = m_dbFactory.CreateCommand()
            Me.SetCommand(m_cmd, sql)
            Dim ret As DbDataReader = m_cmd.ExecuteReader()

            '実行後パラメーターをクリアする
            Me.ClearParameters()

            Return ret
        End Function

        ''' <summary>前回使用したコマンドを使用してSQLステートメントを実行し, DbDataReader を取得します</summary>
        ''' <returns>結果</returns>
        Public Overloads Function ExecuteReader() As DbDataReader

            '前回使用したコマンドオブジェクトを使用して実行
            Dim ret As DbDataReader = m_cmd.ExecuteReader()

            '実行後パラメーターをクリアする
            Me.ClearParameters()

            Return ret
        End Function

        ''' <summary>SQLステートメントを実行し, DataTable を取得します</summary>
        ''' <param name="sql">実行するSQLステートメント</param>
        ''' <param name="dt">レコードを格納するための DataTable</param>
        ''' <returns>正常に追加された行数</returns>
        Public Overloads Function Fill(ByVal sql As String, ByVal dt As DataTable) As Integer
            Dim ret As Integer = 0          '戻り値(正常に追加された行数)

            'Debug.Print(DateTime.Now.ToString("HH:mm:ss,fff") & " " & sql)

            'コマンドオブジェクトの生成と実行
            Using da As DbDataAdapter = m_dbFactory.CreateDataAdapter(), _
                cmd As DbCommand = m_dbFactory.CreateCommand()

                Me.SetCommand(cmd, sql)
                da.SelectCommand = cmd
                ret = da.Fill(dt)
            End Using

            '実行後パラメーターをクリアする.
            Me.ClearParameters()

            Return ret
        End Function

        ''' <summary>SQLステートメントを実行し, DataSet を取得します</summary>
        ''' <param name="sql">実行するSQLステートメント</param>
        ''' <param name="ds">レコードを格納するための DataSet</param>
        ''' <param name="srcTable">テーブルマップに使用するソーステーブルの名前</param>
        ''' <returns>正常に追加された行数</returns>
        Public Overloads Function Fill(ByVal sql As String, ByVal ds As DataSet, ByVal srcTable As String) As Integer
            Dim ret As Integer = 0          '戻り値(正常に追加された行数)

            'Debug.Print(DateTime.Now.ToString("HH:mm:ss,fff") & " " & sql)

            'コマンドオブジェクトの生成と実行
            Using da As DbDataAdapter = m_dbFactory.CreateDataAdapter(), _
                cmd As DbCommand = m_dbFactory.CreateCommand()

                Me.SetCommand(cmd, sql)
                da.SelectCommand = cmd
                ret = da.Fill(ds, srcTable)
            End Using

            '実行後パラメーターをクリアする.
            Me.ClearParameters()

            Return ret
        End Function

        ''' <summary>コマンドオブジェクトを設定します</summary>
        ''' <param name="cmd">設定するコマンドオブジェクト</param>
        ''' <param name="sql">コマンドオブジェクトに設定するSQLステートメント</param>
        Protected Sub SetCommand(ByVal cmd As DbCommand, ByVal sql As String)
            cmd.Connection = m_cn
            cmd.CommandText = sql

            'コマンドタイムアウトが設定されている場合,コマンドタイムアウト値をセット
            If m_cmdTimeOut > 0 Then
                cmd.CommandTimeout = m_cmdTimeOut
            End If

            'パラメーターの追加
            cmd.Parameters.AddRange(m_dbParams.ToArray)

            _LastSql = sql
            _LastParams = m_dbParams.ToArray

            'トランザクション中の場合,トランザクションオブジェクトをセット
            If HasTransaction Then
                cmd.Transaction = m_trans
            End If
        End Sub

        ''' <summary>パラメーターを追加します</summary>
        ''' <param name="paramName">パラメーター名</param>
        ''' <param name="value">値</param>
        ''' <remarks>
        ''' ここで追加されたパラメーターは, コマンドオブジェクト生成時に使用されます.
        ''' なお,SQLステートメント実行後,パラメーターは自動的に削除されます.
        ''' </remarks>
        Public Overloads Sub AddParameter(ByVal paramName As String, ByVal value As Object, _
                                          Optional ByVal direction As ParameterDirection = ParameterDirection.Input, Optional ByVal size As Integer = 0)
            ' TODO 2014.11.10 こうしたい!!
            'AddParameter(paramName, value, dbtype:=Nothing, direction:=direction, size:=size)
            Dim dbParam As Common.DbParameter = m_dbFactory.CreateParameter()  'パラメーター生成

            'Debug.Print(DateTime.Now.ToString("HH:mm:ss,fff") & " @" & paramName & "=" & StringUtil.Nvl(value))

            '値が Nothing なら DBNull.Value
            If value Is Nothing Then value = DBNull.Value

            dbParam.ParameterName = paramName       'パラメーター名
            dbParam.Value = value                   '値
            dbParam.Direction = direction
            If 0 < size Then
                dbParam.Size = size
            End If

            'パラメーター配列に追加
            m_dbParams.Add(dbParam)
        End Sub

        ''' <summary>パラメーターを追加します</summary>
        ''' <param name="paramName">パラメーター名</param>
        ''' <param name="value">値</param>
        ''' <remarks>
        ''' ここで追加されたパラメーターは, コマンドオブジェクト生成時に使用されます.
        ''' なお,SQLステートメント実行後,パラメーターは自動的に削除されます.
        ''' </remarks>
        Public Overloads Sub AddParameter(ByVal paramName As String, ByVal value As Object, ByVal dbtype As DbType?, _
                                          Optional ByVal direction As ParameterDirection = ParameterDirection.Input, Optional ByVal size As Integer = 0)
            Dim dbParam As Common.DbParameter = m_dbFactory.CreateParameter()

            'Debug.Print(DateTime.Now.ToString("HH:mm:ss,fff") & " @" & paramName & "=" & StringUtil.Nvl(value))

            '値が Nothing なら DBNull.Value
            If value Is Nothing Then value = DBNull.Value

            dbParam.ParameterName = paramName       'パラメーター名
            dbParam.Value = value                   '値
            If dbtype.HasValue Then
                dbParam.DbType = dbtype.Value       'データベースの型
            End If
            dbParam.Direction = direction
            If 0 < size Then
                dbParam.Size = size
            End If

            'パラメーター配列に追加
            m_dbParams.Add(dbParam)
        End Sub

        ''' <summary>追加されているパラメーターをすべて削除します</summary>
        Public Sub ClearParameters()
            m_dbParams.Clear()
        End Sub

        ''' <summary>
        ''' パラメーターデバッグ表示
        ''' </summary>
        ''' <remarks></remarks>
        Public Function DebugParameter() As String
            Dim ret As New StringBuilder()

            For Each param As Common.DbParameter In m_dbParams
                ret.AppendLine(String.Format("{0}  =>  '{1}'  Type: {2}", _
                                             param.ParameterName, param.Value, param.DbType.ToString()))
            Next

            Return ret.ToString()
        End Function

        ''' <summary>接続を開くために使用する文字列を取得または設定します</summary>
        Public WriteOnly Property ConnectionString() As String
            Set(ByVal value As String)
                m_connectionString = value
                m_cn.ConnectionString = value
            End Set
        End Property

        ''' <summary>試行を中断してエラー生成するまでの接続の確立時に待機する時間を取得します</summary>
        Public ReadOnly Property ConnectionTimeOut() As Integer
            Get
                Return m_cn.ConnectionTimeout
            End Get
        End Property

        ''' <summary>接続が開いてから現在のデータベースの名前を取得するか、接続が開く前に接続文字列に指定されたデータベース名を取得します</summary>
        Public ReadOnly Property Database() As String
            Get
                Return m_cn.Database
            End Get
        End Property

        ''' <summary>接続するデータベースサーバーの名前を取得します</summary>
        Public ReadOnly Property DataSource() As String
            Get
                Return m_cn.DataSource
            End Get
        End Property

        ''' <summary>接続しているサーバーのバージョンを表す文字列を取得します</summary>
        Public ReadOnly Property ServerVersion() As String
            Get
                Return m_cn.ServerVersion
            End Get
        End Property

        ''' <summary>SQLコマンドタイムアウト値(秒)を取得または設定します</summary>
        Public Property CommandTimeOut() As Integer
            Get
                Return m_cmdTimeOut
            End Get
            Set(ByVal value As Integer)
                m_cmdTimeOut = value
            End Set
        End Property

        ''' <summary>接続の状態を取得します</summary>
        Public ReadOnly Property State() As ConnectionState
            Get
                Return m_cn.State
            End Get
        End Property

        ''' <summary>トランザクション中であるかの状態を取得します</summary>
        Public ReadOnly Property HasTransaction() As Boolean
            Get
                Return m_trans IsNot Nothing AndAlso m_trans.Connection IsNot Nothing
            End Get
        End Property

        Private disposedValue As Boolean = False        ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Me.disposedValue Then
                Return
            End If

            If disposing Then
                ' 明示的に呼び出されたときにマネージ リソースを解放します
            End If

            ' 共有のアンマネージ リソースを解放します

            'トランザクションオブジェクト
            If Not m_trans Is Nothing Then
                If Not m_trans.Connection Is Nothing Then
                    m_trans.Rollback()
                End If

                m_trans.Dispose()
            End If

            'コマンドオブジェクト
            If Not m_cmd Is Nothing Then
                m_cmd.Dispose()
            End If

            'コネクションオブジェクト
            If Not m_cn Is Nothing Then
                If Not m_cn.State = ConnectionState.Closed Then
                    Try
                        m_cn.Close()

                    Catch ex As InvalidOperationException
                        If "PrePush".Equals(ex.TargetSite.Name) _
                                AndAlso "System.Data.ProviderBase.DbConnectionInternal".Equals(ex.TargetSite.DeclaringType.FullName) Then
                            ' DB接続時にコネクションプーリングを初期化していて、
                            ' その時にアプリ終了(ALT+F4)すると当Disposeが動いてm_cn.Close()のときに中断され例外になる模様
                            ' 再現手順：
                            '   1. アプリ起動
                            '   2. Windowが出たらすぐにALT+F4で終了
                            '   3. 以上（例外が発生しなければ数回繰り返してみる）
                            ' この時の例外と考えられる例外は無視する
                            Return
                        End If
                        Throw
                    End Try
                End If

                m_cn.Dispose()
            End If

            Me.disposedValue = True
        End Sub

        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Private _LastSql As String
        ''' <summary>最後に実行したSQL</summary>
        Friend ReadOnly Property LastSql() As String
            Get
                Return _LastSql
            End Get
        End Property

        Private _LastParams As Common.DbParameter()
        ''' <summary>最後に実行したSQLのパラメータ</summary>
        Friend ReadOnly Property LastParams() As Common.DbParameter()
            Get
                Return _LastParams
            End Get
        End Property

    End Class
End Namespace