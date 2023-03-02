Imports System.Collections.Generic
Imports System.Linq.Expressions
Imports Fhi.Fw.Lang.Threading
Imports Fhi.Fw.Util
Imports System.Threading

Namespace Db

    ''' <summary>
    ''' DB操作を担うクラス
    ''' </summary>
    ''' <remarks>DbAccessを使い易くしたクラス</remarks>
    Public MustInherit Class DbClient : Implements IDisposable

        'Protected Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

        Private ReadOnly connectString As String
        Private ReadOnly provider As DbProvider
        ''' <summary>DBコネクションを管理中か</summary>
        Private hasManagingDb As Boolean

#Region "Nested classes..."
        Private Class ProcedureParam
            Public ReadOnly ParamName As String
            Public ReadOnly Value As Object
            Public ReadOnly Direction As ParameterDirection
            Public ReadOnly Size As Integer
            Public Sub New(ByVal paramName As String, ByVal value As Object, ByVal direction As ParameterDirection, ByVal size As Integer)
                Me.ParamName = paramName
                Me.Value = value
                Me.Direction = direction
                Me.Size = size
            End Sub
        End Class
        ''' <summary>
        ''' スレッドごとに一意となるDbクラスとそれに伴うプロパティ値を担うクラス
        ''' </summary>
        ''' <remarks></remarks>
        Private Class DbBag

            ''' <summary>当スレッドで一意のデータベースインスタンス</summary>
            Public Db As DbAccess

            ''' <summary>リンクサーバー経由の場合、リンクサーバー名を設定する</summary>
            Public LinkServerName As String

            ''' <summary>埋め込みパラメータ使用を許可しない場合、true</summary>
            Public DenyBindParameter As Boolean

            ''' <summary>強制的に項目値をTrimする場合、true</summary>
            Public TrimsEndOfValueToForce As Boolean

            ''' <summary>更新を許可するデータベースの場合、true</summary>
            Public AllowUpdate As Boolean

            ''' <summary>OPEN QUERYを使用しない場合、true ※初期値=false</summary>
            Public DenyOpenQuery As Boolean = False

        End Class
#End Region
#Region "Public properties..."
        ''' <summary>ロックタイムアウト(ms) (-1は無期限)</summary>
        Private _LockTimeout As Long

        ''' <summary>キャメライズ処理で一致しないDB項目名の個別設定 key:プロパティ名 / value:DB項目名</summary>
        Private ReadOnly specifiedUnableCamelizeNames As Dictionary(Of String, String)

        ''' <summary>更新を許可するデータベースの場合、true ※初期値=false</summary>
        Public _allowUpdate As Boolean = False
        Public Property AllowUpdate() As Boolean
            Get
                If ExistsDbUnitOnThread() Then
                    Return GetDbUnitOnThread.AllowUpdate
                End If
                Return _allowUpdate
            End Get
            Set(ByVal value As Boolean)
                _allowUpdate = value
                If ExistsDbUnitOnThread() Then
                    GetDbUnitOnThread.AllowUpdate = value
                End If
            End Set
        End Property

        Private _linkServerName As String
        ''' <summary>リンクサーバー経由の場合、リンクサーバー名を設定する</summary>
        Public Property LinkServerName() As String
            Get
                If ExistsDbUnitOnThread() Then
                    Return GetDbUnitOnThread.LinkServerName
                End If
                Return _linkServerName
            End Get
            Set(ByVal value As String)
                _linkServerName = value
                If ExistsDbUnitOnThread() Then
                    GetDbUnitOnThread.LinkServerName = value
                End If
                If Not String.IsNullOrEmpty(value) Then
                    TrimsEndOfValueToForce = True
                End If
            End Set
        End Property

        Private _trimsEndOfValueToForce As Boolean
        ''' <summary>項目値の末尾を強制的にTrimする場合、true ※初期値=false</summary>
        Public Property TrimsEndOfValueToForce() As Boolean
            Get
                If ExistsDbUnitOnThread() Then
                    Return GetDbUnitOnThread.TrimsEndOfValueToForce
                End If
                Return _trimsEndOfValueToForce
            End Get
            Set(ByVal value As Boolean)
                _trimsEndOfValueToForce = value
                If ExistsDbUnitOnThread() Then
                    GetDbUnitOnThread.TrimsEndOfValueToForce = value
                End If
            End Set
        End Property

        Private _denyBindParameter As Boolean = False
        ''' <summary>埋め込みパラメータ使用を許可しない場合、true ※初期値=false</summary>
        Public Property DenyBindParameter() As Boolean
            Get
                If ExistsDbUnitOnThread() Then
                    Return GetDbUnitOnThread.DenyBindParameter
                End If
                Return _denyBindParameter
            End Get
            Set(ByVal value As Boolean)
                _denyBindParameter = value
                If ExistsDbUnitOnThread() Then
                    GetDbUnitOnThread.DenyBindParameter = value
                End If
            End Set
        End Property

        ''' <summary>分離レベル</summary>
        Public IsolationLevel As IsolationLevel = Data.IsolationLevel.ReadCommitted

        ''' <summary>ロックタイムアウト(ms) (-1は無期限)</summary>
        Public Property LockTimeout() As Long
            Get
                Return _LockTimeout
            End Get
            Set(ByVal value As Long)
                _LockTimeout = value
                If ExistsDbUnitOnThread() AndAlso hasManagingDb Then
                    provider.SetLockTimeoutMillis(GetDbUnitOnThread().Db, _LockTimeout)
                End If
            End Set
        End Property

        ''' <summary>
        ''' トランザクション管理開始中かを返す
        ''' </summary>
        ''' <returns>開始中の場合、true</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IsBeginningTransaction() As Boolean
            Get
                If ExistsDbUnitOnThread() Then
                    Return GetDbUnitOnThread().Db.HasTransaction
                End If
                Return False
            End Get
        End Property

        Private _denyOpenQuery As Boolean
        ''' <summary>OPENQUERYを使用しない場合、true</summary>
        Public Property DenyOpenQuery() As Boolean
            Get
                If ExistsDbUnitOnThread() Then
                    Return GetDbUnitOnThread.DenyOpenQuery
                End If
                Return _denyOpenQuery
            End Get
            Set(ByVal value As Boolean)
                _denyOpenQuery = value
                If ExistsDbUnitOnThread() Then
                    GetDbUnitOnThread.DenyOpenQuery = value
                End If
            End Set
        End Property

        ''' <summary>SQLタイムアウト(s)</summary>
        Public Property CommandTimeout As Integer

#End Region
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="connectString">コネクション文字列</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal provider As DbProvider, ByVal connectString As String)
            Me.New(provider, connectString, Nothing)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="connectString">コネクション文字列</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal provider As DbProvider, ByVal connectString As String, ByVal dbFieldNameByPropertyName As Dictionary(Of String, String))
            Me.provider = provider
            Me._LockTimeout = provider.DefaultLockTimeoutMillis
            Me.connectString = connectString
            Me.specifiedUnableCamelizeNames = dbFieldNameByPropertyName
            _commandTimeout = provider.DefaultCommandTimeout
        End Sub

        Private Sub InitializeDb(ByVal db As DbAccess)
            provider.SetLockTimeoutMillis(db, _LockTimeout)
        End Sub

        ''' <summary>
        ''' DbBag を生成して返す
        ''' </summary>
        ''' <returns>新しい DbBag</returns>
        ''' <remarks></remarks>
        Private Function NewDbBag() As DbBag
            Return New DbBag With {.Db = NewDbAccess(), .AllowUpdate = AllowUpdate, _
                                   .DenyBindParameter = DenyBindParameter, .DenyOpenQuery = DenyOpenQuery, _
                                   .LinkServerName = LinkServerName, .TrimsEndOfValueToForce = TrimsEndOfValueToForce}
        End Function

        ''' <summary>
        ''' DbAccess を生成して返す
        ''' </summary>
        ''' <returns>新しい DbAccess</returns>
        ''' <remarks></remarks>
        Private Function NewDbAccess() As DbAccess
            Return NewDbAccess(False)
        End Function
        ''' <summary>
        ''' DbAccess を生成して返す
        ''' </summary>
        ''' <param name="noInitialize">DBの初期化をしない場合、true</param>
        ''' <returns>新しい DbAccess</returns>
        ''' <remarks></remarks>
        Private Function NewDbAccess(ByVal noInitialize As Boolean) As DbAccess
            Dim result As DbAccess = New DbAccess(provider, connectString) With {.CommandTimeOut = CommandTimeout}
            result.Open()
            If Not noInitialize Then
                InitializeDb(result)
            End If
            Return result
        End Function

        ''' <summary>
        ''' Select文を実行した結果を同名プロパティに保持した Object にして返す
        ''' </summary>
        ''' <typeparam name="T">結果の型</typeparam>
        ''' <param name="sql">SELECT文</param>
        ''' <returns>結果を同名プロパティに保持した Object</returns>
        ''' <remarks></remarks>
        Public Function QueryForObject(Of T)(ByVal sql As String) As T
            Return QueryForObject(Of T)(sql, Nothing)
        End Function
        ''' <summary>
        ''' Select文を実行した結果を同名プロパティに保持した Object にして返す
        ''' </summary>
        ''' <typeparam name="T">結果の型</typeparam>
        ''' <param name="sql">SELECT文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <returns>結果を同名プロパティに保持した Object</returns>
        ''' <remarks></remarks>
        Public Function QueryForObject(Of T)(ByVal sql As String, ByVal parameter As Object) As T
            Dim results As List(Of T) = QueryForList(Of T)(sql, parameter)
            If 0 < results.Count Then
                Return results(0)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Select文を実行した結果を同名プロパティに保持した Object の List にして返す
        ''' </summary>
        ''' <typeparam name="T">結果の型</typeparam>
        ''' <param name="sql">SELECT文</param>
        ''' <returns>結果を同名プロパティに保持した Object の List</returns>
        ''' <remarks></remarks>
        Public Function QueryForList(Of T)(ByVal sql As String) As List(Of T)
            Return QueryForList(Of T)(sql, Nothing)
        End Function

        ''' <summary>
        ''' Select文を実行した結果を同名プロパティに保持した Object の List にして返す
        ''' </summary>
        ''' <typeparam name="T">結果の型</typeparam>
        ''' <param name="sql">SELECT文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <returns>結果を同名プロパティに保持した Object の List</returns>
        ''' <remarks></remarks>
        Public Function QueryForList(Of T)(ByVal sql As String, ByVal parameter As Object) As List(Of T)
            Return QueryForList(Of T)(sql, parameter, InternalGetUnableCamelizeNames(Of T))
        End Function

        ''' <summary>
        ''' Select文を実行した結果を同名プロパティに保持した Object の List にして返す
        ''' </summary>
        ''' <typeparam name="T">結果の型</typeparam>
        ''' <param name="sql">SELECT文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <returns>結果を同名プロパティに保持した Object の List</returns>
        ''' <remarks></remarks>
        Public Function QueryForList(Of T)(ByVal sql As String, ByVal parameter As Object, ByVal camelizeIrregulars As Dictionary(Of String, String)) As List(Of T)
            If ExistsDbUnitOnThread() Then
                GetDbUnitOnThread().Db.ClearParameters()
                Return QueryForList(Of T)(sql, parameter, camelizeIrregulars, GetDbUnitOnThread().Db)
            End If
            Using db As DbAccess = NewDbAccess()
                Return QueryForList(Of T)(sql, parameter, camelizeIrregulars, db)
            End Using
        End Function

        ''' <summary>
        ''' Select文を実行した結果を同名プロパティに保持した Object の List にして返す
        ''' </summary>
        ''' <typeparam name="T">結果の型</typeparam>
        ''' <param name="sql">SELECT文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <param name="camelizeIrregulars">Open中の DbAccess</param>
        ''' <param name="db">Open中の DbAccess</param>
        ''' <returns>結果を同名プロパティに保持した Object の List</returns>
        ''' <remarks></remarks>
        Private Function QueryForList(Of T)(ByVal sql As String, ByVal parameter As Object, ByVal camelizeIrregulars As Dictionary(Of String, String), ByVal db As DbAccess) As List(Of T)
            Return CreateDbAccessHelper(db).QueryForList(Of T)(sql, parameter, camelizeIrregulars)
        End Function

        Private Function CreateDbAccessHelper(ByVal db As DbAccess) As DbAccessHelper
            Dim helper As New DbAccessHelper(db)
            helper.LinkServerName = LinkServerName
            helper.DenyBindParameter = DenyBindParameter
            helper.TrimsEndOfValueToForce = TrimsEndOfValueToForce
            helper.DenyOpenQuery = DenyOpenQuery
            Return helper
        End Function

        ''' <summary>
        ''' Insert処理を行う
        ''' </summary>
        ''' <param name="sql">Insert文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <returns>処理件数</returns>
        ''' <remarks></remarks>
        Public Function Insert(ByVal sql As String, Optional ByVal parameter As Object = Nothing) As Integer
            Return Execute(sql, parameter)
        End Function

        ''' <summary>
        ''' Update処理を行う
        ''' </summary>
        ''' <param name="sql">Update文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <returns>処理件数</returns>
        ''' <remarks></remarks>
        Public Function Update(ByVal sql As String, Optional ByVal parameter As Object = Nothing) As Integer
            Return Execute(sql, parameter)
        End Function

        ''' <summary>リンクサーバーのDelete文を許可する場合、true</summary>
        Friend AllowDeleteFromLinkServer As Boolean
        ''' <summary>
        ''' Delete処理を行う
        ''' </summary>
        ''' <param name="sql">Delete文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <returns>処理件数</returns>
        ''' <remarks></remarks>
        Public Function Delete(ByVal sql As String, Optional ByVal parameter As Object = Nothing) As Integer
            If Not AllowDeleteFromLinkServer AndAlso StringUtil.IsNotEmpty(LinkServerName) Then
                Throw New InvalidOperationException(String.Format("リンクサーバー {0} に対して DeleteByPk 以外は使用禁止.", LinkServerName))
            End If
            Return Execute(sql, parameter)
        End Function

        ''' <summary>
        ''' 更新許可されたデータベースであることを保証する
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub AssertAllowUpdate()
            If Not AllowUpdate Then
                Throw New DatabaseException(String.Format("更新許可されたデータベースではありません. {0} - {1}", ConnectionUtil.GetServer(connectString), ConnectionUtil.GetDatabaseName(connectString)))
            End If
        End Sub

        ''' <summary>
        ''' SQL処理を実行する
        ''' </summary>
        ''' <param name="sql">SQL文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <returns>実行結果</returns>
        ''' <remarks></remarks>
        Private Function Execute(ByVal sql As String, ByVal parameter As Object) As Integer
            AssertAllowUpdate()
            If ExistsDbUnitOnThread() Then
                GetDbUnitOnThread().Db.ClearParameters()
                Return Execute(sql, parameter, GetDbUnitOnThread().Db)
            End If
            Using db As DbAccess = NewDbAccess()
                Return Execute(sql, parameter, db)
            End Using
        End Function

        ''' <summary>
        ''' SQL処理を実行する
        ''' </summary>
        ''' <param name="sql">SQL文</param>
        ''' <param name="parameter">埋め込み値をもつオブジェクト</param>
        ''' <param name="db">処理の実行先（Open中の DbAccess）</param>
        ''' <returns>実行結果</returns>
        ''' <remarks></remarks>
        Private Function Execute(ByVal sql As String, ByVal parameter As Object, ByVal db As DbAccess) As Integer
            Return CreateDbAccessHelper(db).Execute(sql, parameter)
        End Function

        Private ReadOnly procedureParams As New List(Of ProcedureParam)

        ''' <summary>
        ''' プロシージャ実行のパラメータを追加する
        ''' </summary>
        ''' <param name="paramName">パラメータ名</param>
        ''' <param name="value">値</param>
        ''' <param name="direction">パラメータの型</param>
        ''' <param name="size">桁数</param>
        ''' <remarks></remarks>
        Public Sub AddProcedureParameter(ByVal paramName As String, ByVal value As Object, _
                                         Optional ByVal direction As ParameterDirection = ParameterDirection.Input, Optional ByVal size As Integer = 0)
            procedureParams.Add(New ProcedureParam(paramName, value, direction, size))
        End Sub

        ''' <summary>
        ''' プロシージャを実行する
        ''' </summary>
        ''' <param name="procedureName">プロシージャ名</param>
        ''' <returns>結果[]</returns>
        ''' <remarks></remarks>
        Public Function ExecuteProcedure(ByVal procedureName As String) As List(Of Object)
            Try
                If ExistsDbUnitOnThread() Then
                    GetDbUnitOnThread().Db.ClearParameters()
                    Return PerformExecuteProcedure(procedureName, procedureParams, GetDbUnitOnThread().Db)
                End If
                Using db As DbAccess = NewDbAccess()
                    Return PerformExecuteProcedure(procedureName, procedureParams, db)
                End Using

            Finally
                procedureParams.Clear()
            End Try
        End Function

        ''' <summary>
        ''' プロシージャを実行する（本体）
        ''' </summary>
        ''' <param name="procedureName">プロシージャ名</param>
        ''' <param name="params">引数[]</param>
        ''' <param name="db">DbAccess</param>
        ''' <returns>結果[]</returns>
        ''' <remarks></remarks>
        Private Function PerformExecuteProcedure(ByVal procedureName As String, ByVal params As IEnumerable(Of ProcedureParam), ByVal db As DbAccess) As List(Of Object)
            Dim helper As DbAccessHelper = CreateDbAccessHelper(db)
            For Each param As ProcedureParam In params
                helper.AddProcedureParameter(param.ParamName, param.Value, param.Direction, param.Size)
            Next
            Return helper.ExecuteProcedure(procedureName)
        End Function

        Private Shared ReadOnly txPerConectStr As New Fhi.Fw.Lang.Threading.ThreadLocal(Of Dictionary(Of String, DbBag))(Function() New Dictionary(Of String, DbBag))

        Private Sub AddDbUnitOnThread(ByVal db As DbBag)
            txPerConectStr.Get().Add(connectString, db)
        End Sub

        Private Sub RemoveDbUnitOnThread()
            txPerConectStr.Get().Remove(connectString)
        End Sub

        Private Function ExistsDbUnitOnThread() As Boolean
            Return txPerConectStr.Get().ContainsKey(connectString)
        End Function

        Private Function GetDbUnitOnThread() As DbBag
            Return txPerConectStr.Get()(connectString)
        End Function

        Private escapeDb As DbBag
        Private Sub SuspendDbOnThread()
            escapeDb = GetDbUnitOnThread()
            RemoveDbUnitOnThread()
        End Sub
        Private Function IsSuspend() As Boolean
            Return escapeDb IsNot Nothing
        End Function
        Private Sub ResumeDbOnThread()
            AddDbUnitOnThread(escapeDb)
        End Sub

        ''' <summary>
        ''' トランザクション管理を開始する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub BeginTransaction()
            BeginTransaction(Me.IsolationLevel)
        End Sub

        ''' <summary>
        ''' トランザクション管理を開始する
        ''' </summary>
        ''' <param name="anIsolationLevel">分離レベル</param>
        ''' <remarks></remarks>
        Public Sub BeginTransaction(ByVal anIsolationLevel As IsolationLevel)
            If ExistsDbUnitOnThread() Then
                If hasManagingDb Then
                    DbAccessHelper.logDebug(GetDbUnitOnThread().Db.ToString & " - BeginTransaction")
                    GetDbUnitOnThread().Db.BeginTransaction(anIsolationLevel)
                    Return
                Else
                    SuspendDbOnThread()
                End If
            End If
            '' DbAccessを作成して、BeginTransactionして、poolに格納して、返す
            Dim bag As DbBag = NewDbBag()
            AddDbUnitOnThread(bag)
            DbAccessHelper.logDebug(bag.Db.ToString & " - BeginTransaction")
            bag.Db.BeginTransaction(anIsolationLevel)
            hasManagingDb = True
        End Sub
        ''' <summary>
        ''' コミットする
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Commit()
            If ExistsDbUnitOnThread() AndAlso hasManagingDb Then
                DbAccessHelper.logDebug(GetDbUnitOnThread().Db.ToString & " - Commit")
                GetDbUnitOnThread().Db.Commit()
            End If
        End Sub
        ''' <summary>
        ''' ロールバックする
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Rollback()
            If ExistsDbUnitOnThread() AndAlso hasManagingDb Then
                DbAccessHelper.logDebug(GetDbUnitOnThread().Db.ToString & " - Rollback")
                GetDbUnitOnThread().Db.Rollback()
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                If ExistsDbUnitOnThread() AndAlso hasManagingDb Then
                    Rollback()
                    Dim db As DbAccess = GetDbUnitOnThread().Db
                    RemoveDbUnitOnThread()
                    DbAccessHelper.logDebug(db.ToString & " - Closing")
                    db.Close()
                End If
                GC.SuppressFinalize(Me)
            Finally
                If IsSuspend() Then
                    ResumeDbOnThread()
                End If
            End Try
        End Sub

        ''' <summary>
        ''' キャメライズ通りの対応が出来ない項目を、項目名に紐付ける
        ''' </summary>
        ''' <remarks></remarks>
        Protected Interface BindTable
            ''' <summary>
            ''' テーブルVOを宣言する
            ''' </summary>
            ''' <param name="aTableVo">テーブルVO</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function IsA(ByVal aTableVo As Object) As BindFieldAndName
        End Interface

        ''' <summary>
        ''' 項目＆項目名を設定する
        ''' </summary>
        ''' <remarks></remarks>
        Public Interface BindFieldAndName
            ''' <summary>
            ''' 項目と項目名を宣言する
            ''' </summary>
            ''' <param name="field">項目プロパティ</param>
            ''' <param name="fieldName">項目名</param>
            ''' <returns>項目＆項目名設定インターフェース</returns>
            ''' <remarks></remarks>
            Function Bind(ByVal field As Object, ByVal fieldName As String) As BindFieldAndName
        End Interface

        Private Class BindTableImpl : Implements BindTable
            Private ReadOnly marker As New VoPropertyMarker
            Public ReadOnly dbNameByPropertyNameByType As New Dictionary(Of Type, Dictionary(Of String, String))
            Private aType As Type
            Public Function IsA(ByVal aTableVo As Object) As BindFieldAndName Implements BindTable.IsA
                marker.Clear()
                marker.MarkVo(aTableVo)
                aType = aTableVo.GetType
                If Not dbNameByPropertyNameByType.ContainsKey(aType) Then
                    dbNameByPropertyNameByType.Add(aType, New Dictionary(Of String, String))
                End If
                Return New BindFieldAndNameImpl(Me)
            End Function
            Friend Sub Message(ByVal field As Object, ByVal fieldName As String)
                dbNameByPropertyNameByType(aType).Add(marker.GetPropertyInfo(field).Name, fieldName)
            End Sub
            Friend Sub Message(Of T)(ByVal field As Expression(Of Func(Of T)), ByVal fieldName As String)
                dbNameByPropertyNameByType(aType).Add(marker.GetPropertyInfo(field).Name, fieldName)
            End Sub
        End Class

        Private Class BindFieldAndNameImpl : Implements BindFieldAndName
            Private ReadOnly binder As BindTableImpl
            Public Sub New(ByVal binder As BindTableImpl)
                Me.binder = binder
            End Sub
            Public Function Bind(ByVal field As Object, ByVal fieldName As String) As BindFieldAndName Implements BindFieldAndName.Bind
                binder.Message(field, fieldName)
                Return Me
            End Function
        End Class

        Private Class InternalWildcard
        End Class
        Private Shared ReadOnly TypeOfWildcard As Type = GetType(InternalWildcard)
        Private dbNameByPropertyNameByType As Dictionary(Of Type, Dictionary(Of String, String))
        Private ReadOnly rwLock As New ReaderWriterLockSlim
        ''' <summary>
        ''' プロパティ名＝DB項目名がキャメライズ通りに成立しない項目名を取得する
        ''' </summary>
        ''' <returns>key:プロパティ名 / value:DB項目名</returns>
        ''' <remarks></remarks>
        Friend Function InternalGetUnableCamelizeNames(Of T)() As Dictionary(Of String, String)
            If dbNameByPropertyNameByType Is Nothing Then
                rwLock.EnterWriteLock()
                Try
                    If dbNameByPropertyNameByType Is Nothing Then
                        Dim binder As New BindTableImpl
                        SettingFieldNameCamelizeIrregular(binder)
                        dbNameByPropertyNameByType = New Dictionary(Of Type, Dictionary(Of String, String))(binder.dbNameByPropertyNameByType)
                        If specifiedUnableCamelizeNames IsNot Nothing Then
                            For Each dbByProperty As Dictionary(Of String, String) In dbNameByPropertyNameByType.Values
                                For Each key As String In specifiedUnableCamelizeNames.Keys
                                    If dbByProperty.ContainsKey(key) Then
                                        Continue For
                                    End If
                                    dbByProperty.Add(key, specifiedUnableCamelizeNames(key))
                                Next
                            Next
                            dbNameByPropertyNameByType.Add(TypeOfWildcard, specifiedUnableCamelizeNames)
                        Else
                            dbNameByPropertyNameByType.Add(TypeOfWildcard, New Dictionary(Of String, String))
                        End If
                    End If
                Finally
                    rwLock.ExitWriteLock()
                End Try
            End If
            Dim aType As Type = GetType(T)
            Dim result As New Dictionary(Of String, String)
            RecurCreateUnableCamelizeNames(aType, result)
            If result.Count = 0 Then
                Return dbNameByPropertyNameByType(TypeOfWildcard)
            End If
            Return result
        End Function

        Private Sub RecurCreateUnableCamelizeNames(ByVal aType As Type, ByVal unableCamelizeNames As Dictionary(Of String, String))
            If aType Is TypeUtil.TypeObject Then
                Return
            End If
            If dbNameByPropertyNameByType.ContainsKey(aType) Then
                For Each value As KeyValuePair(Of String, String) In dbNameByPropertyNameByType(aType)
                    If Not unableCamelizeNames.ContainsKey(value.Key) Then
                        unableCamelizeNames.Add(value.Key, value.Value)
                    End If
                Next
            End If
            RecurCreateUnableCamelizeNames(aType.BaseType, unableCamelizeNames)
        End Sub

        ''' <summary>
        ''' プロパティ名＝DB項目名がキャメライズ通りに成立しない場合、設定する
        ''' </summary>
        ''' <param name="bind">項目と項目名を設定する</param>
        ''' <remarks></remarks>
        Protected Overridable Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)

        End Sub

        ''' <summary>
        ''' DB接続できることを保障する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub AssertDbConnect()
            Using db As New DbAccess(provider, connectString) With {.CommandTimeOut = CommandTimeout}
                db.Open()
            End Using
        End Sub

        ''' <summary>
        ''' 複数行一括Insertできることを保証する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub AssertCanMultiValuesInsert()
            provider.AssertCanMultiValuesInsert()
        End Sub

        ''' <summary>
        ''' 行ロック参照SQLを作成する
        ''' </summary>
        ''' <param name="select">Select句</param>
        ''' <param name="from">From句</param>
        ''' <param name="where">Where句</param>
        ''' <returns>行ロック参照SQL</returns>
        ''' <remarks></remarks>
        Friend Function MakeSelectForUpdate(ByVal [select] As String, ByVal from As String, ByVal where As String) As String
            Return provider.MakeSelectForUpdate([select], from, where)
        End Function

        ''' <summary>
        ''' DB接続を開きっぱなしにする ※ セットで Me#Dispose() が必要（あるいはUsing句）
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub BeginToKeepOpen()
            If ExistsDbUnitOnThread() Then
                If hasManagingDb Then
                    Return
                Else
                    SuspendDbOnThread()
                End If
            End If
            '' DbAccessを作成して、BeginTransactionして、poolに格納して、返す
            Dim bag As DbBag = NewDbBag()
            AddDbUnitOnThread(bag)
            hasManagingDb = True
            DbAccessHelper.logDebug(bag.Db.ToString & " - Opened")
        End Sub

        ''' <summary>
        ''' テーブルの列名を取得する
        ''' </summary>
        ''' <typeparam name="T">取得したいテーブルのVo/Daoの型</typeparam>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FetchTableColumns(Of T)() As String()
            Using db As DbAccess = NewDbAccess()
                Return CreateDbAccessHelper(db).FetchTableColumns(Of T)()
            End Using
        End Function
    End Class

End Namespace
