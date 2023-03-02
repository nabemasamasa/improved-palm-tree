Imports System.Reflection
Imports System.Linq.Expressions
Imports Fhi.Fw.Util
Imports System.Text

Namespace Db.Impl
    ''' <summary>
    ''' 1テーブルのためのCRUDを提供するDAO
    ''' </summary>
    ''' <typeparam name="T">テーブルに対応するVO</typeparam>
    ''' <remarks>実装クラスは、[テーブル名]＆"Dao" という命名規則に従う事</remarks>
    Public MustInherit Class DaoEachTableImpl(Of T) : Implements DaoEachTable(Of T)

        Private ForUpdate As Boolean

        ''' <summary>
        ''' 新しい DbClient のインスタンスを生成して返す
        ''' </summary>
        ''' <returns>新しい DbClient のインスタンス</returns>
        ''' <remarks></remarks>
        Protected Function NewDbClient() As DbClient
            Return NewDbClient(Nothing)
        End Function

        Private Shared Sub AssertParameterIsNotNull(ByVal parameter As Object, ByVal parameterName As String)
            If parameter Is Nothing Then
                Throw New ArgumentException("No specified " & parameterName & ".")
            End If
        End Sub

        Protected Overridable Function GetTableName() As String
            Dim suffixLength As Integer
            If Me.GetType.Name.EndsWith("Dao") Then
                suffixLength = 3
            ElseIf Me.GetType.Name.EndsWith("DaoImpl") Then
                suffixLength = 7
            Else
                Throw New NotSupportedException(Me.GetType.Name & " から、テーブル名を導出できません.")
            End If
            Return StringUtil.DecamelizeIgnoreNumber(Left(Me.GetType.Name, Len(Me.GetType.Name) - suffixLength))
        End Function

        ''' <summary>
        ''' テーブル値を全件取得する
        ''' </summary>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>結果を要素の同名プロパティに保持したList</returns>
        ''' <remarks></remarks>
        Public Function FindAll(Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T) Implements DaoEachTable(Of T).FindAll
            Return FindBy(DirectCast(Nothing, T), selectionCallback)
        End Function

        Protected Function FindByPkMain(ByVal ParamArray values() As Object) As T
            Dim pkWhereClause As T = MakePkVo(values)
            Dim results As List(Of T) = FindByMain(pkWhereClause, GetPkPropertyInfos(), Nothing, selectionCallback:=Nothing)
            If results.Count = 0 Then
                Return Nothing
            End If
            Return results(0)
        End Function

        ''' <summary>
        ''' テーブル値の検索結果を返す
        ''' </summary>
        ''' <param name="where">検索条件</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>結果を要素の同名プロパティに保持したList</returns>
        ''' <remarks></remarks>
        Public Function FindBy(where As T,
                               Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T) Implements DaoEachTable(Of T).FindBy
            Return FindByMain(where, Nothing, Nothing, selectionCallback)
        End Function

        Private Function FindByMain(where As T, onlyFieldsWhere As List(Of PropertyInfo), criteria As Criteria(Of T), selectionCallback As Func(Of SelectionField(Of T), T, Object)) As List(Of T)
            Dim db As DbClient = NewDbClient(GetUnableCamelizeNames)
            Dim unableCamelizeNames As Dictionary(Of String, String) = db.InternalGetUnableCamelizeNames(Of T)()
            Dim selectClause As String = SqlUtil.MakeSelectClause(selectionCallback, unableCamelizeNames)
            Dim whereClause As String
            If criteria IsNot Nothing Then
                whereClause = SqlUtil.MakeWhereClauseByCriteria(criteria, unableCamelizeNames)
            ElseIf onlyFieldsWhere Is Nothing Then
                whereClause = SqlUtil.MakeWhereClause(where, unableCamelizeNames)
            Else
                whereClause = SqlUtil.MakeWhereClauseOnly(where, onlyFieldsWhere, unableCamelizeNames)
            End If
            Dim sql As String = BuildSelectSqlIfForUpdate(selectClause, "FROM " & GetTableName(), whereClause)
            Return db.QueryForList(Of T)(sql, If(DirectCast(criteria, Object), where))
        End Function

        ''' <summary>
        ''' テーブル値の検索結果を返す
        ''' </summary>
        ''' <param name="criteria">検索条件</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>結果を要素の同名プロパティに保持したList</returns>
        ''' <remarks></remarks>
        Public Function FindBy(criteria As Criteria(Of T),
                               Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T) Implements DaoEachTable(Of T).FindBy
            Return FindByMain(where:=Nothing, onlyFieldsWhere:=Nothing, criteria:=criteria, selectionCallback:=selectionCallback)
        End Function

        ''' <summary>
        ''' テーブル値の検索結果を返す
        ''' </summary>
        ''' <param name="criteriaCallback">検索条件を指定するcallback</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>結果を要素の同名プロパティに保持したList</returns>
        Public Function FindBy(criteriaCallback As Func(Of CriteriaBinder, T, CriteriaBinder),
                               Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T) Implements DaoEachTable(Of T).FindBy
            Dim vo As T = VoUtil.NewInstance(Of T)()
            Dim criteria As New Criteria(Of T)(vo)
            criteriaCallback.Invoke(criteria, vo)
            Return FindBy(criteria, selectionCallback)
        End Function

        ''' <summary>
        ''' Select文を構築する（SELECT FOR UPDATE指定を考慮）
        ''' </summary>
        ''' <param name="select">Select句</param>
        ''' <param name="from">From句</param>
        ''' <param name="where">Where句</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Function BuildSelectSqlIfForUpdate([select] As String, [from] As String, where As String) As String
            If Not ForUpdate Then
                Return (New StringBuilder([select])).Append([from]).Append(where).ToString
            End If
            Dim db As DbClient = NewDbClient(GetUnableCamelizeNames)
            If Not db.IsBeginningTransaction Then
                Throw New InvalidOperationException("ForUpdateを実行する時は、トランザクション管理を開始して下さい.")
            End If
            Return db.MakeSelectForUpdate([select], [from], where)
        End Function

        ''' <summary>
        ''' 該当する件数を返す
        ''' </summary>
        ''' <param name="where">検索条件</param>
        ''' <returns>件数</returns>
        ''' <remarks></remarks>
        Public Function CountBy(ByVal where As T) As Integer Implements DaoEachTable(Of T).CountBy
            Dim db As DbClient = NewDbClient(GetUnableCamelizeNames)
            Dim whereClause As String = SqlUtil.MakeWhereClause(where, db.InternalGetUnableCamelizeNames(Of T))
            Dim sql As String = "SELECT COUNT(*) FROM " & GetTableName() & whereClause
            Return db.QueryForObject(Of Integer)(sql, where)
        End Function

        ''' <summary>
        ''' レコードを追加する
        ''' </summary>
        ''' <param name="values">追加する値[]</param>
        ''' <returns>追加した件数</returns>
        ''' <remarks></remarks>
        Public Function InsertBy(ParamArray values As T()) As Integer Implements DaoEachTable(Of T).InsertBy
            Return PerformInsert(values, isWithoutNull:=False)
        End Function

        ''' <summary>
        ''' レコードを追加する(値がNullならデフォルト値とする)
        ''' </summary>
        ''' <param name="values">追加する値[]</param>
        ''' <returns>追加した件数</returns>
        ''' <remarks></remarks>
        Public Function InsertDefaultIfNullBy(ParamArray values As T()) As Integer Implements DaoEachTable(Of T).InsertDefaultIfNullBy
            Return PerformInsert(values, isWithoutNull:=True)
        End Function

        Private Function PerformInsert(values As T(), isWithoutNull As Boolean) As Integer
            AssertParameterIsNotNull(values, "values")
            AssertParameterIsNotNull(values(0), "values(0)")
            Dim db As DbClient = NewDbClient(GetUnableCamelizeNames)
            If 1 < values.Length Then
                db.AssertCanMultiValuesInsert()
            End If
            Dim result As Integer = 0
            For Each splittedValues As T() In CollectionUtil.Split(values, _InsertStatementLimitedRows)
                Dim InsertValues As String
                If isWithoutNull Then
                    InsertValues = SqlUtil.MakeInsertIntoWithoutNull(Of T())(splittedValues, db.InternalGetUnableCamelizeNames(Of T))
                Else
                    InsertValues = SqlUtil.MakeInsertInto(Of T())(splittedValues, db.InternalGetUnableCamelizeNames(Of T))
                End If
                Dim sql As String = "INSERT INTO " & GetTableName() & " " & InsertValues
                result += db.Insert(sql, IIf(splittedValues.Length = 1, splittedValues(0), splittedValues))
            Next
            Return result
        End Function

        ''' <summary>
        ''' 該当レコードを更新する
        ''' </summary>
        ''' <param name="pkWhereAndValue">検索条件（PK項目）と、更新値（その他項目）</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Public Function UpdateByPk(ByVal pkWhereAndValue As T) As Integer Implements DaoEachTable(Of T).UpdateByPk
            Return PerformUpdateIgnoreNullByPk(pkWhereAndValue, isIgnoreNull:=False)
        End Function

        ''' <summary>
        ''' 該当レコードを更新する(値がNullならその項目は更新しない)
        ''' </summary>
        ''' <param name="pkWhereAndValue">検索条件（PK項目）と、更新値（その他項目）</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Public Function UpdateIgnoreNullByPk(ByVal pkWhereAndValue As T) As Integer Implements DaoEachTable(Of T).UpdateIgnoreNullByPk
            Return PerformUpdateIgnoreNullByPk(pkWhereAndValue, isIgnoreNull:=True)
        End Function

        ''' <summary>
        ''' 該当レコードを更新する(値がNullならその項目は更新しない)
        ''' </summary>
        ''' <param name="pkWhereAndValue">検索条件（PK項目）と、更新値（その他項目）</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Private Function PerformUpdateIgnoreNullByPk(ByVal pkWhereAndValue As T, ByVal isIgnoreNull As Boolean) As Integer
            AssertPrimaryKeyIsNotEmpty()
            AssertParameterIsNotNull(pkWhereAndValue, "pkWhereAndValue")
            Dim db As DbClient = NewDbClient(GetUnableCamelizeNames)
            Dim setAndWhere As String
            If isIgnoreNull Then
                setAndWhere = SqlUtil.MakeUpdateSetWithoutNullAndWhere(Of T)(pkWhereAndValue, GetPkPropertyInfos, db.InternalGetUnableCamelizeNames(Of T))
            Else
                setAndWhere = SqlUtil.MakeUpdateSetWhere(Of T)(pkWhereAndValue, GetPkPropertyInfos, db.InternalGetUnableCamelizeNames(Of T))
            End If
            Dim sql As String = "UPDATE " & GetTableName() & " SET " & setAndWhere
            Return db.Update(sql, pkWhereAndValue)
        End Function

        ''' <summary>
        ''' 該当レコードを削除する
        ''' </summary>
        ''' <param name="where">検索条件</param>
        ''' <returns>削除件数</returns>
        ''' <remarks></remarks>
        Public Function DeleteBy(ByVal where As T) As Integer Implements DaoEachTable(Of T).DeleteBy
            Return DeleteByMain(where, Nothing)
        End Function

        ''' <summary>
        ''' <para>リンクサーバーの該当レコードを削除する</para>
        ''' <para>※LinkServerからのDeleteBy()を有効にしたい場合に限りこのメソッドをDaoImplから呼び出して利用すること</para>
        ''' </summary>
        ''' <param name="where">検索条件</param>
        ''' <returns>削除件数</returns>
        ''' <remarks></remarks>
        Protected Function DeleteLinkServerBy(where As T) As Integer
            Return DeleteByMain(where, Nothing, allowLinkServer:=True)
        End Function

        Protected Function DeleteByPkMain(ByVal ParamArray values() As Object) As Integer
            Dim pkWhereClause As T = MakePkVoIfNullError(values)
            Return DeleteByMain(pkWhereClause, GetPkPropertyInfos(), allowLinkServer:=True)
        End Function

        Private Function DeleteByMain(ByVal where As T, ByVal onlyFields As List(Of PropertyInfo), _
                                      Optional ByVal allowLinkServer As Boolean = False) As Integer
            AssertParameterIsNotNull(where, "where")
            Dim db As DbClient = NewDbClient(GetUnableCamelizeNames)
            db.AllowDeleteFromLinkServer = allowLinkServer
            Dim whereClause As String
            If onlyFields Is Nothing Then
                whereClause = SqlUtil.MakeWhereClause(where, db.InternalGetUnableCamelizeNames(Of T))
            Else
                whereClause = SqlUtil.MakeWhereClauseOnly(where, onlyFields, db.InternalGetUnableCamelizeNames(Of T))
            End If
            Dim sql As String = "DELETE FROM " & GetTableName() & whereClause
            Return db.Delete(sql, where)
        End Function

        ''' <summary>
        ''' PrimaryKey設定(テーブル)インターフェース
        ''' </summary>
        ''' <typeparam name="E"></typeparam>
        ''' <remarks></remarks>
        Protected Interface PkTable(Of E)
            ''' <summary>
            ''' テーブルVOを宣言する
            ''' </summary>
            ''' <param name="aTableVo">テーブルVO</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function IsA(ByVal aTableVo As E) As PkInfoField
        End Interface

        ''' <summary>
        ''' PrimaryKey設定(フィールド)インターフェース
        ''' </summary>
        ''' <remarks></remarks>
        Protected Interface PkInfoField
            ''' <summary>
            ''' PrimaryKeyのプロパティを宣言する
            ''' </summary>
            ''' <param name="aFields">PrimaryKeyに設定するプロパティ[]</param>
            ''' <returns>PrimaryKey設定インターフェース</returns>
            ''' <remarks></remarks>
            Function PkField(ByVal ParamArray aFields As Object()) As PkInfoField

            ''' <summary>
            ''' PrimaryKeyのプロパティを宣言する
            ''' </summary>
            ''' <typeparam name="U">型</typeparam>
            ''' <param name="aExpression">ラムダ式</param>
            ''' <returns>PrimaryKey設定インターフェース</returns>
            ''' <remarks></remarks>
            Function PkField(Of U)(ByVal aExpression As Expression(Of Func(Of U))) As PkInfoField
        End Interface

        Private Class PkTableImpl : Implements PkTable(Of T)
            Private ReadOnly marker As New VoPropertyMarker
            Public pkFields As New List(Of PropertyInfo)
            Public Function IsA(ByVal aTableVo As T) As PkInfoField Implements PkTable(Of T).IsA
                marker.Clear()
                marker.MarkVo(aTableVo)
                Return New PkInfoFieldImpl(Me)
            End Function
            Friend Sub Message(ByVal aField As Object)
                pkFields.Add(marker.GetPropertyInfo(aField))
            End Sub
            Friend Sub Message(Of U)(ByVal aField As Expression(Of Func(Of U)))
                pkFields.Add(marker.GetPropertyInfo(aField))
            End Sub
        End Class

        Private Class PkInfoFieldImpl : Implements PkInfoField
            Private _table As PkTableImpl
            Public Sub New(ByVal _table As PkTableImpl)
                Me._table = _table
            End Sub
            Public Function PkField(ByVal ParamArray aFields As Object()) As PkInfoField Implements PkInfoField.PkField
                For Each aField As Object In aFields
                    _table.Message(aField)
                Next
                Return Me
            End Function
            Public Function PkField(Of U)(ByVal aExpression As Expression(Of Func(Of U))) As PkInfoField Implements PkInfoField.PkField
                _table.Message(Of U)(aExpression)
                Return Me
            End Function
        End Class

        Private _pkFieldNames As List(Of PropertyInfo)
        Private Function GetPkPropertyInfos() As List(Of PropertyInfo)
            If _pkFieldNames Is Nothing Then
                Dim info As New PkTableImpl
                SettingPkField(info)
                _pkFieldNames = info.pkFields
            End If
            Return _pkFieldNames
        End Function

        Private Sub AssertPrimaryKeyIsNotEmpty()
            If CollectionUtil.IsEmpty(GetPkPropertyInfos()) Then
                Throw New InvalidOperationException("PrimaryKeyがありません。")
            End If
        End Sub

        ''' <summary>
        ''' PrimaryKeyのVOを作成して返す
        ''' </summary>
        ''' <param name="values">PrimaryKeyを構成する値</param>
        ''' <returns>VO</returns>
        ''' <remarks></remarks>
        Public Function MakePkVo(ByVal ParamArray values() As Object) As T Implements DaoEachTable(Of T).MakePkVo
            Return PerformMakePkVo(True, values)
        End Function
        Private Function MakePkVoIfNullError(ByVal ParamArray values() As Object) As T
            Return PerformMakePkVo(False, values)
        End Function
        ''' <summary>
        ''' PrimaryKeyのVOを作成して返す
        ''' </summary>
        ''' <param name="allowNull">PK値のnullを許可する場合、true</param>
        ''' <param name="values">PrimaryKeyを構成する値</param>
        ''' <returns>VO</returns>
        ''' <remarks></remarks>
        Private Function PerformMakePkVo(ByVal allowNull As Boolean, ByVal ParamArray values() As Object) As T
            If GetPkPropertyInfos.Count <> values.Length Then
                Throw New ArgumentException("PKの数が一致しません.")
            End If
            If Not allowNull AndAlso values.Any(Function(val) val Is Nothing) Then
                Throw New InvalidOperationException("PKにnullが含まれており危険です.")
            End If
            Dim aType As Type = GetType(T)
            Dim result As T = CType(Activator.CreateInstance(aType), T)
            For i As Integer = 0 To values.Length - 1
                With GetPkPropertyInfos(i)
                    If values(i) IsNot Nothing AndAlso TypeUtil.GetTypeIfNullable(.PropertyType) IsNot values(i).GetType Then
                        Throw New ArgumentException(.Name & " の型は " & .PropertyType.ToString & " です.", "values(" & i & ")")
                    End If
                    .SetValue(result, values(i), Nothing)
                End With
            Next
            Return result
        End Function

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
            Function IsA(ByVal aTableVo As T) As BindFieldAndName
        End Interface

        ''' <summary>
        ''' 項目＆項目名を設定する
        ''' </summary>
        ''' <remarks></remarks>
        Protected Interface BindFieldAndName
            ''' <summary>
            ''' 項目と項目名を宣言する
            ''' </summary>
            ''' <param name="field">項目プロパティ</param>
            ''' <param name="fieldName">項目名</param>
            ''' <returns>項目＆項目名設定インターフェース</returns>
            ''' <remarks></remarks>
            Function Bind(ByVal field As Object, ByVal fieldName As String) As BindFieldAndName

            ''' <summary>
            ''' 項目と項目名を宣言する
            ''' </summary>
            ''' <typeparam name="U">型</typeparam>
            ''' <param name="aExpression">ラムダ式</param>
            ''' <param name="fieldName">項目名</param>
            ''' <returns>項目＆項目名設定インターフェース</returns>
            ''' <remarks></remarks>
            Function Bind(Of U)(ByVal aExpression As Expression(Of Func(Of U)), ByVal fieldName As String) As BindFieldAndName
        End Interface

        Private Class BindTableImpl : Implements BindTable
            Private ReadOnly marker As New VoPropertyMarker
            Public ReadOnly propertyAndNames As New Dictionary(Of String, String)
            Public Function IsA(ByVal aTableVo As T) As BindFieldAndName Implements BindTable.IsA
                marker.Clear()
                marker.MarkVo(aTableVo)
                Return New BindFieldAndNameImpl(Me)
            End Function
            Friend Sub Message(ByVal field As Object, ByVal fieldName As String)
                propertyAndNames.Add(marker.GetPropertyInfo(field).Name, fieldName)
            End Sub
            Friend Sub Message(Of U)(ByVal field As Expression(Of Func(Of U)), ByVal fieldName As String)
                propertyAndNames.Add(marker.GetPropertyInfo(field).Name, fieldName)
            End Sub
        End Class

        Private Class BindFieldAndNameImpl : Implements BindFieldAndName
            Private binder As BindTableImpl
            Public Sub New(ByVal binder As BindTableImpl)
                Me.binder = binder
            End Sub
            Public Function Bind(ByVal field As Object, ByVal fieldName As String) As BindFieldAndName Implements BindFieldAndName.Bind
                binder.Message(field, fieldName)
                Return Me
            End Function
            Public Function Bind(Of U)(ByVal aExpression As Expression(Of Func(Of U)), ByVal fieldName As String) As BindFieldAndName Implements BindFieldAndName.Bind
                binder.Message(Of U)(aExpression, fieldName)
                Return Me
            End Function
        End Class

        Private _propertyAndNames As Dictionary(Of String, String)
        Private Function GetUnableCamelizeNames() As Dictionary(Of String, String)
            If _propertyAndNames Is Nothing Then
                Dim binder As New BindTableImpl
                SettingFieldNameCamelizeIrregular(binder)
                _propertyAndNames = binder.propertyAndNames
            End If
            Return _propertyAndNames
        End Function

        ''' <summary>
        ''' SELECT FOR UPDATE を行うか？の値を設定する
        ''' </summary>
        ''' <param name="ForUpdate">SELECT FOR UPDATE を行う場合、true</param>
        ''' <remarks></remarks>
        Public Sub SetForUpdate(ByVal ForUpdate As Boolean) Implements DaoEachTable(Of T).SetForUpdate
            Me.ForUpdate = ForUpdate
        End Sub

        ''' <summary>
        ''' プロパティ名＝DB項目名がキャメライズ通りに成立しない場合、設定する
        ''' </summary>
        ''' <param name="bind">項目と項目名を設定する</param>
        ''' <remarks></remarks>
        Protected Overridable Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)

        End Sub

        ''' <summary>
        ''' 新しい DbClient のインスタンスを生成して返す
        ''' </summary>
        ''' <returns>新しい DbClient のインスタンス</returns>
        ''' <remarks></remarks>
        Protected MustOverride Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient

        ''' <summary>
        ''' PrimaryKey を設定する
        ''' </summary>
        ''' <param name="table">テーブルに対応するVOのインスタンス</param>
        ''' <remarks></remarks>
        Protected MustOverride Sub SettingPkField(ByVal table As PkTable(Of T))

        ''' <summary>
        ''' PKのみ保持するVoへ抽出する
        ''' </summary>
        ''' <param name="vo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExtractHasOnlyPkVo(vo As T) As T
            Dim result As T = VoUtil.NewInstance(Of T)()
            GetPkPropertyInfos.Select(Function(info) info.Name).ToList.ForEach(Sub(name) VoUtil.SetValue(result, name, VoUtil.GetValue(vo, name)))
            Return result
        End Function

#If DEBUG Then
        ''' <summary> (テスト用) FetchTableColumnsFromDbの実体</summary>
        Public InternalFetchTableColumnsFromDb As Func(Of String()) = Function() NewDbClient(GetUnableCamelizeNames).FetchTableColumns(Of T)()
#Else
        ''' <summary> FetchTableColumnsFromDbの実体</summary>
        Public ReadOnly InternalFetchTableColumnsFromDb As Func(Of String()) = Function() NewDbClient(GetUnableCamelizeNames).FetchTableColumns(Of T)()
#End If
        ''' <summary>
        ''' DBからテーブルが持つ列名[]を取得する
        ''' </summary>
        ''' <returns>列名[]</returns>
        ''' <remarks></remarks>
        Public Function FetchTableColumnsFromDb() As String()
            Return InternalFetchTableColumnsFromDb.Invoke
        End Function

        ''' <summary>
        ''' Voを元に列名を特定する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DetectColumnNamesByVoProperties() As String()
            Dim dbFieldNameByPropertyName As Dictionary(Of String, String) = NewDbClient(GetUnableCamelizeNames).InternalGetUnableCamelizeNames(Of T)()
            Return VoUtil.GetPropertyNames(Of T).Select(Function(propName)
                                                            If dbFieldNameByPropertyName.ContainsKey(propName) Then
                                                                Return dbFieldNameByPropertyName.Item(propName)
                                                            End If
                                                            Return StringUtil.DecamelizeIgnoreNumber(propName)
                                                        End Function).ToArray
        End Function

        ''' <summary>
        ''' Voのプロパティとテーブルの列名が一致しているかどうか
        ''' </summary>
        ''' <returns>一致している場合True</returns>
        ''' <remarks></remarks>
        Public Function MatchesVoPropertiesBetweenDbColumns() As Boolean
            Dim propNames As List(Of String) = DetectColumnNamesByVoProperties().ToList
            Dim dbColumns As List(Of String) = FetchTableColumnsFromDb().ToList
            propNames.Sort()
            dbColumns.Sort()
            Return propNames.SequenceEqual(dbColumns)
        End Function

        ''' <summary>
        ''' 指定項目を、テーブル列名にする
        ''' </summary>
        ''' <param name="selectionCallback">項目指定のcallback</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ConvertToDbColumnNames(selectionCallback As Func(Of SelectionField(Of T), T, Object)) As String()
            Dim unableCamelizeNames As Dictionary(Of String, String) = NewDbClient(GetUnableCamelizeNames).InternalGetUnableCamelizeNames(Of T)()
            Return SqlUtil.ConvertPropertyNameToDbFieldName(selectionCallback, unableCamelizeNames)
        End Function

        Private _InsertStatementLimitedRows As Integer = 100
        ''' <summary>
        ''' Insert時の行数制限を設定する
        ''' </summary>
        ''' <param name="limitedRows"></param>
        ''' <remarks></remarks>
        Public Sub SetInsertStatementLimitedRows(limitedRows As Integer) Implements DaoEachTable(Of T).SetInsertStatementLimitedRows
            _InsertStatementLimitedRows = limitedRows
        End Sub

    End Class
End Namespace