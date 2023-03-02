Namespace Db
    ''' <summary>
    ''' 1テーブルのためのCRUDを提供するDAO
    ''' </summary>
    ''' <typeparam name="T">テーブルに対応するVO</typeparam>
    ''' <remarks>実装クラスは、[テーブル名]＆"Dao" という命名規則に従う事</remarks>
    Public Interface DaoEachTable(Of T)
        ''' <summary>
        ''' テーブル値を全件取得する
        ''' </summary>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>結果を要素の同名プロパティに保持したList</returns>
        ''' <remarks></remarks>
        Function FindAll(Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T)

        ''' <summary>
        ''' テーブル値の検索結果を返す
        ''' </summary>
        ''' <param name="where">検索条件</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>結果を要素の同名プロパティに保持したList</returns>
        ''' <remarks></remarks>
        Function FindBy(where As T, Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T)

        ''' <summary>
        ''' テーブル値の検索結果を返す
        ''' </summary>
        ''' <param name="criteria">検索条件</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>結果を要素の同名プロパティに保持したList</returns>
        ''' <remarks></remarks>
        Function FindBy(criteria As Criteria(Of T),
                        Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T)

        ''' <summary>
        ''' テーブル値の検索結果を返す
        ''' </summary>
        ''' <param name="criteriaCallback">検索条件を作るラムダ式</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>結果を要素の同名プロパティに保持したList</returns>
        ''' <remarks></remarks>
        Function FindBy(criteriaCallback As Func(Of CriteriaBinder, T, CriteriaBinder),
                        Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T)

        ''' <summary>
        ''' 該当する件数を返す
        ''' </summary>
        ''' <param name="where">検索条件</param>
        ''' <returns>件数</returns>
        ''' <remarks></remarks>
        Function CountBy(ByVal where As T) As Integer

        ''' <summary>
        ''' レコードを追加する
        ''' </summary>
        ''' <param name="values">追加する値[]</param>
        ''' <returns>追加した件数</returns>
        ''' <remarks></remarks>
        Function InsertBy(ParamArray values As T()) As Integer

        ''' <summary>
        ''' 該当レコードを更新する
        ''' </summary>
        ''' <param name="pkWhereAndValue">検索条件（PK項目）と、更新値（その他項目）</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Function UpdateByPk(ByVal pkWhereAndValue As T) As Integer

        ''' <summary>
        ''' レコードを削除する
        ''' </summary>
        ''' <param name="where">削除条件</param>
        ''' <returns>削除した件数</returns>
        ''' <remarks></remarks>
        Function DeleteBy(ByVal where As T) As Integer

        ''' <summary>
        ''' 更新ロックを設定する
        ''' </summary>
        ''' <param name="forUpdate">更新ロックをする場合、true</param>
        ''' <remarks></remarks>
        Sub SetForUpdate(ByVal ForUpdate As Boolean)
        ''' <summary>
        ''' PrimaryKeyのVOを作成して返す
        ''' </summary>
        ''' <param name="values">PrimaryKeyを構成する値</param>
        ''' <returns>VO</returns>
        ''' <remarks></remarks>
        Function MakePkVo(ByVal ParamArray values() As Object) As T

        ''' <summary>
        ''' レコードを追加する(値がNullならデフォルト値とする)
        ''' </summary>
        ''' <param name="values">追加する値[]</param>
        ''' <returns>追加した件数</returns>
        ''' <remarks></remarks>
        Function InsertDefaultIfNullBy(ParamArray values As T()) As Integer

        ''' <summary>
        ''' 該当レコードを更新する(値がNullならその項目は更新しない)
        ''' </summary>
        ''' <param name="pkWhereAndValue">検索条件（PK項目）と、更新値（その他項目）</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Function UpdateIgnoreNullByPk(ByVal pkWhereAndValue As T) As Integer

        ''' <summary>
        ''' Insert時の行数制限を設定する
        ''' </summary>
        ''' <param name="limitedRows"></param>
        ''' <remarks></remarks>
        Sub SetInsertStatementLimitedRows(limitedRows As Integer)

    End Interface
End Namespace