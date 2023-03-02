Imports System.Reflection
Imports Fhi.Fw.Domain

Namespace Db
    ''' <summary>
    ''' 個別FakeDao用の擬似テーブルDaoクラス
    ''' </summary>
    ''' <typeparam name="T">擬似テーブルの型</typeparam>
    ''' <remarks></remarks>
    Public Class FakeTableDao(Of T)

        ''' <summary>条件一致デリゲート</summary>
        Public Delegate Function MatchCallback(ByVal vo As T) As Boolean
        ''' <summary>キー作成デリゲート</summary>
        Public Delegate Function MakeKeyCallback(ByVal vo As T) As Object
        ''' <summary>レコード複製デリゲート</summary>
        Public Delegate Function CloneRecordCallback(ByVal vo As T) As T

#Region "Public properties..."
        ''' <summary>プライマリキーの一意キー作成</summary>
        Public MakePrimaryUniqueKey As MakeKeyCallback
        ''' <summary>Criteriaの大/小文字違いを無視するか(初期値:False)</summary>
        Public IgnoreCaseOfCriteria As Boolean = False

        ''' <summary>レコードを複製する ※デフォルトはProperty値のみをShallowコピー</summary>
        Public CloneRecord As CloneRecordCallback

        Private ReadOnly _records As New List(Of T)

        ''' <summary>擬似テーブルのレコード[]</summary>
        Public ReadOnly Property Records() As T()
            Get
                Return _records.ToArray
            End Get
        End Property
#End Region

        ''' <summary>
        ''' 擬似テーブルを初期化する
        ''' </summary>
        ''' <param name="records">擬似テーブルのレコード一覧</param>
        ''' <remarks></remarks>
        Public Sub InitializeTable(ByVal ParamArray records As T())
            Me._records.Clear()
            AddRecords(records)
        End Sub

        ''' <summary>
        ''' 擬似テーブルにレコードを追加する
        ''' </summary>
        ''' <param name="records">追加レコード[]</param>
        ''' <remarks></remarks>
        Public Sub AddRecords(ByVal ParamArray records As T())
            Me._records.AddRange(records.Select(Function(r) Clone(r)))
        End Sub

        ''' <summary>
        ''' 擬似テーブル内から条件に一致したレコードを返す
        ''' </summary>
        ''' <param name="where">条件</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>レコード ※見つからなければNothing</returns>
        ''' <remarks></remarks>
        Public Function FindForObject(where As T,
                                      Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As T
            Return FindForObject(MakeTheMatchCallbackItsOnlyHasValue(where), selectionCallback)
        End Function

        ''' <summary>
        ''' 擬似テーブル内から条件に一致したレコードを返す
        ''' </summary>
        ''' <param name="findClause">条件 ※ラムダ式</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>レコード ※見つからなければNothing</returns>
        ''' <remarks></remarks>
        Public Function FindForObject(findClause As MatchCallback,
                                      Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As T
            For Each vo As T In _records
                If findClause(vo) Then
                    Return Clone(vo, selectionCallback)
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' 擬似テーブル内から条件に一致したレコードを返す
        ''' </summary>
        ''' <param name="criteria">条件</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>レコード ※見つからなければNothing</returns>
        ''' <remarks></remarks>
        Public Function FindForObject(criteria As Criteria(Of T),
                                      Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As T
            Return FindForList(criteria, selectionCallback).FirstOrDefault()
        End Function

        ''' <summary>
        ''' 擬似テーブル内から条件に一致したレコードを返す
        ''' </summary>
        ''' <param name="criteriaCallback">検索条件を作るラムダ式</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>レコード ※見つからなければNothing</returns>
        ''' <remarks></remarks>
        Public Function FindForObject(criteriaCallback As Func(Of CriteriaBinder, T, CriteriaBinder),
                                      Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As T
            Return FindForList(criteriaCallback, selectionCallback).FirstOrDefault()
        End Function

        ''' <summary>
        ''' 擬似テーブル内から条件に一致したレコード[]をすべて返す
        ''' </summary>
        ''' <param name="where">条件</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>レコード[]</returns>
        ''' <remarks></remarks>
        Public Function FindForList(where As T,
                                    Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T)
            Return FindForList(MakeTheMatchCallbackItsOnlyHasValue(where), selectionCallback)
        End Function

        ''' <summary>
        ''' 擬似テーブル内から条件に一致したレコード[]をすべて返す
        ''' </summary>
        ''' <param name="findClause">条件 ※ラムダ式</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>レコード[]</returns>
        ''' <remarks></remarks>
        Public Function FindForList(findClause As MatchCallback,
                                    Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T)
            Dim results As New List(Of T)
            For Each vo As T In _records
                If findClause(vo) Then
                    results.Add(Clone(vo, selectionCallback))
                End If
            Next
            Return results
        End Function

        ''' <summary>
        ''' 擬似テーブル内から条件に一致したレコード[]をすべて返す
        ''' </summary>
        ''' <param name="criteria">条件</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>レコード[]</returns>
        ''' <remarks></remarks>
        Public Function FindForList(criteria As Criteria(Of T),
                                    Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T)
            Dim findClause As MatchCallback = MakeTheMatchCallback(criteria, IgnoreCaseOfCriteria)
            Return _records.Where(Function(vo) findClause(vo)).Select(Function(vo) Clone(vo, selectionCallback)).ToList()
        End Function

        ''' <summary>
        ''' 擬似テーブル内から条件に一致したレコード[]をすべて返す
        ''' </summary>
        ''' <param name="criteriaCallback">検索条件を作るラムダ式</param>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns>レコード ※見つからなければNothing</returns>
        ''' <remarks></remarks>
        Public Function FindForList(criteriaCallback As Func(Of CriteriaBinder, T, CriteriaBinder),
                                    Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T)
            Dim vo As T = VoUtil.NewInstance(Of T)()
            Dim criteria As New Criteria(Of T)(vo)
            criteriaCallback.Invoke(criteria, vo)
            Return FindForList(criteria, selectionCallback)
        End Function

        ''' <summary>
        ''' 擬似テーブルを特定のキーごとのレコード一覧にして返す
        ''' </summary>
        ''' <param name="aMakeKey">特定のキーを作成 ※ラムダ式</param>
        ''' <returns>キーごとのレコード一覧</returns>
        ''' <remarks></remarks>
        Public Function MakeVosByKey(ByVal aMakeKey As MakeKeyCallback) As Dictionary(Of Object, List(Of T))
            Dim mapByOya As New Dictionary(Of Object, List(Of T))
            For Each vo As T In _records
                Dim key As Object = aMakeKey(vo)
                If Not mapByOya.ContainsKey(key) Then
                    mapByOya.Add(key, New List(Of T))
                End If
                mapByOya(key).Add(vo)
            Next
            Return mapByOya
        End Function

        ''' <summary>
        ''' 削除する
        ''' </summary>
        ''' <param name="whereForDeletion">削除データの抽出条件</param>
        ''' <returns>削除件数</returns>
        ''' <remarks></remarks>
        Public Function DeleteBy(whereForDeletion As T) As Integer
            Return DeleteBy(MakeTheMatchCallbackItsOnlyHasValue(whereForDeletion))
        End Function

        ''' <summary>
        ''' 削除する
        ''' </summary>
        ''' <param name="deleteClause">削除データの抽出条件</param>
        ''' <returns>削除件数</returns>
        ''' <remarks></remarks>
        Public Function DeleteBy(ByVal deleteClause As MatchCallback) As Integer
            Dim dests As New List(Of T)
            For Each record As T In _records
                If deleteClause(record) Then
                    dests.Add(record)
                End If
            Next
            For Each dest As T In dests
                _records.Remove(dest)
            Next
            Return dests.Count
        End Function

        Private Function Clone(vo As T, Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As T
            Dim result As T = If(CloneRecord Is Nothing, VoUtil.NewInstance(Of T)(vo), CloneRecord.Invoke(vo))
            If selectionCallback Is Nothing Then
                Return result
            End If
            Dim specifiedResult As T = VoUtil.NewInstance(Of T)()
            SqlUtil.GetSpecifyPropertyNames(selectionCallback).ToList.ForEach(Sub(name)
                                                                                  VoUtil.SetValue(specifiedResult, name, VoUtil.GetValue(result, name))
                                                                              End Sub)
            Return specifiedResult
        End Function

        Private Sub AssertRequiredMakePrimaryUniqueKey()

            If MakePrimaryUniqueKey Is Nothing Then
                Throw New NullReferenceException("#MakePrimaryUniqueKey が必要")
            End If
        End Sub

        ''' <summary>
        ''' 追加する
        ''' </summary>
        ''' <param name="vo">追加データ</param>
        ''' <param name="makePrimaryUniqueKey">プライマリキーで一意キー作成</param>
        ''' <returns>登録件数</returns>
        ''' <remarks></remarks>
        Public Function InsertBy(ByVal vo As T, ByVal makePrimaryUniqueKey As MakeKeyCallback) As Integer
            Me.MakePrimaryUniqueKey = makePrimaryUniqueKey
            Return InsertBy(vo)
        End Function

        ''' <summary>
        ''' 追加する
        ''' </summary>
        ''' <param name="vos">追加データ[]</param>
        ''' <returns>登録件数</returns>
        ''' <remarks></remarks>
        Public Function InsertBy(ParamArray vos As T()) As Integer
            Return vos.Sum(Function(vo) InsertBy(vo, ignoresPrimaryKey:=False))
        End Function

        ''' <summary>
        ''' 追加する
        ''' </summary>
        ''' <param name="vo">追加データ</param>
        ''' <param name="ignoresPrimaryKey">プライマリキーを無視するかどうか</param>
        ''' <returns>登録件数</returns>
        ''' <remarks></remarks>
        Public Function InsertBy(ByVal vo As T, ByVal ignoresPrimaryKey As Boolean) As Integer
            If Not ignoresPrimaryKey Then
                AssertRequiredMakePrimaryUniqueKey()
                Dim primaryKey As Object = MakePrimaryUniqueKey(vo)
                If MakeVosByKey(MakePrimaryUniqueKey).ContainsKey(primaryKey) Then
                    Throw New InvalidProgramException(String.Format("PK重複エラー. {0} は登録済み", primaryKey))
                End If
            End If
            AddRecords(vo)
            Return 1
        End Function

        ''' <summary>
        ''' プライマリキーの一致するレコードを更新する
        ''' </summary>
        ''' <param name="vo">更新値</param>
        ''' <param name="makePrimaryUniqueKey">プライマリキーで一意キー作成</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Public Function UpdateByPk(ByVal vo As T, ByVal makePrimaryUniqueKey As MakeKeyCallback) As Integer
            Me.MakePrimaryUniqueKey = makePrimaryUniqueKey
            Return UpdateByPk(vo)
        End Function

        ''' <summary>
        ''' プライマリキーの一致するレコードを更新する
        ''' </summary>
        ''' <param name="vo">更新値</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Public Function UpdateByPk(ByVal vo As T) As Integer
            AssertRequiredMakePrimaryUniqueKey()
            Return UpdateBy(vo, MakePrimaryUniqueKey)
        End Function

        ''' <summary>
        ''' キーの一致するレコードを更新する
        ''' </summary>
        ''' <param name="vo">更新値</param>
        ''' <param name="aMakeKey">一意キー作成</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Public Function UpdateBy(ByVal vo As T, ByVal aMakeKey As MakeKeyCallback) As Integer
            EzUtil.AssertParameterIsNotNull(aMakeKey, "aMakeKey")
            Dim vosByKey As Dictionary(Of Object, List(Of T)) = MakeVosByKey(aMakeKey)
            Dim updateKey As Object = aMakeKey(vo)
            If Not vosByKey.ContainsKey(updateKey) Then
                Return 0
            End If
            For Each record As T In vosByKey(updateKey)
                VoUtil.CopyProperties(vo, record)
            Next
            Return vosByKey(updateKey).Count
        End Function

        ''' <summary>
        ''' プライマリキーの一致するレコードを更新する（値がNullならその項目は更新しない）
        ''' </summary>
        ''' <param name="vo">更新値</param>
        ''' <param name="makePrimaryUniqueKey">プライマリキーで一意キー作成</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Public Function UpdateIgnoreNullByPk(ByVal vo As T, ByVal makePrimaryUniqueKey As MakeKeyCallback) As Integer
            Me.MakePrimaryUniqueKey = makePrimaryUniqueKey
            Return UpdateIgnoreNullByPk(vo)
        End Function

        ''' <summary>
        ''' プライマリキーの一致するレコードを更新する（値がNullならその項目は更新しない）
        ''' </summary>
        ''' <param name="vo">更新値</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Public Function UpdateIgnoreNullByPk(ByVal vo As T) As Integer
            AssertRequiredMakePrimaryUniqueKey()
            Return UpdateIgnoreNullBy(vo, MakePrimaryUniqueKey)
        End Function

        ''' <summary>
        ''' キーの一致するレコードを更新する（値がNullならその項目は更新しない）
        ''' </summary>
        ''' <param name="vo">更新値</param>
        ''' <param name="aMakeKey">一意キー作成</param>
        ''' <returns>更新件数</returns>
        ''' <remarks></remarks>
        Public Function UpdateIgnoreNullBy(ByVal vo As T, ByVal aMakeKey As MakeKeyCallback) As Integer
            EzUtil.AssertParameterIsNotNull(aMakeKey, "aMakeKey")
            Dim count As Integer = 0
            Dim vosByKey As Dictionary(Of Object, List(Of T)) = MakeVosByKey(aMakeKey)
            Dim updateKey As Object = aMakeKey(vo)
            If Not vosByKey.ContainsKey(updateKey) Then
                Return 0
            End If

            Dim propertyInfos As PropertyInfo() = GetType(T).GetProperties
            For Each record As T In vosByKey(updateKey)
                For Each info As PropertyInfo In propertyInfos
                    If info.GetGetMethod Is Nothing OrElse info.GetSetMethod Is Nothing Then
                        Continue For
                    End If
                    Dim value As Object = info.GetValue(vo, Nothing)
                    If value Is Nothing Then
                        Continue For
                    End If
                    info.SetValue(record, value, Nothing)
                Next
            Next
            Return vosByKey(updateKey).Count
        End Function

        ''' <summary>
        ''' 設定値のPropertyInfoを返す
        ''' </summary>
        ''' <param name="vo">値</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function ExtractPropertiesOnlyHasValue(ByVal vo As T) As IEnumerable(Of PropertyInfo)

            Dim results As New List(Of PropertyInfo)
            Dim propertyInfos As PropertyInfo() = vo.GetType.GetProperties
            For Each info As PropertyInfo In propertyInfos
                If info.GetGetMethod Is Nothing OrElse info.GetSetMethod Is Nothing Then
                    Continue For
                End If
                If info.GetValue(vo, Nothing) Is Nothing Then
                    Continue For
                End If
                results.Add(info)
            Next
            Return results
        End Function

        ''' <summary>
        ''' 設定値でのMatchCallbackを作成する
        ''' </summary>
        ''' <param name="vo">値</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MakeTheMatchCallbackItsOnlyHasValue(ByVal vo As T) As MatchCallback
            Dim aMakeKeyCallback As MakeKeyCallback = MakeTheMakeKeyCallbackItsOnlyHasValue(vo)
            Dim matchValue As Object = aMakeKeyCallback.Invoke(vo)
            If matchValue Is Nothing Then
                Return New MatchCallback(Function(vo2) True)
            End If
            Return New MatchCallback(Function(vo2) matchValue.Equals(aMakeKeyCallback(vo2)))
        End Function

        ''' <summary>
        ''' 設定値でのMakeKeyCallbackを作成する
        ''' </summary>
        ''' <param name="vo">値</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MakeTheMakeKeyCallbackItsOnlyHasValue(ByVal vo As T) As MakeKeyCallback
            Dim propertiesOnlyHasValue As IEnumerable(Of PropertyInfo) = ExtractPropertiesOnlyHasValue(vo)
            Return New MakeKeyCallback(Function(vo2) PerformMakeKeyByInfos(vo2, propertiesOnlyHasValue))
        End Function

        ''' <summary>
        ''' 検索条件でのMatchCallbackを作成する
        ''' </summary>
        ''' <param name="criteria">検索条件</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MakeTheMatchCallback(criteria As Criteria(Of T), ignoreCase As Boolean) As MatchCallback
            If criteria.GetParameters().Length = 0 Then
                Return New MatchCallback(Function(vo) True)
            End If
            Return PerformMakeTheMatchCallback(criteria, ignoreCase)
        End Function

        Private Shared Function PerformMakeTheMatchCallback(criteria As Criteria(Of T), ignoreCase As Boolean) As MatchCallback
            Dim callbackStack As New List(Of MatchCallback)
            For Each param As CriteriaParameter In ConvCriteriaParamsToReturnPolishNotation(criteria.GetParameters())
                If param.Condition < CriteriaParameter.SearchCondition.And Then
                    callbackStack.Add(MakeTheMatchCallbackByCriteriaParameter(param, ignoreCase))
                    Continue For
                ElseIf param.Condition = CriteriaParameter.SearchCondition.Not Then
                    Dim beforeCallback As MatchCallback = callbackStack(callbackStack.Count - 1)
                    callbackStack.Remove(beforeCallback)
                    callbackStack.Add(New MatchCallback(Function(vo) Not beforeCallback(vo)))
                    Continue For
                End If

                Dim callback1 As MatchCallback = callbackStack(callbackStack.Count - 2)
                Dim callback2 As MatchCallback = callbackStack(callbackStack.Count - 1)
                callbackStack.RemoveRange(callbackStack.Count - 2, 2)
                If param.Condition = CriteriaParameter.SearchCondition.And Then
                    callbackStack.Add(New MatchCallback(Function(vo) callback1(vo) AndAlso callback2(vo)))
                    Continue For
                ElseIf param.Condition = CriteriaParameter.SearchCondition.Or Then
                    callbackStack.Add(New MatchCallback(Function(vo) callback1(vo) OrElse callback2(vo)))
                    Continue For
                End If
                Throw New InvalidProgramException(String.Format("AND/OR以外の演算子が残っている。 Condition：'{0}'", param.Condition))
            Next

            If callbackStack.Count <> 1 Then
                Throw New InvalidOperationException("書式が間違っている？式が正常に処理できていない")
            End If

            Return callbackStack(0)
        End Function

        ''' <summary>
        ''' 検索条件を逆ポーランド記法に変換する
        ''' </summary>
        ''' <param name="params"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function ConvCriteriaParamsToReturnPolishNotation(params As CriteriaParameter()) As CriteriaParameter()
            Dim result As New List(Of CriteriaParameter)
            Dim stack As New List(Of CriteriaParameter)
            Dim insertsNotOperator As Boolean = False
            For Each param As CriteriaParameter In params
                Select Case param.Condition
                    Case CriteriaParameter.SearchCondition.And,
                         CriteriaParameter.SearchCondition.Bracket
                        stack.Add(param)
                    Case CriteriaParameter.SearchCondition.AndBracket
                        stack.Add(New CriteriaParameter(Nothing, Nothing, CriteriaParameter.SearchCondition.And, Nothing))
                        stack.Add(New CriteriaParameter(Nothing, Nothing, CriteriaParameter.SearchCondition.Bracket, Nothing))
                    Case CriteriaParameter.SearchCondition.NotBracket
                        stack.Add(New CriteriaParameter(Nothing, Nothing, CriteriaParameter.SearchCondition.Not, Nothing))
                        stack.Add(New CriteriaParameter(Nothing, Nothing, CriteriaParameter.SearchCondition.Bracket, Nothing))
                    Case CriteriaParameter.SearchCondition.Or,
                         CriteriaParameter.SearchCondition.OrBracket
                        If 0 < stack.Count AndAlso stack(stack.Count - 1).Condition = CriteriaParameter.SearchCondition.And Then
                            ' 演算子スタックの最新が"AND"なら、ANDをresultへ移動
                            result.Add(New CriteriaParameter(Nothing, Nothing, CriteriaParameter.SearchCondition.And, Nothing))
                            stack.RemoveAt(stack.Count - 1)
                        End If
                        stack.Add(New CriteriaParameter(Nothing, Nothing, CriteriaParameter.SearchCondition.Or, Nothing))
                        If param.Condition = CriteriaParameter.SearchCondition.OrBracket Then
                            stack.Add(New CriteriaParameter(Nothing, Nothing, CriteriaParameter.SearchCondition.Bracket, Nothing))
                        End If
                    Case CriteriaParameter.SearchCondition.Not
                        insertsNotOperator = True
                        Continue For
                    Case CriteriaParameter.SearchCondition.CloseBracket
                        Dim lastIndex As Integer = stack.Count - 1
                        While stack(lastIndex).Condition <> CriteriaParameter.SearchCondition.Bracket
                            ' "("が見つかるまで演算子スタックをどんどんresultへ移動
                            result.Add(stack(lastIndex))
                            stack.RemoveAt(lastIndex)
                            lastIndex -= 1
                        End While
                        stack.RemoveAt(lastIndex)
                    Case Else
                        result.Add(param)
                End Select

                If Not insertsNotOperator Then
                    Continue For
                End If
                insertsNotOperator = False
                result.Add(New CriteriaParameter(Nothing, Nothing, CriteriaParameter.SearchCondition.Not, Nothing))
            Next
            While 0 < stack.Count
                result.Add(stack(stack.Count - 1))
                stack.RemoveAt(stack.Count - 1)
            End While
            Return result.ToArray()
        End Function

        Private Shared Function MakeTheMatchCallbackByCriteriaParameter(param As CriteriaParameter, ignoreCase As Boolean) As MatchCallback
            Select Case param.Condition
                Case CriteriaParameter.SearchCondition.Equal
                    Return New MatchCallback(Function(vo) IsEqual(param.Value, VoUtil.GetValue(vo, param.PropertyName), ignoreCase))
                Case CriteriaParameter.SearchCondition.Any
                    Dim values As Object() = DirectCast(param.Value, Object())
                    Return New MatchCallback(Function(vo) Contains(values, VoUtil.GetValue(vo, param.PropertyName), ignoreCase))
                Case CriteriaParameter.SearchCondition.[Like]
                    Dim paramValue As String = ConvSqlWildcardToDotnetWildcard(StringUtil.ToString(param.Value))
                    If ignoreCase Then
                        If StringUtil.IsNotEmpty(paramValue) Then
                            paramValue = paramValue.ToUpper()
                        End If
                        Return New MatchCallback(Function(vo) StringUtil.ToUpper(StringUtil.ToString(VoUtil.GetValue(vo, param.PropertyName))) Like paramValue)
                    End If
                    Return New MatchCallback(Function(vo) StringUtil.ToString(VoUtil.GetValue(vo, param.PropertyName)) Like paramValue)
                Case CriteriaParameter.SearchCondition.GreaterThan
                    Return New MatchCallback(Function(vo) IsInvokedCompareSafety(vo, param.PropertyName, Function(value) CompareUtil.IsGreaterThan(param.Value, value, ignorePvo:=True)))
                Case CriteriaParameter.SearchCondition.LessThan
                    Return New MatchCallback(Function(vo) IsInvokedCompareSafety(vo, param.PropertyName, Function(value) CompareUtil.IsLessThan(param.Value, value, ignorePvo:=True)))
                Case CriteriaParameter.SearchCondition.GreaterEqual
                    Return New MatchCallback(Function(vo) IsInvokedCompareSafety(vo, param.PropertyName, Function(value) CompareUtil.IsGreaterEqual(param.Value, value, ignorePvo:=True)))
                Case CriteriaParameter.SearchCondition.LessEqual
                    Return New MatchCallback(Function(vo) IsInvokedCompareSafety(vo, param.PropertyName, Function(value) CompareUtil.IsLessEqual(param.Value, value, ignorePvo:=True)))
            End Select
            Throw New NotImplementedException(String.Format("検索条件'{0}'はFakeTableDaoに未実装", param.Condition))
        End Function

        Private Shared Function IsEqual(paramValue As Object, valueInVo As Object, ignoreCase As Boolean) As Boolean
            If ignoreCase AndAlso IsString(paramValue) AndAlso IsString(valueInVo) Then
                Return StringUtil.EqualsIgnoreCaseIfNull(paramValue, valueInVo)
            End If
            Return EzUtil.IsEqualIgnorePvoIfNull(paramValue, valueInVo)
        End Function

        Private Shared Function Contains(paramValues As Object, valueInVo As Object, ignoreCase As Boolean) As Boolean
            Dim values As Object() = DirectCast(paramValues, Object())
            If values.Length = 0 Then
                Return False
            End If
            If ignoreCase AndAlso IsString(values.First) AndAlso IsString(valueInVo) Then
                Return StringUtil.ContainsIgnoreCase(values.Select(Function(vo) vo.ToString()).ToArray(), valueInVo.ToString())
            End If
            Return CollectionUtil.ContainsIgnorePvo(values, valueInVo)
        End Function

        Private Shared Function IsString(value As Object) As Boolean
            Return TypeOf value Is String OrElse (TypeOf value Is PrimitiveValueObject _
                                                  AndAlso TypeOf DirectCast(value, PrimitiveValueObject).Value Is String)
        End Function

        Private Shared Function IsInvokedCompareSafety(vo As T, propertyName As String, compareCallback As Func(Of Object, Boolean)) As Boolean
            Dim value As Object = VoUtil.GetValue(vo, propertyName)
            If value Is Nothing Then
                Return False
            End If
            If TypeOf value Is PrimitiveValueObject AndAlso DirectCast(value, PrimitiveValueObject).Value Is Nothing Then
                Return False
            End If
            Return compareCallback.Invoke(value)
        End Function

        ''' <summary>
        ''' SQLのワイルドカード書式を.NETのワイルドカード書式に変換する
        ''' </summary>
        ''' <param name="str"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvSqlWildcardToDotnetWildcard(str As String) As String
            Const EVACUATE_SQL_WILDCARD_SINGLE As String = "/@\-/@\"
            Const EVACUATE_SQL_WILDCARD_MULTI As String = "/@\=/@\"
            If StringUtil.IsEmpty(str) Then
                Return str
            End If
            Return str.Replace("?_", EVACUATE_SQL_WILDCARD_SINGLE) _
                      .Replace("?%", EVACUATE_SQL_WILDCARD_MULTI) _
                      .Replace("?", "[?]").Replace("*", "[*]").Replace("#", "[#]") _
                      .Replace("_", "?").Replace("%", "*") _
                      .Replace(EVACUATE_SQL_WILDCARD_SINGLE, "_") _
                      .Replace(EVACUATE_SQL_WILDCARD_MULTI, "%")
        End Function

        Private Shared Function PerformMakeKeyByInfos(ByVal obj As Object, ByVal infos As IEnumerable(Of PropertyInfo)) As Object
            Dim values As New List(Of Object)
            For Each info As PropertyInfo In infos
                values.Add(info.GetValue(obj, Nothing))
            Next
            Return EzUtil.MakeKey(values.ToArray)
        End Function

    End Class

End Namespace