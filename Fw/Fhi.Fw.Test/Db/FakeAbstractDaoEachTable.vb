Namespace Db
    ''' <summary>
    ''' テスト用のDAO擬装クラス
    ''' </summary>
    ''' <typeparam name="T">テーブルVOの型</typeparam>
    ''' <remarks></remarks>
    Public MustInherit Class FakeAbstractDaoEachTable(Of T) : Implements DaoEachTable(Of T)

        Public ResultCountBy As Integer?
        Public ParamCountBy As New List(Of T)

        Public Function FindBy(criteriaCallback As Func(Of CriteriaBinder, T, CriteriaBinder),
                               Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T) Implements DaoEachTable(Of T).FindBy
            Throw New NotImplementedException
        End Function

        Public Overridable Function CountBy(ByVal where As T) As Integer Implements DaoEachTable(Of T).CountBy
            If ResultCountBy Is Nothing Then
                Throw New NotImplementedException
            End If
            ParamCountBy.Add(where)
            Return Convert.ToInt32(ResultCountBy)
        End Function

        Public ResultDeleteBy As Integer?
        Public ParamDeleteBy As New List(Of T)
        Public Overridable Function DeleteBy(ByVal where As T) As Integer Implements DaoEachTable(Of T).DeleteBy
            If ResultDeleteBy Is Nothing Then
                Throw New NotImplementedException
            End If
            ParamDeleteBy.Add(where)
            Return Convert.ToInt32(ResultDeleteBy)
        End Function

        Public ResultFindByAll As List(Of T)
        Public Overridable Function FindAll(Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T) Implements DaoEachTable(Of T).FindAll
            If ResultFindByAll Is Nothing Then
                Throw New NotImplementedException
            End If
            Return ResultFindByAll
        End Function

        Public ResultInsertBy As Integer?
        Public ParamInsertBy As New List(Of T)
        Public Overridable Function InsertBy(ParamArray values As T()) As Integer Implements DaoEachTable(Of T).InsertBy
            If ResultInsertBy Is Nothing Then
                Throw New NotImplementedException
            End If
            ParamInsertBy.AddRange(values)
            Return Convert.ToInt32(ResultInsertBy)
        End Function

        Public ResultInsertDefaultIfNullBy As Integer?
        Public ParamInsertDefaultIfNullBy As New List(Of T)
        Public Overridable Function InsertDefaultIfNullBy(ParamArray values As T()) As Integer Implements DaoEachTable(Of T).InsertDefaultIfNullBy
            If ResultInsertDefaultIfNullBy Is Nothing Then
                Throw New NotImplementedException
            End If
            ParamInsertDefaultIfNullBy.AddRange(values)
            Return Convert.ToInt32(ResultInsertDefaultIfNullBy)
        End Function

        Public ResultUpdateIgnoreNullByPk As Integer?
        Public ParamUpdateIgnoreNullByPk As New List(Of T)
        Public Function UpdateIgnoreNullByPk(ByVal pkWhereAndValue As T) As Integer Implements DaoEachTable(Of T).UpdateIgnoreNullByPk
            If ResultUpdateIgnoreNullByPk Is Nothing Then
                Throw New NotImplementedException
            End If
            ParamUpdateIgnoreNullByPk.Add(pkWhereAndValue)
            Return ResultUpdateIgnoreNullByPk.Value
        End Function

        Public Sub SetInsertStatementLimitedRows(limitedRows As Integer) Implements DaoEachTable(Of T).SetInsertStatementLimitedRows
            ' nop
        End Sub

        Public ResultUpdateByPk As Integer?
        Public ParamUpdateByPk As New List(Of T)
        Public Overridable Function UpdateByPk(ByVal pkWhereAndValue As T) As Integer Implements DaoEachTable(Of T).UpdateByPk
            If ResultUpdateByPk Is Nothing Then
                Throw New NotImplementedException
            End If
            ParamUpdateByPk.Add(pkWhereAndValue)
            Return Convert.ToInt32(ResultUpdateByPk)
        End Function

        Public ResultFindBy As List(Of T)
        Public ParamFindBy As New List(Of T)
        Public Overridable Function FindBy(where As T,
                                           Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T) Implements DaoEachTable(Of T).FindBy
            If ResultFindBy Is Nothing Then
                Throw New NotImplementedException
            End If
            ParamFindBy.Add(where)
            Return ResultFindBy
        End Function

        Public Function FindBy(criteria As Criteria(Of T), Optional selectionCallback As Func(Of SelectionField(Of T), T, Object) = Nothing) As List(Of T) Implements DaoEachTable(Of T).FindBy
            Throw New NotImplementedException
        End Function

        Public ResultMakePkVo As T
        Public ParamMakePkVo As New List(Of Object())
        Public Overridable Function MakePkVo(ByVal ParamArray values() As Object) As T Implements DaoEachTable(Of T).MakePkVo
            If ResultMakePkVo Is Nothing Then
                Throw New NotImplementedException
            End If
            ParamMakePkVo.Add(values)
            Return ResultMakePkVo
        End Function

        Public Overridable Sub SetForUpdate(ByVal ForUpdate As Boolean) Implements DaoEachTable(Of T).SetForUpdate
            Throw New NotImplementedException
        End Sub

    End Class
End Namespace
