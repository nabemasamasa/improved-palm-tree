Imports System.Reflection

Namespace Db
    ''' <summary>
    ''' テーブル排他制御クラス
    ''' </summary>
    ''' <typeparam name="T">排他制御を行うテーブル用のVO</typeparam>
    ''' <remarks></remarks>
    Public Class TableExclusion(Of T)
        Private ReadOnly tableExclusionDao As DaoEachTable(Of T)
        Private ReadOnly updatedUserId As PropertyInfo
        Private ReadOnly updatedDate As PropertyInfo
        Private ReadOnly updatedTime As PropertyInfo
        Protected saveVo As T
        Protected checkVo As T
        Private pkVo As T
        Private lastWasNotUpdated As Boolean
        Private _systemDate As SystemDate
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="tableExclusionDao">排他制御を行うテーブル用のDAO</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal tableExclusionDao As DaoEachTable(Of T))
            Me.New(tableExclusionDao, New SystemDate)
        End Sub

        ''' <summary>
        ''' コンストラクタ(通常は使用しない)
        ''' </summary>
        ''' <param name="tableExclusionDao">排他制御を行うテーブル用のDAO</param>
        ''' <param name="_systemDate">標準日時クラス</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal tableExclusionDao As DaoEachTable(Of T), ByVal _systemDate As SystemDate)
            Me.tableExclusionDao = tableExclusionDao
            Me._systemDate = _systemDate
            Dim genericType As Type = GetType(T)
            updatedUserId = genericType.GetProperty("UpdatedUserId")
            updatedDate = genericType.GetProperty("UpdatedDate")
            updatedTime = genericType.GetProperty("UpdatedTime")
            If updatedUserId Is Nothing OrElse updatedDate Is Nothing OrElse updatedTime Is Nothing Then
                Throw New NotSupportedException("UpdatedUserId / UpdatedDate / UpdatedTime をもつテーブルでない場合、排他制御できません.")
            End If
        End Sub
        ''' <summary>
        ''' 排他制御に必要な情報を保存します
        ''' </summary>
        ''' <param name="pkFields">（可変引数）主キー値</param>
        ''' <remarks></remarks>
        Public Overridable Sub Save(ByVal ParamArray pkFields() As Object)
            Me.pkVo = tableExclusionDao.MakePkVo(pkFields)
            saveVo = FindExclusionRecordByPk()
            If saveVo Is Nothing Then
                Throw New InvalidOperationException("PK (" & Join(pkFields, ", ") & ") に該当するレコードは存在しません.")
            End If
            checkVo = Nothing
            lastWasNotUpdated = False
        End Sub

        ''' <summary>
        ''' 更新されているかを返す
        ''' </summary>
        ''' <returns>更新された場合、true</returns>
        ''' <remarks>#Save()のあとに、使用する</remarks>
        Public Function WasUpdatedBySomeone() As Boolean
            lastWasNotUpdated = False
            tableExclusionDao.SetForUpdate(True)
            Try
                checkVo = FindExclusionRecordByPk()
                If checkVo Is Nothing Then
                    Return True
                End If
                If Same(updatedUserId.GetValue(saveVo, Nothing), updatedUserId.GetValue(checkVo, Nothing)) _
                        And Same(updatedDate.GetValue(saveVo, Nothing), updatedDate.GetValue(checkVo, Nothing)) _
                        And Same(updatedTime.GetValue(saveVo, Nothing), updatedTime.GetValue(checkVo, Nothing)) Then
                    lastWasNotUpdated = True
                    Return False
                End If
            Finally
                tableExclusionDao.SetForUpdate(False)
            End Try
            Return True
        End Function

        ''' <summary>
        ''' 最終更新ユーザーIDを返す
        ''' </summary>
        ''' <returns>最終更新ユーザーID</returns>
        ''' <remarks></remarks>
        Public Function GetUpdatedUserId() As String
            If checkVo Is Nothing Then
                If saveVo Is Nothing Then
                    Throw New InvalidOperationException("#Save()か、#WasUpdatedBySomeone()を実行して下さい.")
                End If
                Return StringUtil.ToString(updatedUserId.GetValue(saveVo, Nothing))
            End If
            Return StringUtil.ToString(updatedUserId.GetValue(checkVo, Nothing))
        End Function

        ''' <summary>
        ''' 更新する
        ''' </summary>
        ''' <param name="loginId">ログインID</param>
        ''' <remarks></remarks>
        Public Sub Update(ByVal loginId As String)
            Update(loginId, False)
        End Sub

        ''' <summary>
        ''' 更新した後、排他制御に使用する情報を保存する
        ''' </summary>
        ''' <param name="loginId">ログインID</param>
        ''' <remarks></remarks>
        Public Sub UpdateAndSave(ByVal loginId As String)
            Update(loginId, True)
        End Sub

        ''' <summary>
        ''' 更新した後、排他制御に使用する情報を保存する
        ''' </summary>
        ''' <param name="loginId">ログインID</param>
        ''' <param name="isSave">保存する場合、true</param>
        ''' <remarks></remarks>
        Private Sub Update(ByVal loginId As String, ByVal isSave As Boolean)
            If Not lastWasNotUpdated Then
                Throw New InvalidOperationException("#WasUpdatedBySomeone()で更新されていない事を確認して下さい.")
            End If
            _systemDate.Clear()
            Dim newVo As T = FindExclusionRecordByPk()
            If newVo Is Nothing Then
                Throw New InvalidOperationException("排他制御中のデータを削除しています。続行出来ません。")
            End If
            updatedUserId.SetValue(newVo, loginId, Nothing)
            updatedDate.SetValue(newVo, _systemDate.CurrentDateDbFormat, Nothing)
            updatedTime.SetValue(newVo, _systemDate.CurrentTimeDbFormat, Nothing)
            tableExclusionDao.UpdateByPk(newVo)
            If isSave Then
                saveVo = newVo
                checkVo = Nothing
                lastWasNotUpdated = False
            End If
        End Sub

        ''' <summary>
        ''' 排他制御に必要な情報を返す
        ''' </summary>
        ''' <returns>排他制御に必要な情報</returns>
        ''' <remarks></remarks>
        Private Function FindExclusionRecordByPk() As T
            Dim results As List(Of T) = tableExclusionDao.FindBy(pkVo)
            If results.Count = 0 Then
                Return Nothing
            End If
            Return results(0)
        End Function

        ''' <summary>
        ''' Nothingを含めて同値か、を返す
        ''' </summary>
        ''' <param name="a">値A</param>
        ''' <param name="b">値B</param>
        ''' <returns>同値の場合、true</returns>
        ''' <remarks></remarks>
        Private Function Same(ByVal a As Object, ByVal b As Object) As Boolean
            If a Is Nothing And b Is Nothing Then
                Return True
            ElseIf a Is Nothing Then
                Return False
            ElseIf b Is Nothing Then
                Return False
            End If
            Return a.Equals(b)
        End Function
    End Class
End Namespace