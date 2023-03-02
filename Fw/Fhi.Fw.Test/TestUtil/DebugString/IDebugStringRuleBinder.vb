Imports System.Linq.Expressions
Imports Fhi.Fw.Domain

Namespace TestUtil.DebugString
    ''' <summary>
    ''' 検証用文字列の項目定義を担うインターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IDebugStringRuleBinder
        ''' <summary>
        ''' 出力する列を設定する
        ''' </summary>
        ''' <param name="fields">列[]</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Bind(ByVal ParamArray fields As Object()) As IDebugStringRuleBinder

        ''' <summary>
        ''' 出力する列とそのタイトルを設定する
        ''' </summary>
        ''' <param name="field">列</param>
        ''' <param name="title">出力タイトル</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function BindWithTitle(ByVal field As Object, ByVal title As String) As IDebugStringRuleBinder

        ''' <summary>
        ''' 出力する値Lambdaとそのタイトルを設定する
        ''' </summary>
        ''' <param name="fieldLambda">値Lambda</param>
        ''' <param name="title">出力タイトル</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function BindFuncWithTitle(Of T)(ByVal fieldLambda As Expression(Of Func(Of T)), ByVal title As String) As IDebugStringRuleBinder

        ''' <summary>
        ''' 出力する明細列を設定する
        ''' </summary>
        ''' <param name="field">列</param>
        ''' <param name="detailConfigure">明細列設定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function JoinDetails(Of T)(ByVal field As IEnumerable(Of T), detailConfigure As DebugStringRuleBuilder(Of T).Configure) As IDebugStringRuleBinder
        ''' <summary>
        ''' 出力する明細列を設定する
        ''' </summary>
        ''' <param name="field">列</param>
        ''' <param name="detailConfigure">明細列設定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function JoinDetails(Of T)(ByVal field As CollectionObject(Of T), detailConfigure As DebugStringRuleBuilder(Of T).Configure) As IDebugStringRuleBinder

        ''' <summary>
        ''' 横並び出力する明細列を設定する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="field">列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function JoinFromSideBySide(Of T)(ByVal field As IEnumerable(Of T)) As IDebugStringRuleBinder
        ''' <summary>
        ''' 横並び出力する明細列を設定する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="field">列</param>
        ''' <param name="sideBySideConfigure">横並び明細列設定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function JoinFromSideBySide(Of T)(ByVal field As IEnumerable(Of T), ByVal sideBySideConfigure As DebugStringRuleBuilder(Of T).Configure) As IDebugStringRuleBinder
        ''' <summary>
        ''' 横並び出力する明細列を設定する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="field">列</param>
        ''' <param name="fixedRepeatCount">固定表示繰り返し数</param>
        ''' <param name="sideBySideConfigure">横並び明細列設定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function JoinFromSideBySide(Of T)(ByVal field As IEnumerable(Of T), ByVal fixedRepeatCount As Integer?, ByVal sideBySideConfigure As DebugStringRuleBuilder(Of T).Configure) As IDebugStringRuleBinder
        ''' <summary>
        ''' 横並び出力する明細列を設定する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="field">列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function JoinFromSideBySide(Of T)(ByVal field As CollectionObject(Of T)) As IDebugStringRuleBinder
        ''' <summary>
        ''' 横並び出力する明細列を設定する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="field">列</param>
        ''' <param name="sideBySideConfigure">横並び明細列設定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function JoinFromSideBySide(Of T)(ByVal field As CollectionObject(Of T), ByVal sideBySideConfigure As DebugStringRuleBuilder(Of T).Configure) As IDebugStringRuleBinder
        ''' <summary>
        ''' 横並び出力する明細列を設定する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="field">列</param>
        ''' <param name="fixedRepeatCount">固定表示繰り返し数</param>
        ''' <param name="sideBySideConfigure">横並び明細列設定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function JoinFromSideBySide(Of T)(ByVal field As CollectionObject(Of T), ByVal fixedRepeatCount As Integer?, ByVal sideBySideConfigure As DebugStringRuleBuilder(Of T).Configure) As IDebugStringRuleBinder
    End Interface
End Namespace