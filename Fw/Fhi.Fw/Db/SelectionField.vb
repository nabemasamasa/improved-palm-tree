Imports System.Linq.Expressions

Namespace Db
    Public Interface SelectionField(Of T)
        ''' <summary>
        ''' Select句の項目を指定する
        ''' </summary>
        ''' <param name="fields">Select句の項目[]</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function [Is](ParamArray fields As Object()) As SelectionField(Of T)

        ''' <summary>
        ''' Select句の項目をdelegateで指定する
        ''' </summary>
        ''' <typeparam name="F">項目の型（※意識する必要なし）</typeparam>
        ''' <param name="aField">Select句の項目を返すdelegate</param>
        ''' <returns></returns>
        ''' <remarks>Boolean三項目以上持つVoだと内部使用しているVoPropertyMarkerでプロパティを特定できない</remarks>
        Function [Is](Of F)(aField As Expression(Of Func(Of F))) As SelectionField(Of T)
    End Interface
End Namespace