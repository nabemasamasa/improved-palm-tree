Namespace Util.LabelValue
    ''' <summary>
    ''' ラベルと値を持つVo
    ''' </summary>
    ''' <remarks></remarks>
    <DebuggerDisplay("Label: {Label}, Value: {Value}")>
    <Serializable()> Public Class LabelValueVo
        '' ラベル
        Private _label As String
        '' 値
        Private _value As String
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            ' nop
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="label">ラベル</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal label As String, ByVal value As String)
            Me.Label = label
            Me.Value = value
        End Sub

        ''' <summary>ラベル</summary>
        ''' <value>ラベル</value>
        ''' <returns>ラベル</returns>
        Public Property Label() As String
            Get
                Return _label
            End Get
            Set(ByVal value As String)
                _label = value
            End Set
        End Property

        ''' <summary>値</summary>
        ''' <value>値</value>
        ''' <returns>値</returns>
        Public Property Value() As String
            Get
                Return _value
            End Get
            Set(ByVal value As String)
                _value = value
            End Set
        End Property

    End Class
End Namespace