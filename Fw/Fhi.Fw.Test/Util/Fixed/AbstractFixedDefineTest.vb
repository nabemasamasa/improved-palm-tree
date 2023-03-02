Imports Fhi.Fw.Util.Fixed
Imports NUnit.Framework

Namespace Util.Fixed
    Public Class AbstractFixedDefineTest

        Private Class TestingDefine : Inherits AbstractFixedDefine

            Private group As FixedGroup

            Public Sub New()
                Me.New(Nothing)
            End Sub

            Public Sub New(ByVal group As FixedGroup)
                Me.group = group
            End Sub

            Protected Overrides Function GetRootEntryImpl() As FixedGroup
                If group Is Nothing Then
                    Return New FixedGroup( _
                                New FixedField("NUM", 4, False), _
                                New FixedGroup("CUSTOMER", 3, _
                                               New FixedField("TANTO", 12, True), _
                                               New FixedNumber("WEIGHT", GetType(Decimal), 10, 2) _
                                               ), _
                                New FixedGroup("INSU", 5, _
                                                New FixedField("INSU", 2, False) _
                                               ) _
                            )
                End If
                Return group
            End Function
        End Class

        <Test()> Public Sub IsZeroPaddingIfNull_数値nullの属性はゼロ埋め()
            Dim dbpf As New TestingDefine
            dbpf.IsZeroPaddingIfNull = True
            dbpf.InitializeFixedString()

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "0000000000" _
                            & "　　　　　　　　　　　　" & "0000000000" _
                            & "　　　　　　　　　　　　" & "0000000000" _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub

        <Test()> Public Sub SetValue_全角属性に半角を入れても全角になる()
            Dim dbpf As New TestingDefine
            dbpf.SetValue("CUSTOMER[1].TANTO", "aa")

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "ａａ　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub

        <Test()> Public Sub SetValue_属性パス名は大文字小文字同一視()
            Dim dbpf As New TestingDefine
            dbpf.SetValue("Customer[1].tanTo", "aa")

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "ａａ　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub


        <Test()> Public Sub SetValue_桁数超を設定しても指定桁に切られる()
            Dim dbpf As New TestingDefine
            dbpf.SetValue("INSU[2].INSU", "abcde")

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "ab" _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub

        <Test()> Public Sub SetValue_数値属性に小数値を()
            Dim dbpf As New TestingDefine
            dbpf.SetValue("CUSTOMER[0].WEIGHT", 123.45@)

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "0000012345" _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub

        <Test()> Public Sub SetValue_数値属性に整数値()
            Dim dbpf As New TestingDefine
            dbpf.SetValue("CUSTOMER[2].WEIGHT", 12345&)

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "0001234500" _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub

        <Test()> Public Sub SetValue_数値属性に数値文字も可能()
            Dim dbpf As New TestingDefine
            dbpf.SetValue("CUSTOMER[2].WEIGHT", "123.456")

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "0000012345" _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub

        <Test()> Public Sub SetValue_数値属性に数値以外の文字は例外()
            Dim dbpf As New TestingDefine
            Try
                dbpf.SetValue("CUSTOMER[2].WEIGHT", "asdfg")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub SetValue_数値属性に大きい小数値を入れたら桁数で切られる()
            Dim dbpf As New TestingDefine
            dbpf.SetValue("CUSTOMER[0].WEIGHT", 1234567890.123456789@)

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "3456789012" _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub

        <Test()> Public Sub SetValue_数値属性にゼロ値を入れたらゼロ埋め()
            Dim dbpf As New TestingDefine
            dbpf.SetValue("CUSTOMER[1].WEIGHT", 0@)

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "0000000000" _
                            & "　　　　　　　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , dbpf.FixedString)
        End Sub

        <Test()> Public Sub InitializeFixedString_直後の値_文字列型は半角全角ともに空文字()
            Dim dbpf As New TestingDefine
            dbpf.InitializeFixedString()
            Assert.AreEqual("", dbpf.GetValue("NUM"))
            Assert.AreEqual("", dbpf.GetValue("CUSTOMER[0].TANTO"))
        End Sub

        <Test()> Public Sub InitializeFixedString_直後の値_数値型はNull()
            Dim dbpf As New TestingDefine
            dbpf.InitializeFixedString()
            Assert.IsNull(dbpf.GetValue("CUSTOMER[0].WEIGHT"))
        End Sub

        <Test()> Public Sub GetValue_数値属性の取得()
            Dim dbpf As New TestingDefine
            dbpf.FixedString = "    " _
                            & "　　　　　　　　　　　　" & "3456789012" _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  "
            Assert.AreEqual(34567890.12@, dbpf.GetValue("CUSTOMER[0].WEIGHT"))
            Assert.IsNull(dbpf.GetValue("CUSTOMER[1].WEIGHT"))
        End Sub

        <Test()> Public Sub GetValue_半角属性の取得_末尾空白はtrim()
            Dim dbpf As New TestingDefine
            dbpf.FixedString = "    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "ａｂｃ　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "x " _
                            & " y" _
                            & "  " _
                            & "  " _
                            & "  "
            Assert.AreEqual("x", dbpf.GetValue("INSU[0].INSU"))
            Assert.AreEqual(" y", dbpf.GetValue("INSU[1].INSU"))
            Assert.AreEqual("", dbpf.GetValue("INSU[2].INSU"))
        End Sub

        <Test()> Public Sub GetValue_全角属性の取得_末尾空白はtrim()
            Dim dbpf As New TestingDefine
            dbpf.FixedString = "    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "ａｂｃ　　　　　　　　　" & "          " _
                            & "　　　ｄｅｆ　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  "
            Assert.AreEqual("", dbpf.GetValue("CUSTOMER[0].TANTO"))
            Assert.AreEqual("ａｂｃ", dbpf.GetValue("CUSTOMER[1].TANTO"))
            Assert.AreEqual("　　　ｄｅｆ", dbpf.GetValue("CUSTOMER[2].TANTO"))
        End Sub

    End Class
End Namespace