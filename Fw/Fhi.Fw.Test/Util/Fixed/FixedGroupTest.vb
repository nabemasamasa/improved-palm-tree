Imports Fhi.Fw.Util.Fixed
Imports NUnit.Framework

Namespace Util.Fixed
    Public Class FixedGroupTest

        <Test()> Public Sub Format_nothingでグループ内を初期化()
            Dim root As New FixedGroup( _
                    New FixedField("NUM", 4, False), _
                    New FixedGroup("CUSTOMER", 3, _
                                   New FixedField("TANTO", 12, True), _
                                   New FixedNumber("WEIGHT", GetType(Decimal), 10, 2) _
                                   ), _
                    New FixedGroup("INSU", 10, _
                                    New FixedField("INSU", 2, False) _
                                   ) _
                )

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "　　　　　　　　　　　　" & "          " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , root.Format(Nothing))
        End Sub

        <Test()> Public Sub Format_nothingでグループ内を初期化_NumberのIsZeroPaddingIfNullをtrueにすれば数値は0padding()
            Dim root As New FixedGroup( _
                    New FixedField("NUM", 4, False), _
                    New FixedGroup("CUSTOMER", 3, _
                                   New FixedField("TANTO", 12, True), _
                                   New FixedNumber("WEIGHT", GetType(Decimal), 10, 2) _
                                   ), _
                    New FixedGroup("INSU", 10, _
                                    New FixedField("INSU", 2, False) _
                                   ) _
                )
            For Each number As FixedNumber In root.DetectEntries(Of FixedNumber)()
                number.IsZeroPaddingIfNull = True
            Next

            Assert.AreEqual("    " _
                            & "　　　　　　　　　　　　" & "0000000000" _
                            & "　　　　　　　　　　　　" & "0000000000" _
                            & "　　　　　　　　　　　　" & "0000000000" _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            & "  " _
                            , root.Format(Nothing))
        End Sub

        <Test()> Public Sub Name無しのGroupは_他のGroupに内包出来ない_例外になる()
            Try
                Dim root As New FixedGroup( _
                        New FixedGroup( _
                                       New FixedField("TANTO", 12, True), _
                                       New FixedNumber("WEIGHT", GetType(Decimal), 10, 2) _
                                       ), _
                        New FixedField("hoge", 3, False) _
                    )
                Assert.Fail()
            Catch expected As InvalidOperationException
                Assert.IsTrue(True)
            End Try

        End Sub

    End Class
End Namespace