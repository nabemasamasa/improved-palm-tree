Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Diagnostics
Imports NUnit.Framework
Imports Fhi.Fw.TestUtil.DebugString
Imports Fhi.Fw.Domain

Namespace App.Xls
    Public MustInherit Class ExcelExTest

        ''' <summary>
        ''' ExcelのCOMオブジェクトがメモリに残っていないか精査する場合、true。※その時、Excelをすべて閉じておく必要がある
        ''' </summary>
        ''' <remarks>一テストごとに一定時間待機しているので、取扱いに注意</remarks>
        Private Const IS_PROCESS_CHECK As Boolean = False

        Private Const EXCEL_PROCESS_NAME As String = "Excel"
        ''' <summary>
        ''' タスクマネージャに、Excel.exe が残っていないか確認する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub AssertNotExistsProcess(ByVal ignoreProcessIds As List(Of Integer))
            ' Excelプロセスを終了するのに時間がかかる場合がある為、待機する
            Const SLEEP_MILLS As Integer = 100
            Const RETRY_COUNT As Integer = 30
            For i As Integer = 1 To RETRY_COUNT
                System.Threading.Thread.Sleep(SLEEP_MILLS)
                Dim notExistsProcess As Boolean = True
                For Each aProces As Process In Process.GetProcessesByName(EXCEL_PROCESS_NAME)
                    If ignoreProcessIds Is Nothing OrElse Not ignoreProcessIds.Contains(aProces.Id) Then
                        notExistsProcess = False
                        Exit For
                    End If
                Next
                If notExistsProcess Then
                    Return
                End If
            Next
            Assert.Fail("Excelプロセスが残っている " & Me.GetType.Name)
        End Sub

        Private Function GetDefaultCellSize() As Size
            Dim standardFontSize As Double = xls.StandardFontSize
            If Math.Abs(standardFontSize - 11.0R) < Double.Epsilon Then
                Return New Size(54, 14)
            ElseIf Math.Abs(standardFontSize - 10.0R) < Double.Epsilon Then
                Return New Size(48, 12)
            End If
            Throw New InvalidProgramException("未対応のフォントサイズ=" & standardFontSize)
        End Function

        Private xls As ExcelEx
        Private excelProcessIds As List(Of Integer)
        Private testingFileName As String

        Private Function ExtractProcessIds() As List(Of Integer)
            Dim excelProcessIds = New List(Of Integer)
            For Each aProcess As Process In Process.GetProcessesByName(EXCEL_PROCESS_NAME)
                excelProcessIds.Add(aProcess.Id)
            Next
            Return excelProcessIds
        End Function

        Private Function GetPathAndFileName(ByVal excelName As String) As String
            Return Path.Combine(AssemblyUtil.GetPath, excelName)
        End Function

        Private Function MakeExcel(ByVal fileName As String) As ExcelEx
            Dim aExcelEx As New ExcelEx(fileName)
            While aExcelEx.SheetCount < 3
                aExcelEx.AddSheet()
            End While
            aExcelEx.SetActiveSheet(1)
            Return aExcelEx
        End Function

        Private Shared Function AssignFileNameByExtension(extension As String) As String
            Return Path.Combine(AssemblyUtil.GetPath, Path.ChangeExtension(Path.GetRandomFileName, extension))
        End Function
        Private Shared Function AssignRandomFileName() As String
            Return AssignFileNameByExtension("xls")
        End Function
        Private Shared Function AssignXlsxFileName() As String
            Return AssignFileNameByExtension("xlsx")
        End Function

        <SetUp()> Public Overridable Sub SetUp()
            excelProcessIds = ExtractProcessIds()
            testingFileName = AssignRandomFileName()
            FileUtil.DeleteFileIfExist(testingFileName)
            xls = MakeExcel(testingFileName)
        End Sub

        <TearDown()> Public Overridable Sub TearDown()
            xls.Dispose()
            FileUtil.DeleteFileIfExist(testingFileName)
            If IS_PROCESS_CHECK Then
                AssertNotExistsProcess(excelProcessIds)
            End If
        End Sub

        <TestFixtureTearDown()> Public Sub TearDownOnce()
            If Not IS_PROCESS_CHECK Then
                AssertNotExistsProcess(excelProcessIds)
            End If
        End Sub

        <Category("RequireExcel")> Public Class [Default] : Inherits ExcelExTest

            <Test()> Public Sub 新規ファイルを開いたら_シートIndexは1で_シート名はSheet1()
                Assert.AreEqual(1, xls.GetActiveSheetIndex)
                Assert.AreEqual("Sheet1", xls.GetActiveSheetName)
            End Sub

            <Test()> Public Sub 新規ファイルのシートを削除したら_シートIndexは1だけど_シート名はSheet2()
                xls.DeleteActiveSheet()
                Assert.AreEqual(1, xls.GetActiveSheetIndex)
                Assert.AreEqual("Sheet2", xls.GetActiveSheetName, "Sheet1が削除されたのだから、ズレている")
            End Sub

            <Test()> Public Sub SetPrintRange_印刷範囲を設定する()
                xls.SetPrintRange(2, 2, 5, 5)
                xls.Save()
                Assert.IsTrue(True, "目視チェック：B2:E5が印刷範囲になる")
            End Sub

            <Test()> Public Sub SetPrintZoom＿拡大縮小印刷の倍率を設定する()
                xls.SetPrintRange(2, 2, 5, 5)
                xls.SetPrintZoom(200)
                xls.Save()
                Assert.IsTrue(True, "目視チェック：拡大縮小印刷の倍率が200%になる")
            End Sub

            <Test()> Public Sub SetPrintZoom＿共用範囲外は例外(<Values(1, 900, -100)> ByVal zoomValue As Integer)
                xls.SetPrintRange(2, 2, 5, 5)
                Try
                    xls.SetPrintZoom(zoomValue)
                    Assert.Fail()
                Catch ex As COMException
                    Assert.AreEqual("PageSetup クラスの Zoom プロパティを設定できません。", ex.Message, "許容範囲外はエラー")
                End Try
            End Sub

            <Test()> Public Sub SetPrintOrientation_印刷の向きを設定する()
                xls.SetPrintRange(2, 2, 5, 5)
                xls.SetPrintOrientation(XlPageOrientation.xlLandscape)
                xls.Save()
                Assert.IsTrue(True, "目視チェック：印刷の向きが「横」になる")
            End Sub

            <Test()> Public Sub SetPrintOrientation_用紙の大きさを設定する()
                xls.SetPrintRange(2, 2, 5, 5)
                xls.SetPrintPaper(XlPaperSize.xlPaperA3)
                xls.Save()
                Assert.IsTrue(True, "目視チェック：用紙サイズがA3になる")
            End Sub

            <Test()> Public Sub SetPrintOrientation_用紙の大きさを設定する_A5()
                xls.SetPrintRange(2, 2, 5, 5)
                xls.SetPrintPaper(XlPaperSize.xlPaperA5)
                xls.Save()
                Assert.IsTrue(True, "目視チェック：用紙サイズがA5になる")
            End Sub

        End Class

        <Category("RequireExcel")> Public Class AddShapeLine : Inherits ExcelExTest

            <Test()> Public Sub セル位置で線を引く()
                xls.SetValueRC(2, 2, "zzz")
                xls.AddShapeLineRC(2, 3, 3, 1)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：R2C2の右上からR2C1の左下へ線が引かれる")
            End Sub

            <Test()> Public Sub セル位置で線を引いてクロス()
                xls.SetValueRC(2, 2, "zzz")
                xls.AddShapeLineRC(1, 4, 4, 1)
                xls.AddShapeLineRC(2, 2, 3, 3)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：R1C3の右上からR3C1の左下へ線が引かれ、R2C2の左上から右下に線が引かれる")
            End Sub

            <Test()> Public Sub 座標位置で線を引く()
                xls.SetValueRC(2, 2, "zzz")
                xls.AddShapeLine(New Point(100, 100), New Point(10, 40))

                xls.Save()
                Assert.IsTrue(True, "目視チェック：R2C2の右上からR2C1の左下へ線が引かれる")
            End Sub

            <Test()> Public Sub 座標位置で線を引いてクロス()
                xls.SetValueRC(2, 2, "zzz")
                xls.AddShapeLine(New Point(100, 100), New Point(10, 40))
                xls.AddShapeLine(New Point(100, 10), New Point(40, 100))

                xls.Save()
                Assert.IsTrue(True, "目視チェック：R1C3の右上からR3C1の左下へ線が引かれ、R2C2の左上から右下に線が引かれる")
            End Sub

        End Class

        <Category("RequireExcel")> Public Class SetLocked_セルの書式_保護_ロック_を設定 : Inherits ExcelExTest

            <Test()> Public Sub Lockする()
                xls.SetValueRC(1, 2, "Lockセル")

                xls.SetLockedRC(1, 2, True)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(1,2)の「セルの書式 | 保護 | ロック」がon")
            End Sub

            <Test()> Public Sub Lock解除する()
                xls.SetValueRC(9, 2, "Lock解除セル")

                xls.SetLockedRC(9, 2, True)
                xls.SetLockedRC(9, 2, False)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(9,2)の「セルの書式 | 保護 | ロック」がoff")
            End Sub

            <Test()> Public Sub マージセルをLockする()
                xls.MergeCellsRC(3, 4, 10, 6, True)
                xls.SetValueRC(3, 4, "マージセルでLockセル")

                xls.SetLockedRC(3, 4, True)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(3,4)の「セルの書式 | 保護 | ロック」がon")
            End Sub

            <Test()> Public Sub マージセルをLock解除する()
                xls.MergeCellsRC(1, 3, 6, 3, True)
                xls.SetValueRC(1, 3, "マージセルでLock解除セル")

                xls.SetLockedRC(1, 3, True)
                xls.SetLockedRC(1, 3, False)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(1,3)の「セルの書式 | 保護 | ロック」がoff")
            End Sub

        End Class

        <Category("RequireExcel")> Public Class CopySheet_シートをコピー挿入する : Inherits ExcelExTest

            <Test()> Public Sub シートをコピー挿入する()
                xls.SetValueRC(2, 2, "qwe")
                xls.CopySheet(1, 2)

                Assert.AreEqual(2, xls.GetActiveSheetIndex)
                Assert.AreEqual("qwe", xls.GetValueRC(2, 2))
            End Sub

            <Test()> Public Sub 範囲外のシートIndexは例外(<Values(-1, 0, 4)> ByVal destSheetIndex As Integer)
                xls.SetValueRC(2, 2, "qwe")
                Try
                    xls.CopySheet(1, destSheetIndex)
                    Assert.Fail()

                Catch expected As ArgumentOutOfRangeException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub シート末尾にコピー挿入する()
                xls.SetValueRC(2, 2, "zzz")
                xls.CopySheetToLast(1)

                Assert.AreEqual(4, xls.GetActiveSheetIndex, "コピー後は、コピーしたシート位置index")
                Assert.AreEqual("zzz", xls.GetValueRC(2, 2))
            End Sub

            <Test()> Public Sub シートを右隣にコピー_シート名で()
                xls.SetValueRC(2, 2, "Hoge")

                xls.CopySheet("Sheet1")

                Assert.AreEqual(2, xls.GetSheetIndexByName("Sheet1 (2)"), "Sheet1の右隣だから index=2")
                Assert.AreEqual(2, xls.GetActiveSheetIndex, "コピー後は、コピーしたシート位置index")
                Assert.AreEqual("Hoge", xls.GetValueRC(2, 2), "コピー前の値がある")
            End Sub

            <Test()> Public Sub シートを右隣にコピー_シート名で2()
                xls.SetActiveSheet(3)
                xls.SetValueRC(2, 2, "Fuga")
                xls.SetActiveSheet(1)

                xls.CopySheet("Sheet3")

                Assert.AreEqual(4, xls.GetSheetIndexByName("Sheet3 (2)"), "Sheet3の右隣だから index=4")
                Assert.AreEqual(4, xls.GetActiveSheetIndex, "コピー後は、コピーしたシート位置index")
                Assert.AreEqual("Fuga", xls.GetValueRC(2, 2), "コピー前の値がある")
            End Sub

            <Test()> Public Sub シートを右隣にコピー_シート名で_名前指定()
                xls.SetValueRC(2, 2, "Piyo")

                xls.CopySheet("Sheet1", "CopiedSheet")

                Assert.AreEqual(2, xls.GetSheetIndexByName("CopiedSheet"), "Sheet1の右隣だから index=2")
                Assert.AreEqual(2, xls.GetActiveSheetIndex, "コピー後は、コピーしたシート位置index")
                Assert.AreEqual("Piyo", xls.GetValueRC(2, 2), "コピー前の値がある")
            End Sub
        End Class

        <Category("RequireExcel")> Public Class DeleteSheet_ : Inherits ExcelExTest

            <Test()> Public Sub Activeシートを削除する()
                xls.SetValueRC(2, 2, "Hoge")

                xls.DeleteActiveSheet()

                xls.Save()
                Assert.IsTrue(True, "目視チェック：Sheet1が削除されている")
            End Sub

            <Test()> Public Sub 指定したシートを削除する()

                xls.DeleteSheet(2)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：Sheet2が削除されている")
            End Sub

        End Class

        <Category("RequireExcel")> Public Class SetPageSetupHeader_ : Inherits ExcelExTest

            <Test()> Public Sub Activeシートを削除する()
                xls.SetValueRC(2, 2, "Hoge")

                xls.SetPageSetupHeader("LEFT", "CENTER", "RIGHT")

                xls.Save()
                Assert.IsTrue(True, "目視チェック：ファイル | ページ設定 の ヘッダー/フッター に LEFT,CENTER,RIGHT が設定されている")
            End Sub

        End Class

        <Category("RequireExcel")> Public Class [Default2] : Inherits ExcelExTest

            <Test()> Public Sub ClearContents_値をクリア()
                xls.SetValueRC(2, 2, "zzz")
                xls.SetValueRC(3, 3, "yyy")
                xls.SetValueRC(4, 4, "xxx")

                xls.ClearContentsRC(2, 2, 4, 4)

                Assert.AreEqual(Nothing, xls.GetValueRC(2, 2))
                Assert.AreEqual(Nothing, xls.GetValueRC(3, 3))
                Assert.AreEqual(Nothing, xls.GetValueRC(4, 4))
            End Sub

            <Test()> Public Sub GetCellTopLeft_セルの左上座標を返す()
                Dim point1 As Point = xls.GetCellTopLeftRC(2, 2)
                Dim defaultCellSize As Size = GetDefaultCellSize()
                Assert.AreEqual(defaultCellSize.Width, point1.X)
                Assert.AreEqual(defaultCellSize.Height, point1.Y)
            End Sub

            <Test()> Public Sub GetCellBottomRight_セルの右下座標を返す()
                Dim point1 As Point = xls.GetCellBottomRightRC(1, 1)
                Dim defaultCellSize As Size = GetDefaultCellSize()
                Assert.AreEqual(defaultCellSize.Width, point1.X)
                Assert.AreEqual(defaultCellSize.Height, point1.Y)
            End Sub

            <Test()> Public Sub GetCellSize_高さと幅を返す()
                Dim actual As Size = xls.GetCellSizeRC(1, 1)
                Dim defaultCellSize As Size = GetDefaultCellSize()
                Assert.AreEqual(defaultCellSize.Width, actual.Width)
                Assert.AreEqual(defaultCellSize.Height, actual.Height)
            End Sub

            <Test()> Public Sub GetSheetIndexByName_シート名のシート位置indexを返す()
                Assert.AreEqual(1, xls.GetSheetIndexByName("Sheet1"))
                Assert.AreEqual(2, xls.GetSheetIndexByName("Sheet2"))
                Assert.AreEqual(3, xls.GetSheetIndexByName("Sheet3"))
            End Sub

            <Test()> Public Sub GetSheetIndexesByNames_シート名一覧からシートインデックス一覧を返す()
                Dim actual As Integer() = xls.GetSheetIndexesByNames({"Sheet1", "Sheet1", "Sheet3"})
                '結果の重複は許さない
                Assert.AreEqual(1, actual(0))
                Assert.AreEqual(3, actual(1))
            End Sub

            <Test()> Public Sub IsVisibleSheet_終了後のRCWエラーにならない()
                xls.SetActiveSheet(2)
                Assert.AreEqual(True, xls.IsVisibleSheet(2))
                Assert.AreEqual(Nothing, xls.GetValueRC(2, 2))
            End Sub

            <Test()> Public Sub SetActiveCell_終了後のRCWエラーにならない()
                xls.SetActiveSheet(2)
                xls.SetActiveCell(2, "C3")
                Assert.AreEqual(Nothing, xls.GetValueRC(2, 2))
            End Sub

            <Test()> Public Sub GetSheetNameByIndex_シート位置Indexからシート名を返せる()
                Assert.AreEqual("Sheet1", xls.GetSheetNameBySheetIndex(1))
                Assert.AreEqual("Sheet2", xls.GetSheetNameBySheetIndex(2))
                Assert.AreEqual("Sheet3", xls.GetSheetNameBySheetIndex(3))
            End Sub

            <Test()> Public Sub GetSheetNamesBySheetIndexes_シートインデックス一覧からシート名一覧を返せる()
                Dim actual As String() = xls.GetSheetNamesBySheetIndexes({1, 1, 3})
                '結果の重複は許さない
                Assert.AreEqual("Sheet1", actual(0))
                Assert.AreEqual("Sheet3", actual(1))
            End Sub

        End Class

        <Category("RequireExcel")> Public Class SetFormula_計算式を設定する : Inherits ExcelExTest
            <Test()> Public Sub 空のセルの場合_計算結果は0になる()
                xls.SetFormulaRC(1, 1, "=A1*A2")
                Assert.That(xls.GetValueRC(1, 1), [Is].EqualTo(0))
            End Sub

            <Test()> Public Sub 空のセルでも_条件文を利用すれば_計算結果を空にできる()
                xls.SetFormulaRC(1, 1, "=IF(OR(A1="""", B1=""""), """", A1*B1)")
                Assert.That(xls.GetValueRC(1, 1), [Is].Empty)
            End Sub

            <Test()> Public Sub 正しくない計算式の場合_ISERROR関数で判定できる()
                Const WRONG_FORMULA As String = "A1*あ"
                xls.SetFormulaRC(1, 1, "=ISERROR(" & WRONG_FORMULA & ")")
                Assert.That(xls.GetValueRC(1, 1), [Is].True)
            End Sub

            <Test()> Public Sub 正しい計算式を設定すれば_結果を値として取得できる_四則演算の場合()
                xls.SetValueRC(2, 2, 1)
                xls.SetValueRC(3, 2, 2)
                xls.SetValueRC(4, 2, 3)
                xls.SetValueRC(5, 2, 5)
                xls.SetFormulaRC(7, 2, "=(B2+B3)*B4/B5")

                Assert.That(xls.GetValueRC(7, 2), [Is].EqualTo(1.8))
            End Sub

            <Test()> Public Sub 正しい計算式を設定すれば_結果を値として取得できる_その他の演算の場合()
                xls.SetValueRC(2, 2, 1)
                xls.SetValueRC(3, 2, 2)
                xls.SetValueRC(4, 2, 3)
                xls.SetValueRC(5, 2, 5)
                xls.SetFormulaRC(7, 2, "=SUM(B2:B5)")

                Assert.That(xls.GetValueRC(7, 2), [Is].EqualTo(11))
            End Sub
        End Class

        <Category("RequireExcel")> Public Class SetFormulaR1C1_計算式を設定する : Inherits ExcelExTest

            <Test()> Public Sub 計算式を設定すれば_結果を値として取得できる()
                xls.SetValueRC(2, 2, 1)
                xls.SetValueRC(3, 2, 2)
                xls.SetValueRC(4, 2, 3)
                xls.SetValueRC(5, 2, 5)
                xls.SetFormulaR1C1RC(7, 2, "=SUM(R[-5]C:R[-2]C)")

                Assert.That(xls.GetValueRC(7, 2), [Is].EqualTo(11))
            End Sub

        End Class

        <Category("RequireExcel")> Public Class SetDisplayGridlines_シートの枠線_表示非表示 : Inherits ExcelExTest

            <Test()> Public Sub 枠線を非表示にする()
                xls.SetValueRC(2, 2, "aiu")
                xls.SetDisplayGridlines(False)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：枠線が非表示になっている")
            End Sub

            <Test()> Public Sub 枠線設定を取得できる()
                xls.SetDisplayGridlines(False)
                Assert.That(xls.IsDisplayGridlines, [Is].False)

                xls.SetDisplayGridlines(True)
                Assert.That(xls.IsDisplayGridlines, [Is].True)
            End Sub

        End Class

        <Category("RequireExcel")> Public Class SetBordersTest : Inherits ExcelExTest

            <Test()> Public Sub _2_2_3_3の範囲を斜線_DiagonalUp_できる()
                xls.SetValueRC(2, 2, "DiagonalUp")
                xls.SetBordersRC(2, 2, 3, 3, New XlBorders With {.DiagonalUp = XlBorders.Continuous})

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)-(3,3)範囲の各セルに左下→右上の罫線が出力される")

                Dim actual As XlBorders = xls.GetBordersRC(2, 2, 3, 3)
                Assert.That(actual.DiagonalUp.LineStyle, [Is].EqualTo(XlLineStyle.xlContinuous))
                Assert.That(actual.DiagonalUp.Weight, [Is].EqualTo(XlBorderWeight.xlThin))
            End Sub

            <Test()> Public Sub _2_2_3_3の範囲全体に格子状の罫線を設定できる()
                xls.SetValueRC(2, 2, "LineStyle-Weight")
                xls.SetBordersRC(2, 2, 3, 3, New XlBorders With {.LineStyle = XlLineStyle.xlContinuous,
                                                                 .Weight = XlBorderWeight.xlThin})

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)-(3,3)範囲の各セルに四方を囲む罫線が出力される")

                Dim actual As XlBorders = xls.GetBordersRC(2, 2, 3, 3)
                Assert.That(actual.LineStyle, [Is].EqualTo(XlLineStyle.xlContinuous))
                Assert.That(actual.Weight, [Is].EqualTo(XlBorderWeight.xlThin))
            End Sub

            <Test()> Public Sub _2_2_3_3の範囲全体に格子状の罫線を設定できる_これでもできる()
                xls.SetValueRC(2, 2, "Continuous")
                xls.SetBordersRC(2, 2, 3, 3, XlBorders.Continuous)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)-(3,3)範囲の各セルに四方を囲む罫線が出力される")

                Dim actual As XlBorders = xls.GetBordersRC(2, 2, 3, 3)
                Assert.That(actual.LineStyle, [Is].EqualTo(XlLineStyle.xlContinuous))
                Assert.That(actual.Weight, [Is].EqualTo(XlBorderWeight.xlThin))
            End Sub

            <Test()> Public Sub _2_2_5_5の範囲の四方を囲む罫線を設定できる()
                xls.SetValueRC(2, 2, "left-top-right-bottom")
                xls.SetBordersRC(2, 2, 5, 5, New XlBorders With {.EdgeLeft = XlBorders.Continuous,
                                                                 .EdgeTop = XlBorders.Dash,
                                                                 .EdgeRight = XlBorders.Dot,
                                                                 .EdgeBottom = XlBorders.Double})
                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)-(5,5)範囲の四方を囲む罫線が出力される")

                Dim actual As XlBorders = xls.GetBordersRC(2, 2, 5, 5)
                Assert.That(actual.EdgeLeft.LineStyle, [Is].EqualTo(XlLineStyle.xlContinuous))
                Assert.That(actual.EdgeLeft.Weight, [Is].EqualTo(XlBorderWeight.xlThin))
                Assert.That(actual.EdgeTop.LineStyle, [Is].EqualTo(XlLineStyle.xlDash))
                Assert.That(actual.EdgeTop.Weight, [Is].EqualTo(XlBorderWeight.xlThin))
                Assert.That(actual.EdgeRight.LineStyle, [Is].EqualTo(XlLineStyle.xlDot))
                Assert.That(actual.EdgeRight.Weight, [Is].EqualTo(XlBorderWeight.xlThin))
                Assert.That(actual.EdgeBottom.LineStyle, [Is].EqualTo(XlLineStyle.xlDouble))
                Assert.That(actual.EdgeBottom.Weight, [Is].EqualTo(XlBorderWeight.xlThick))
            End Sub

            <Test()> Public Sub SetLine()
                xls.SetValueRC(2, 2, "zzz")
                xls.SetLineRC(2, 2, 3, 3, Fhi.Fw.App.Xls.XlBordersIndex.xlDiagonalUp)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)-(3,3)範囲の各セルに四方を囲む罫線と、各セルに左下→右上の罫線とが出力される")
            End Sub

            <Test()> Public Sub SetLine_2()
                xls.SetValueRC(2, 2, "zzz")
                xls.SetLineRC(2, 2, 3, 3)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)-(3,3)範囲の各セルに四方を囲む罫線が出力される")
            End Sub

            <Test()> Public Sub SetLine_3()
                xls.SetValueRC(2, 2, "zzz")
                xls.SetLineRC(2, 2, 5, 5, Fhi.Fw.App.Xls.XlBordersIndex.xlEdgeTop)
                xls.SetLineRC(2, 2, 5, 5, Fhi.Fw.App.Xls.XlBordersIndex.xlEdgeBottom)
                xls.SetLineRC(2, 2, 5, 5, Fhi.Fw.App.Xls.XlBordersIndex.xlEdgeLeft)
                xls.SetLineRC(2, 2, 5, 5, Fhi.Fw.App.Xls.XlBordersIndex.xlEdgeRight)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)-(5,5)範囲の四方を囲む罫線が出力される")
            End Sub

        End Class

        <Category("RequireExcel")> Public Class [Default3] : Inherits ExcelExTest

            <Test()> Public Sub Protect()
                xls.SetValueRC(2, 2, "編集可能")
                xls.SetLockedRC(2, 2, False)
                xls.Protect(1)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)以外のセルは編集不可能になる")
            End Sub

            <Test()> Public Sub Protect_allowFormattingCellsでセル書式のみ編集可能にできる()
                xls.SetValueRC(2, 2, "編集可能")
                xls.SetLockedRC(2, 2, False)
                xls.Protect(1, allowFormattingCells:=True)

                xls.Save()
                Assert.IsTrue(True, "目視チェック：(2,2)以外のセルは編集不可能だが、セル書式は変更できる")
            End Sub

            <Test()> Public Sub AddBarCode39_通常パターン()
                xls.SetValueRC(2, 2, "zzz")
                xls.AddBarCode39RC(1, 4, 4, 7, "C12345")

                xls.Save()
                Assert.IsTrue(True, "目視チェック：R1C4の右上からR3C6の左下までの大きさのバーコードが出力される")
            End Sub

            '<Test()> Public Sub SaveAsPDF_PDF出力()
            '    Const TESTING_PDF_NAME As String = TESTING_XLS_NAME & ".PDF"
            '    xls.SetValue(2, 2, "xxx")
            '    xls.SetValue(3, 3, "yyy")
            '    xls.SetValue(4, 4, "zzz")

            '    xls.SaveAsPDF(TESTING_PDF_NAME)
            '    Assert.IsTrue(True, "目視チェック：PDFが出力され表示できる")
            'End Sub

        End Class

        <Category("RequireExcel")> Public Class SetCharactersFontSize_R2C2に設定されている文字列の先頭から5文字のフォントサイズが8ポイントで出力される : Inherits ExcelExTest
            <Test()> Public Sub SetCharactersFontSize_先頭から5文字のフォントサイズが8ポイントで出力される()
                xls.SetValueRC(2, 2, "00000" & vbLf & "科目")
                xls.SetCharactersFontSizeRC(2, 2, 0, 5, 8)
                xls.Save()
                Assert.IsTrue(True, "目視チェック：R2C2に設定されている文字列の先頭から5文字のフォントサイズが8ポイントで出力される")
            End Sub
        End Class

        <Category("RequireExcel")> Public Class SetLockedRangeTest : Inherits ExcelExTest

            <Test()> Public Sub マージセルを丸々含む範囲でLockできる()
                xls.MergeCellsRC(1, 2, 3, 4, True)
                xls.SetLockedRangeRC(1, 2, 3, 4, isLocked:=True)
                xls.Save()
            End Sub

            <Test()> Public Sub マージセルを一部含む範囲でもLockできる()
                xls.MergeCellsRC(2, 1, 5, 6, True)
                xls.SetLockedRangeRC(3, 4, 2, 5, isLocked:=True)
                xls.Save()
            End Sub

        End Class

        <Category("RequireExcel")> Public Class ExcelXlsSaveTest : Inherits ExcelExTest
            Private testingXlsxFileName As String
            Public Overrides Sub SetUp()
                MyBase.SetUp()
                testingXlsxFileName = AssignXlsxFileName()
            End Sub

            Public Overrides Sub TearDown()
                MyBase.TearDown()
                FileUtil.DeleteFileIfExist(testingXlsxFileName)
            End Sub

            <Test()> Public Sub xlsで開き_xlsで保存_再度xlsで開くことができる()
                xls.Save()
                Using New ExcelEx(xls.FileName)
                End Using
                Assert.IsTrue(True)
            End Sub

            <Test()> Public Sub xlsで開き_xlsxで保存_再度xlsxで開くことができる()
                Dim oldFileName As String = xls.FileName
                Dim newFileName As String = testingXlsxFileName
                xls.SaveAs(newFileName)
                Using New ExcelEx(newFileName)
                End Using
                Assert.IsTrue(True)
                Assert.That(xls.FileName, [Is].EqualTo(oldFileName))
            End Sub
        End Class

        <Category("RequireExcel")> Public Class ExcelXlsxSaveTest : Inherits ExcelExTest

            Private testingXlsxFileName As String
            Public Overrides Sub SetUp()
                MyBase.SetUp()
                MyBase.TearDown()
                excelProcessIds = ExtractProcessIds()
                testingXlsxFileName = AssignXlsxFileName()
                xls = MakeExcel(testingXlsxFileName)
            End Sub

            Public Overrides Sub TearDown()
                MyBase.TearDown()
                FileUtil.DeleteFileIfExist(testingXlsxFileName)
            End Sub

            <Test()> Public Sub xlsxで開き_xlsで保存_再度xlsで開くことができる()
                Dim oldFileName As String = xls.FileName
                Dim newFileName As String = testingFileName
                xls.SaveAs(newFileName)
                Using New ExcelEx(newFileName)
                End Using
                Assert.IsTrue(True)
                Assert.That(xls.FileName, [Is].EqualTo(oldFileName))
            End Sub

            <Test()> Public Sub xlsxで開き_xlsxで保存_再度xlsxで開くことができる()
                xls.Save()
                Using New ExcelEx(xls.FileName)
                End Using
                Assert.IsTrue(True)
            End Sub
        End Class

        <Category("RequireExcel")> Public Class CopySheetOfAnotherBook : Inherits ExcelExTest
            Private Const TEST_EXCEL_NAME As String = "TEST_EXCEL.xls"
            Private Const TEST_VALUE As String = "TEST_VALUE"
            Private Const TEST_SHEET_NAME As String = "TEST_SHEET_NAME"

            Private testBookPathAndFileName As String

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                testBookPathAndFileName = GetPathAndFileName(TEST_EXCEL_NAME)
                FileUtil.DeleteFileIfExist(testBookPathAndFileName)
                MakeTestExcel(testBookPathAndFileName)
            End Sub

            Private Sub MakeTestExcel(ByVal fileName As String)
                Using testExcel As New ExcelEx(fileName)
                    While testExcel.SheetCount < 3
                        testExcel.AddSheet()
                    End While
                    testExcel.SetActiveSheet(1)
                    testExcel.SetValueRC(1, 1, TEST_VALUE)
                    testExcel.SetSheetName(TEST_SHEET_NAME)
                    testExcel.SaveAs(fileName)
                End Using
            End Sub

            <Test()> Public Sub 別のブックを指定して_3枚目の後にシートをコピーできる()
                xls.CopySheetOfAnotherBookSupportedExcel2013(testBookPathAndFileName, 1, 3)

                Dim insertedSheetIndex As Integer = xls.GetSheetIndexByName(TEST_SHEET_NAME)
                xls.SetActiveSheet(insertedSheetIndex)
                Assert.That(insertedSheetIndex, [Is].EqualTo(4))
                Assert.That(xls.SheetCount, [Is].EqualTo(4))
                Assert.That(xls.GetValueRC(1, 1), [Is].EqualTo(TEST_VALUE))
            End Sub

            <Test()> Public Sub 別のブックを指定して_3枚目の手前にシートをコピーできる()
                xls.CopySheetOfAnotherBookSupportedExcel2013(testBookPathAndFileName, 1, 3, isAfterSheet:=False)

                Dim insertedSheetIndex As Integer = xls.GetSheetIndexByName(TEST_SHEET_NAME)
                xls.SetActiveSheet(insertedSheetIndex)
                Assert.That(insertedSheetIndex, [Is].EqualTo(3), "3枚目の手前に挿入される")
                Assert.That(xls.SheetCount, [Is].EqualTo(4))
                Assert.That(xls.GetValueRC(1, 1), [Is].EqualTo(TEST_VALUE))
            End Sub

            <Test()> Public Sub 存在しないExcelファイル名を指定すると_例外となる()
                Try
                    xls.CopySheetOfAnotherBookSupportedExcel2013("BadName.xlsx", 1, 1)
                    Assert.Fail()
                Catch ex As ArgumentException
                    Assert.That(ex.Message, [Is].EqualTo("コピー対象として指定されたブックが存在しません"))
                End Try
            End Sub

            <Test()> Public Sub コピー元がDisposeされるとコピー先のファイルが削除できることを確認()
                Dim copyFileName As String = GetPathAndFileName("HOGE.xls")
                Dim fileName As String = GetPathAndFileName("FUGA.xls")
                FileUtil.DeleteFileIfExist(copyFileName)
                FileUtil.DeleteFileIfExist(fileName)

                Using testXls As ExcelEx = MakeExcel(fileName)
                    testXls.Save()
                End Using

                MakeTestExcel(copyFileName)

                Using testXls As ExcelEx = MakeExcel(fileName)
                    testXls.CopySheetOfAnotherBookSupportedExcel2013(copyFileName, 1, 3, isAfterSheet:=False)
                End Using

                FileUtil.DeleteFileIfExist(copyFileName)
                FileUtil.DeleteFileIfExist(fileName)
                Assert.That(File.Exists(copyFileName), [Is].False)
            End Sub

        End Class

        <Category("RequireExcel")> Public Class New_引数なしでの初期化 : Inherits ExcelExTest
            Public Overrides Sub SetUp()
                excelProcessIds = ExtractProcessIds()
                xls = MakeExcel(Nothing)
            End Sub

            <Test()> Public Sub 引数なしで_ExcelExを作成できる()
                Assert.That(xls.FileName, [Is].Null)
                Assert.That(xls.EndCol, [Is].EqualTo(1), "データがない場合、最終列は1")
                Assert.That(xls.EndRow, [Is].EqualTo(1), "データがない場合、最終行は1")
                Assert.That(xls.EnableEvents, [Is].True)
                Assert.That(xls.IsBookOpen, [Is].True)
            End Sub
        End Class

        <Category("RequireExcel")> Public Class CopySheetOfAnotherBookForAllTest : Inherits ExcelExTest

            Private Const TESTING_COPY_XLS_NAME As String = "ExcelExCopyTest.xls"
            Private testCopyBookFileName As String

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                testCopyBookFileName = GetPathAndFileName(TESTING_COPY_XLS_NAME)
                File.Delete(testCopyBookFileName)
                Using copyXls As ExcelEx = MakeExcel(testCopyBookFileName)
                    copyXls.SetSheetName(1, "A1")
                    copyXls.SetSheetName(2, "B2")
                    copyXls.SetSheetName(3, "C3")
                    copyXls.Save()
                End Using
            End Sub

            Public Overrides Sub TearDown()
                MyBase.TearDown()
                File.Delete(testCopyBookFileName)
            End Sub

            <Test()> Public Sub シート2の後ろにワークブックをコピーすると_コピーしたシートは3シート目になる()
                xls.CopySheetOfAnotherBookForAllSupportedExcel2013(testCopyBookFileName, 2, True)

                xls.SetActiveSheet(3)
                Assert.That(xls.GetActiveSheetName, [Is].EqualTo("A1"))
            End Sub

            <Test()> Public Sub シート2の前にワークブックをコピーすると_コピーしたシートは2シート目になる()
                xls.CopySheetOfAnotherBookForAllSupportedExcel2013(testCopyBookFileName, 2, False)

                xls.SetActiveSheet(2)
                Assert.That(xls.GetActiveSheetName, [Is].EqualTo("A1"))
            End Sub

            <Test()> Public Sub A1_B2_C3_シートを持つワークブックをコピーすると_順番通りにコピーされる()
                xls.CopySheetOfAnotherBookForAllSupportedExcel2013(testCopyBookFileName, 3, True)

                xls.SetActiveSheet(4)
                Assert.That(xls.GetActiveSheetName, [Is].EqualTo("A1"))
                xls.SetActiveSheet(5)
                Assert.That(xls.GetActiveSheetName, [Is].EqualTo("B2"))
                xls.SetActiveSheet(6)
                Assert.That(xls.GetActiveSheetName, [Is].EqualTo("C3"))
            End Sub

            <Test()> Public Sub 存在しないファイル名が渡された場合_例外になる()
                Try
                    xls.CopySheetOfAnotherBookForAllSupportedExcel2013("nonexistence.xlsx", xls.SheetCount, True)
                    Assert.Fail()
                Catch ex As ArgumentException
                    Assert.That(ex.Message, [Is].EqualTo("コピー対象として指定されたブックが存在しません"))
                End Try
            End Sub

            <Test(), Sequential()> Public Sub コピー先とコピー元の拡張子が違う場合_例外になる(<Values(".xls", ".xlsx")> ByVal copyExtension As String, _
                                                                                        <Values(".xlsx", ".xls")> ByVal destExtension As String)
                Dim copyFileName As String = GetPathAndFileName("copyFile" & copyExtension)
                Dim copyDestFileName As String = GetPathAndFileName("copyDestFile" & destExtension)
                Try

                    Using copyXls As ExcelEx = MakeExcel(copyFileName)
                        copyXls.Save()
                    End Using

                    Using destXls As ExcelEx = MakeExcel(copyDestFileName)
                        destXls.CopySheetOfAnotherBookForAllSupportedExcel2013(copyFileName, 1)
                    End Using
                    Assert.Fail()
                Catch ex As ArgumentException
                    Assert.That(ex.Message, [Is].EqualTo("コピー先のブックとコピー元のブックで拡張子が違います"))
                Finally
                    File.Delete(copyFileName)
                End Try
            End Sub

            <Test()> Public Sub コピー元がDisposeされるとコピー先のファイルが削除できることを確認()
                Dim copyFileName As String = GetPathAndFileName("HOGE.xls")
                Dim fileName As String = GetPathAndFileName("FUGA.xls")
                FileUtil.DeleteFileIfExist(copyFileName)
                FileUtil.DeleteFileIfExist(fileName)

                Using testXls As ExcelEx = MakeExcel(fileName)
                    testXls.Save()
                End Using

                Using copyXls As ExcelEx = MakeExcel(copyFileName)
                    copyXls.SetSheetName(1, "A1")
                    copyXls.SetSheetName(2, "B2")
                    copyXls.SetSheetName(3, "C3")
                    copyXls.Save()
                End Using

                Using testXls As ExcelEx = MakeExcel(fileName)
                    testXls.CopySheetOfAnotherBookForAllSupportedExcel2013(copyFileName, 2, False)
                End Using

                FileUtil.DeleteFileIfExist(copyFileName)
                FileUtil.DeleteFileIfExist(fileName)
                Assert.That(File.Exists(copyFileName), [Is].False)
            End Sub

        End Class

        <Category("RequireExcel")> Public Class New_指定されたファイル名によって生成するブックの形式が変わる : Inherits ExcelExTest

            <Test(), Sequential()> Public Sub 指定されたファイル名が_xlsなら_256列目_65536行目を超えて値を取得すると例外になる(<Values(1, 257)> ByVal col As Integer, <Values(65537, 1)> ByVal row As Integer)
                Try
                    Using tmpXls As ExcelEx = MakeExcel(GetPathAndFileName("test.xls"))
                        tmpXls.GetValueRC(row, col)
                    End Using
                    Assert.Fail()
                Catch ex As COMException
                    Assert.Pass()
                End Try
            End Sub

            <Test()> Public Sub 指定されたファイル名が_xlsなら_256列目_65536行目の値を取得しても例外にならない()
                Using tmpXls As ExcelEx = MakeExcel(GetPathAndFileName("test.xls"))
                    tmpXls.GetValueRC(65536, 256)
                End Using
                Assert.Pass()
            End Sub

            <Test(), Sequential()> Public Sub 指定されたファイル名が_xlsxなら_16384列目_1048576行目を超えて値を取得すると例外になる(<Values(1, 16385)> ByVal col As Integer, <Values(1048577, 1)> ByVal row As Integer)
                Try
                    Using tmpXls As ExcelEx = MakeExcel(GetPathAndFileName("test.xlsx"))
                        tmpXls.GetValueRC(row, col)
                    End Using
                    Assert.Fail()
                Catch ex As COMException
                    Assert.Pass()
                End Try
            End Sub

            <Test()> Public Sub 指定されたファイル名が_xlsxなら_16384列目_1048576行目の値を取得しても例外にならない()
                Using tmpXls As ExcelEx = MakeExcel(GetPathAndFileName("test.xlsx"))
                    tmpXls.GetValueRC(1048576, 16384)
                End Using
                Assert.Pass()
            End Sub

        End Class

        <Category("RequireExcel")> Public Class オートシェイプをコピーして指定シートにペーストする : Inherits ExcelExTest

            <Test()> Public Sub 直線が同じ配置でシートにコピーされる()
                xls.AddShapeLineRC(2, 3, 3, 1)
                xls.ShapeCopyPaste(2, "直線コネクタ 1")
                xls.Save()
                Assert.IsTrue(True, "目視チェック：R2C2の右上からR2C1の左下へ線が引かれ、Sheet2でも同じ結果になる")
            End Sub
        End Class

        <Category("RequireExcel")> Public Class ColumnWidthTest : Inherits ExcelExTest
            Private Const TESTING_FONT_NAME As String = "ＭＳ Ｐゴシック"
            Private Const TESTING_FONT_SIZE As Double = 10.0R
            Private defaultFontSize As Double
            Private defaultFontName As String

            Public Overrides Sub SetUp()
                '標準フォントの設定はExcelを再起動しないと有効にならないので設定して一度閉じる
                Using tmpXls As New ExcelEx
                    defaultFontSize = tmpXls.StandardFontSize
                    defaultFontName = tmpXls.StandardFont
                    tmpXls.StandardFont = TESTING_FONT_NAME
                    tmpXls.StandardFontSize = TESTING_FONT_SIZE
                End Using
                MyBase.SetUp()
            End Sub

            Public Overrides Sub TearDown()
                xls.StandardFont = defaultFontName
                xls.StandardFontSize = defaultFontSize
                MyBase.TearDown()
            End Sub

            <Test()> Public Sub 設定した値を取得できる(<Values(100.0R, 3.14R)> width As Double)
                xls.SetColumnWidth(2, width)
                Assert.That(xls.GetColumnWidth(2), [Is].EqualTo(width))
            End Sub

        End Class

        <Category("RequireExcel")> Public Class HiddenTest : Inherits ExcelExTest

            <Test()> Public Sub 非表示にできる()
                xls.SetColumnHidden(2, hidden:=True)
                xls.SetRowHidden(3, hidden:=True)
            End Sub

            <Test()> Public Sub 非表示にできる_範囲で()
                xls.SetColumnHidden(5, 8, hidden:=True)
                xls.SetRowHidden(13, 21, hidden:=True)
            End Sub

        End Class

        <Category("RequireExcel")> Public Class RowHeightTest : Inherits ExcelExTest

            <Test()> Public Sub 設定した値を取得できる(<Values(100.0R, 3.25R)> ByVal height As Double)
                xls.SetRowHeight(5, height)
                Assert.That(xls.GetRowHeight(5), [Is].EqualTo(height))
            End Sub

        End Class

        <Category("RequireExcel")> Public Class AlignmentTest : Inherits ExcelExTest

            <Test()> Public Sub Orientation_90度からマイナス90度まで指定できる(<Values(89, 45, -45, -89)> orientation As Integer)
                xls.SetAlignmentRC(2, 2, New ExcelEx.Alignment With {.Orientation = DirectCast([Enum].ToObject(GetType(XlOrientation), orientation), XlOrientation)})
                Dim actual As ExcelEx.Alignment = xls.GetAlignmentRC(2, 2)
                Assert.That(actual.Orientation, [Is].EqualTo(DirectCast([Enum].ToObject(GetType(XlOrientation), orientation), XlOrientation)))
            End Sub

            <Test()> Public Sub Orientation_定数で指定できる(<Values(XlOrientation.xlDownward, XlOrientation.xlVertical)> orientation As XlOrientation)
                xls.SetAlignmentRC(3, 3, New ExcelEx.Alignment With {.Orientation = orientation})
                Dim actual As ExcelEx.Alignment = xls.GetAlignmentRC(3, 3)
                Assert.That(actual.Orientation, [Is].EqualTo(orientation))
            End Sub

            <TestCase(90, XlOrientation.xlUpward)> <TestCase(-90, XlOrientation.xlDownward)> _
            Public Sub Orientation_90度_マイナス90度だと_適宜定数になる(orientation As Integer, expected As XlOrientation)
                xls.SetAlignmentRC(4, 4, New ExcelEx.Alignment With {.Orientation = DirectCast([Enum].ToObject(GetType(XlOrientation), orientation), XlOrientation)})
                Dim actual As ExcelEx.Alignment = xls.GetAlignmentRC(4, 4)
                Assert.That(actual.Orientation, [Is].EqualTo(expected))
            End Sub

        End Class

        <Category("RequireExcel")> Public Class ConvToTwoJaggedArrayTest : Inherits ExcelExTest

            <Test()> Public Sub _1次元配列を渡したら_2段階配列になる_ConvGotValueToStringsはすべて文字列型になる()
                Dim data As Object() = {"a", "b", 3}
                Dim expected As String()() = {New String() {"a", "b", "3"}}
                Dim actuals As String()() = ExcelEx.ConvGotValueToStrings(data)
                Assert.That(actuals, [Is].EqualTo(expected))
            End Sub

        End Class

        <Category("RequireExcel")> Public Class SetValuesRC_GetValuesRC_Test : Inherits ExcelExTest

            <Test()> Public Sub _2次元配列を設定して_取得できる()
                Dim expected As Object()() = {New Object() {"a", "b", 3}, New Object() {7, "y", "z"}}
                Dim data As Object(,) = {{"a", "b", 3}, {7, "y", "z"}}
                xls.SetValuesRC(2, 3, data)
                Dim actuals As Object()() = xls.GetValuesRC(2, 3, 3, 5)
                Assert.That(actuals, [Is].EqualTo(expected))
            End Sub

            <Test()> Public Sub _2次元配列を設定して_取得できる_数値文字列なら数値文字列のまま取得できる()
                Dim expected As Object()() = {New Object() {"a", "b", 3}, New Object() {"7", "y", "z"}}
                Dim data As Object(,) = {{"a", "b", 3}, {"7", "y", "z"}}
                xls.SetValuesRC(2, 3, data)
                Dim actuals As Object()() = xls.GetValuesRC(2, 3, 3, 5)
                Assert.That(actuals, [Is].EqualTo(expected))
            End Sub

            <Test()> Public Sub _2段階配列を設定して_取得できる()
                Dim expected As Object()() = {New Object() {"a", "b", 3}, New Object() {7, "y", "z"}}
                Dim data As Object()() = {New Object() {"a", "b", 3}, New Object() {7, "y", "z"}}
                xls.SetValuesRC(2, 3, data)
                Dim actuals As Object()() = xls.GetValuesRC(2, 3, 3, 5)
                Assert.That(actuals, [Is].EqualTo(expected))
            End Sub

        End Class

        <Category("RequireExcel")> Public Class SetValueRC_GetValueAsStringRC_Test : Inherits ExcelExTest

            <Test()> Public Sub 値を設定して_取得できる()
                xls.SetValueRC(2, 3, "a")
                Dim actual As String = xls.GetValueAsStringRC(2, 3)
                Assert.That(actual, [Is].EqualTo("a"))
            End Sub
        End Class

        <Category("RequireExcel")> Public Class CellRangeDefinition_SetValuesRC_GetValuesRC_Test : Inherits ExcelExTest

            Private Class HogeVo
                Private _HogeId As Integer?
                Private _HogeName As String
                Private _HogeTime As DateTime?
                Public Property HogeId As Integer?
                    Get
                        Return _HogeId
                    End Get
                    Set(ByVal value As Integer?)
                        _HogeId = value
                    End Set
                End Property
                Public Property HogeName As String
                    Get
                        Return _HogeName
                    End Get
                    Set(ByVal value As String)
                        _HogeName = value
                    End Set
                End Property
                Public Property HogeTime As DateTime?
                    Get
                        Return _HogeTime
                    End Get
                    Set(ByVal value As DateTime?)
                        _HogeTime = value
                    End Set
                End Property
                Public Property HogePVO As HogePVO
            End Class

            Private Class HogePVO : Inherits PrimitiveValueObject(Of String)
                Public Sub New(ByVal value As String)
                    MyBase.New(value)
                End Sub
            End Class

            Private Class FugaVo
                ' nop
            End Class

            Private Overloads Function ToString(ByVal vos As IEnumerable(Of HogeVo)) As String
                Dim maker As New DebugStringMaker(Of HogeVo)(Function(define, vo)
                                                                 Return define.Bind(vo.HogeId, vo.HogeName, vo.HogeTime)
                                                             End Function)
                Return maker.MakeString(vos)
            End Function

            <Test()> Public Sub 定義を作成して_コレクションを出し入れできる()
                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId, vo.HogeName, vo.HogeTime))

                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "salty", .HogeTime = Nothing},
                    New HogeVo With {.HogeId = 2, .HogeName = "shiomix", .HogeTime = DateTime.Parse("2018-09-21 00:00:00.000")}
                )

                xls.SetValuesRC(hogeVos)
                Dim actuals As HogeVo() = xls.GetValuesRC(Of HogeVo)(count:=2)

                Assert.That(ToString(actuals), [Is].EqualTo( _
                            "HogeId HogeName  HogeTime            " & vbCrLf & _
                            "     1 'salty'   null                " & vbCrLf & _
                            "     2 'shiomix' '2018/09/21 0:00:00'"
                            ))
            End Sub

            <Test()> Public Sub 実データより多い場合_空白行はnullになる()
                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId, vo.HogeName, vo.HogeTime))

                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "salty", .HogeTime = Nothing},
                    New HogeVo With {.HogeId = 2, .HogeName = "shiomix", .HogeTime = DateTime.Parse("2018-09-21 00:00:00.000")}
                )

                xls.SetValuesRC(hogeVos)
                Dim actuals As HogeVo() = xls.GetValuesRC(Of HogeVo)(count:=3)

                Assert.That(ToString(actuals), [Is].EqualTo( _
                            "HogeId HogeName  HogeTime            " & vbCrLf & _
                            "     1 'salty'   null                " & vbCrLf & _
                            "     2 'shiomix' '2018/09/21 0:00:00'" & vbCrLf & _
                            "null   null      null                "
                            ))
            End Sub

            <Test()> Public Sub 実データより少ない件数を指定して_データを取得できる()
                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId, vo.HogeName, vo.HogeTime))

                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "salty", .HogeTime = Nothing},
                    New HogeVo With {.HogeId = 2, .HogeName = "shiomix", .HogeTime = DateTime.Parse("2018-09-21 00:00:00.000")}
                )

                xls.SetValuesRC(hogeVos)
                Dim actuals As HogeVo() = xls.GetValuesRC(Of HogeVo)(count:=1)

                Assert.That(ToString(actuals), [Is].EqualTo( _
                            "HogeId HogeName HogeTime" & vbCrLf & _
                            "     1 'salty'  null    "
                            ))
            End Sub

            <Test()> Public Sub 定義を複数作成した場合_最新の定義でコレクションが操作される()
                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId, vo.HogeName, vo.HogeTime))

                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "salty", .HogeTime = Nothing},
                    New HogeVo With {.HogeId = 2, .HogeName = "shiomix", .HogeTime = DateTime.Parse("2018-09-21 00:00:00.000")}
                )
                xls.SetValuesRC(hogeVos)

                Dim actuals1 As HogeVo() = xls.GetValuesRC(Of HogeVo)(count:=2)
                Assert.That(ToString(actuals1), [Is].EqualTo( _
                            "HogeId HogeName  HogeTime            " & vbCrLf & _
                            "     1 'salty'   null                " & vbCrLf & _
                            "     2 'shiomix' '2018/09/21 0:00:00'"
                            ))

                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId))
                Dim actuals2 As HogeVo() = xls.GetValuesRC(Of HogeVo)(count:=2)
                Assert.That(ToString(actuals2), [Is].EqualTo( _
                            "HogeId HogeName HogeTime" & vbCrLf & _
                            "     1 null     null    " & vbCrLf & _
                            "     2 null     null    "
                            ))
            End Sub

            <Test()> Public Sub 定義がない場合_SetValuesRCは例外になる()
                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "salty", .HogeTime = Nothing},
                    New HogeVo With {.HogeId = 2, .HogeName = "shiomix", .HogeTime = DateTime.Parse("2018-09-21 00:00:00.000")}
                )

                Try
                    xls.SetValuesRC(hogeVos)
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("型 HogeVo のセル範囲定義がされていません. SetUpDefinition()で定義を作成してください."))
                End Try
            End Sub

            <Test()> Public Sub 定義がない場合_GetValuesRCは例外になる()
                Try
                    Dim ignore As HogeVo() = xls.GetValuesRC(Of HogeVo)(count:=2)
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("型 HogeVo のセル範囲定義がされていません. SetUpDefinition()で定義を作成してください."))
                End Try
            End Sub

            <Test()> Public Sub 異なる定義が指定された場合_SetValuesRCは例外になる()
                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId, vo.HogeName, vo.HogeTime))

                Dim fugaVos As List(Of FugaVo) = EzUtil.NewList(New FugaVo)

                Try
                    xls.SetValuesRC(fugaVos)
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("型 HogeVo のセル範囲定義と異なる型 FugaVo が指定されました. SetUpDefinition()で再定義してください."))
                End Try
            End Sub

            <Test()> Public Sub 異なる定義が指定された場合_GetValuesRCは例外になる()
                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId, vo.HogeName, vo.HogeTime))

                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "salty", .HogeTime = Nothing},
                    New HogeVo With {.HogeId = 2, .HogeName = "shiomix", .HogeTime = DateTime.Parse("2018-09-21 00:00:00.000")}
                )

                xls.SetValuesRC(hogeVos)

                Try
                    Dim ignore As FugaVo() = xls.GetValuesRC(Of FugaVo)(count:=1)
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("型 HogeVo のセル範囲定義と異なる型 FugaVo が指定されました. SetUpDefinition()で再定義してください."))
                End Try
            End Sub

            <Test()> Public Sub SetValuesRC_割り当てたくない列をskipできる()
                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId, Nothing, vo.HogeName))

                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "salty"},
                    New HogeVo With {.HogeId = 2, .HogeName = "dog"}
                )

                xls.SetValuesRC(hogeVos)

                Dim actualData As Object()() = xls.GetValuesRC(1, 1, 2, 3)
                Assert.That(actualData(0)(0), [Is].EqualTo(1))
                Assert.That(actualData(0)(1), [Is].Null)
                Assert.That(actualData(0)(2), [Is].EqualTo("salty"))
                Assert.That(actualData(1)(0), [Is].EqualTo(2))
                Assert.That(actualData(1)(1), [Is].Null)
                Assert.That(actualData(1)(2), [Is].EqualTo("dog"))
            End Sub

            <Test()> Public Sub GetValuesRC_割り当てたくない列をskipできる()
                xls.SetUpDefinition(Of HogeVo)(1, 1, Function(selector, vo) selector.Use(vo.HogeId, Nothing, vo.HogeName))

                xls.SetValuesRC(1, 1, New Object()() {New Object() {1, 2, "kahlua"}, New Object() {3, 4, "milk"}})

                Dim actuals As HogeVo() = xls.GetValuesRC(Of HogeVo)(count:=2)
                Assert.That(ToString(actuals), [Is].EqualTo( _
                            "HogeId HogeName HogeTime" & vbCrLf & _
                            "     1 'kahlua' null    " & vbCrLf & _
                            "     3 'milk'   null    "
                            ))
            End Sub

            <Test()>
            Public Sub SetValuesRC_デリゲートを元に出力できる()

                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "ガイア"},
                    New HogeVo With {.HogeId = 3, .HogeName = "マッシュ", .HogePVO = New HogePVO("ヒートサーベル")},
                    New HogeVo With {.HogeId = 2, .HogeName = "オルテガ", .HogePVO = New HogePVO("ジャイアントバズ")}
                )
                Dim func As XlVoRuleBuilder(Of HogeVo).Configure = Function(selector, vo)
                                                                       selector.Use(vo.HogeId, Nothing, vo.HogeName) _
                                                                       .UseWithFunc(vo.HogePVO,
                                                                                 toVoDecorator:=Function(obj)
                                                                                                    If obj Is Nothing Then
                                                                                                        Return Nothing
                                                                                                    End If
                                                                                                    Return New HogePVO(StringUtil.ToString(obj))
                                                                                                End Function,
                                                                                 toXlsDecorator:=Function(obj)
                                                                                                     Return StringUtil.ToString(obj)
                                                                                                 End Function)
                                                                       Return selector
                                                                   End Function
                xls.SetValuesRC(1, 1, hogeVos, func)

                Dim actualData As Object()() = xls.GetValuesRC(1, 1, 2, 4)

                Assert.That(actualData.Count, [Is].EqualTo(2))

                Assert.That(actualData(0)(0), [Is].EqualTo(1))
                Assert.That(actualData(0)(1), [Is].Null)
                Assert.That(actualData(0)(2), [Is].EqualTo("ガイア"))
                Assert.That(actualData(0)(3), [Is].Null)
                Assert.That(actualData(1)(0), [Is].EqualTo(3))
                Assert.That(actualData(1)(1), [Is].Null)
                Assert.That(actualData(1)(2), [Is].EqualTo("マッシュ"))
                Assert.That(actualData(1)(3), [Is].EqualTo("ヒートサーベル"))
            End Sub

            <Test()>
            Public Sub GetValuesRC_デリゲートを元に入力できる()

                xls.SetValuesRC(1, 1, {{"2", "hoge", "フラッシュ", "mix"},
                                     {"4", "fuga", "ミラクル", "blue"},
                                     {"5", "musa", "ストロング", "red"}})

                Dim vos As List(Of HogeVo) = xls.GetValuesRC(Of HogeVo)(startRow:=1, startColumn:=1, count:=3, configure:=Function(selector, vo)
                                                                                                                              selector.Use(vo.HogeId, Nothing, vo.HogeName) _
                                                                                                                              .UseWithFunc(vo.HogePVO,
                                                                                                                                        toVoDecorator:=Function(obj)
                                                                                                                                                           If obj Is Nothing Then
                                                                                                                                                               Return Nothing
                                                                                                                                                           End If
                                                                                                                                                           Return New HogePVO(StringUtil.ToString(obj))
                                                                                                                                                       End Function,
                                                                                                                                        toXlsDecorator:=Function(obj)
                                                                                                                                                            Return StringUtil.ToString(obj)
                                                                                                                                                        End Function)
                                                                                                                              Return selector
                                                                                                                          End Function).ToList

                Dim actual As String = (New DebugStringMaker(Of HogeVo)(Function(define, vo)
                                                                            define.BindWithTitle(vo.HogeId, "Id")
                                                                            define.BindWithTitle(vo.HogeName, "Name")
                                                                            define.BindWithTitle(vo.HogeTime, "Time")
                                                                            define.BindWithTitle(vo.HogePVO, "PVO")
                                                                            Return define
                                                                        End Function)).MakeString(vos)
                Assert.That(actual, [Is].EqualTo(
                            "Id Name         Time PVO   " & vbCrLf &
                            " 2 'フラッシュ' null 'mix' " & vbCrLf &
                            " 4 'ミラクル'   null 'blue'" & vbCrLf &
                            " 5 'ストロング' null 'red' "))
            End Sub

            <Test()>
            Public Sub SetValuesRC_Decoratorの設定がされていなければエラーになる()
                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(New HogeVo With {.HogeId = 1, .HogeName = "salty"},
                                                                New HogeVo With {.HogeId = 2, .HogeName = "dog", .HogePVO = New HogePVO("val")})
                Try
                    xls.SetValuesRC(1, 1, hogeVos, Function(selector, vo) selector.UseWithFunc(vo.HogePVO))
                    Assert.Fail()
                Catch ex As ArgumentException
                    Assert.That(ex.Message, [Is].EqualTo("toVoDecoratorもしくはtoXlsDecoratorのいずれかを設定する必要があります"))
                End Try
            End Sub

            <Test()>
            Public Sub GetValuesRC_Decoratorの設定がされていなければエラーになる()
                Try
                    Dim hogeVos As HogeVo() = xls.GetValuesRC(Of HogeVo)(1, 1, 1, Function(selector, vo) selector.UseWithFunc(vo.HogePVO))
                    Assert.Fail()
                Catch ex As ArgumentException
                    Assert.That(ex.Message, [Is].EqualTo("toVoDecoratorもしくはtoXlsDecoratorのいずれかを設定する必要があります"))
                End Try
            End Sub

            <Test()>
            Public Sub SetValuesRC_PVO抱えてても出力できる()

                Dim hogeVos As List(Of HogeVo) = EzUtil.NewList(
                    New HogeVo With {.HogeId = 1, .HogeName = "ザンバード", .HogePVO = New HogePVO("ザンボエース")},
                    New HogeVo With {.HogeId = 2, .HogeName = "ザンベース"},
                    New HogeVo With {.HogeId = 3, .HogeName = "ザンブル", .HogePVO = New HogePVO("ザンボット３")}
                )
                xls.SetValuesRC(1, 1, hogeVos, Function(selector, vo) selector.Use(vo.HogeId, vo.HogeName, vo.HogePVO))

                Dim actualData As Object()() = xls.GetValuesRC(1, 1, 3, 3)

                Assert.That(actualData.Count, [Is].EqualTo(3))

                Assert.That(actualData(0)(0), [Is].EqualTo(1))
                Assert.That(actualData(0)(1), [Is].EqualTo("ザンバード"))
                Assert.That(actualData(0)(2), [Is].EqualTo("ザンボエース"))
                Assert.That(actualData(1)(0), [Is].EqualTo(2))
                Assert.That(actualData(1)(1), [Is].EqualTo("ザンベース"))
                Assert.That(actualData(1)(2), [Is].Null)
                Assert.That(actualData(2)(0), [Is].EqualTo(3))
                Assert.That(actualData(2)(1), [Is].EqualTo("ザンブル"))
                Assert.That(actualData(2)(2), [Is].EqualTo("ザンボット３"))
            End Sub

            <Test()>
            Public Sub GetValuesRC_PVOを抱えててもVoに変換できる()

                xls.SetValuesRC(1, 1, {{"1", "一号", "one-line"},
                                     {"2", "二号", "two-line"},
                                     {"3", "V3", "red-mask"}})

                Dim vos As List(Of HogeVo) = xls.GetValuesRC(Of HogeVo)(1, 1, 3, Function(selector, vo) selector.Use(vo.HogeId, vo.HogeName, vo.HogePVO)).ToList

                Dim actual As String = (New DebugStringMaker(Of HogeVo)(Function(define, vo)
                                                                            define.BindWithTitle(vo.HogeId, "Id")
                                                                            define.BindWithTitle(vo.HogeName, "Name")
                                                                            define.BindWithTitle(vo.HogeTime, "Time")
                                                                            define.BindWithTitle(vo.HogePVO, "PVO")
                                                                            Return define
                                                                        End Function)).MakeString(vos)
                Assert.That(actual, [Is].EqualTo(
                            "Id Name   Time PVO       " & vbCrLf &
                            " 1 '一号' null 'one-line'" & vbCrLf &
                            " 2 '二号' null 'two-line'" & vbCrLf &
                            " 3 'V3'   null 'red-mask'"))
            End Sub


        End Class

        <Category("RequireExcel")> Public Class FindAllTest : Inherits ExcelExTest

            <Test()> Public Sub 見つかった座標をすべて取得できる()
                xls.SetValueRC(2, 2, "A1")
                xls.SetValueRC(2, 4, "A2")
                xls.SetValueRC(3, 2, "A3")
                Dim actuals As XlAxis() = xls.FindAll("*")
                Assert.That(actuals, [Is].EquivalentTo({New XlAxis(2, 2), New XlAxis(2, 4), New XlAxis(3, 2)}))
            End Sub

            <Test()> Public Sub 見つからなければ_Empty()
                xls.SetValueRC(2, 2, "A1")
                Dim actuals As XlAxis() = xls.FindAll("B")
                Assert.That(actuals, [Is].Empty)
            End Sub

        End Class

        <Category("RequireExcel")> Public Class CommentTest : Inherits ExcelExTest

            <Test()> Public Sub コメントを設定して取得できる()
                xls.SetCommentRC(2, 2, "abc")
                Dim actual As String = xls.GetCommentRC(2, 2)
                Assert.That(actual, [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub コメントが未設定ならNullになる()
                Dim actual As String = xls.GetCommentRC(2, 2)
                Assert.That(actual, [Is].Null)
            End Sub

        End Class

        <Category("RequireExcel")> Public Class ExtractComboBoxValuesTest : Inherits ExcelExTest

            <Test()> Public Sub コンボボックスに設定した_カンマ区切り値_から選択値を取得できる()
                xls.SetComboBoxFormula(2, 2, "a,b,c")
                Dim actuals As String() = xls.ExtractComboBoxValues(2, 2)
                Assert.That(actuals, [Is].EquivalentTo({"a", "b", "c"}))
            End Sub

            <Test()> Public Sub コンボボックスに設定した_セル範囲_から選択値を取得できる()
                xls.SetValueRC(5, 3, "A")
                xls.SetValueRC(6, 3, "Z")
                xls.SetComboBoxFormula(2, 2, "=Sheet1!$C$5:$C$6")
                Dim actuals As String() = xls.ExtractComboBoxValues(2, 2)
                Assert.That(actuals, [Is].EquivalentTo({"A", "Z"}))
            End Sub

            <Test()> Public Sub コンボボックスに設定した_セル範囲_から選択値を取得できる_2()
                xls.SetValueRC(5, 2, "O")
                xls.SetValueRC(6, 2, "P")
                xls.SetValueRC(7, 2, "Q")
                xls.SetComboBoxFormula(3, 3, "=B5:B7")
                Dim actuals As String() = xls.ExtractComboBoxValues(3, 3)
                Assert.That(actuals, [Is].EquivalentTo({"O", "P", "Q"}))
            End Sub

        End Class

        <Category("RequireExcel")> Public Class DetectAxisesOfAllValidationTest_データの入力規則を設定した座標を取得できる : Inherits ExcelExTest

            <Test()> Public Sub 取得できる()
                xls.SetValueRC(5, 3, "A")
                xls.SetValueRC(6, 3, "Z")
                xls.SetComboBoxFormula(2, 2, "=C5:C6")
                Dim actuals As XlAxis() = xls.DetectAxisesOfAllValidation
                Assert.That(actuals, [Is].EquivalentTo({New XlAxis(2, 2)}))
            End Sub

            <Test()> Public Sub 複数でも取得できる()
                xls.SetValueRC(5, 2, "O")
                xls.SetValueRC(6, 2, "P")
                xls.SetValueRC(7, 2, "Q")
                xls.SetComboBoxFormula(3, 3, "=B5:B7")
                xls.SetComboBoxFormula(4, 4, "=B5:B7")
                xls.SetComboBoxFormula(5, 5, "=B5:B7")
                Dim actuals As XlAxis() = xls.DetectAxisesOfAllValidation
                Assert.That(actuals, [Is].EquivalentTo({New XlAxis(3, 3), New XlAxis(4, 4), New XlAxis(5, 5)}))
            End Sub

            <Test()> Public Sub 見つからないときは_長さゼロの配列()
                Dim actuals As XlAxis() = xls.DetectAxisesOfAllValidation
                Assert.That(actuals, [Is].Empty)
            End Sub

            <Test()> Public Sub 見つからないときは_長さゼロの配列_2()
                xls.SetValueRC(6, 2, "P")
                Dim actuals As XlAxis() = xls.DetectAxisesOfAllValidation
                Assert.That(actuals, [Is].Empty)
            End Sub

        End Class

        <Category("RequireExcel")> Public Class シェイプの操作Test : Inherits ExcelExTest

            <Test()> Public Sub DetectShapeTextAndName_テキストボックスシェイプの内容を取得できる()
                Dim text1 As String = xls.AddTextBox(0, 0, 100, 25, "abcdef")
                Dim text2 As String = xls.AddTextBox(0, 30, 100, 25, "sss")
                Dim text3 As String = xls.AddTextBox(0, 60, 100, 25, "qwerty")
                Dim actual As Dictionary(Of String, String) = xls.DetectShapeTextAndName()
                Assert.That(actual(text1), [Is].EqualTo("abcdef"))
                Assert.That(actual(text2), [Is].EqualTo("sss"))
                Assert.That(actual(text3), [Is].EqualTo("qwerty"))
                Assert.That(actual.Values.ToArray, [Is].EquivalentTo({"abcdef", "sss", "qwerty"}))
            End Sub

            <Test()> Public Sub CloneShape_既存シェイプを複製できる()
                Dim baseShapeName As String = xls.AddTextBox(0, 0, 100, 25, "hogehoge")

                xls.CloneShape(baseShapeName, "Clone - 複製 - ")

                Dim actual As Dictionary(Of String, String) = xls.DetectShapeTextAndName()

                Assert.That(actual.Keys, [Is].EquivalentTo({baseShapeName, "Clone - 複製 - "}))

            End Sub

            <Test()> Public Sub CloneShape_存在しないシェイプ名を指定すれば_例外を返す()
                Dim baseShapeName As String = xls.AddTextBox(0, 0, 100, 25, "hogehoge")

                Try
                    xls.CloneShape("404 - not found - ", "Clone - 複製 - ")

                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("指定した名前のアイテムが見つかりませんでした。"))
                End Try
                Dim actual As Dictionary(Of String, String) = xls.DetectShapeTextAndName()

                Assert.That(actual.Keys, [Is].EquivalentTo({baseShapeName}))
            End Sub


            <Test()> Public Sub CloneShape_複製元と複製先が同じ場合_例外を返す()
                Dim baseShapeName As String = xls.AddTextBox(0, 0, 100, 25, "hogehoge")
                Try
                    xls.CloneShape(baseShapeName, baseShapeName)
                    Assert.Fail()
                Catch expected As ArgumentException

                    Assert.That(expected.Message, [Is].EqualTo("元となるShapeと新しいShapeの名前が同じ"))
                End Try
                Dim actual As Dictionary(Of String, String) = xls.DetectShapeTextAndName()

                Assert.That(actual.Keys, [Is].EquivalentTo({baseShapeName}), "Shapeを複製できないので一個だけ")
            End Sub

            <Test()> Public Sub DeleteShapes_指定したシェイプのみ削除できる()
                Dim shapeName1 As String = xls.AddTextBox(0, 0, 100, 25, "hogehoge")
                Dim shapeName2 As String = xls.AddTextBox(0, 40, 100, 25, "fugafuga")
                Dim shapeName3 As String = xls.AddTextBox(0, 80, 100, 25, "piyopiyo")

                xls.DeleteShapes({shapeName1, shapeName3, "non-exitsts-shape"})

                Dim actual As Dictionary(Of String, String) = xls.DetectShapeTextAndName()

                Assert.That(actual.Keys, [Is].EquivalentTo({shapeName2}))

            End Sub

        End Class

        <Category("RequireExcel")> Public Class 列番号から列名への変換Test : Inherits ExcelExTest
            <Test()> Public Sub TestConvertToLetter()
                Assert.AreEqual("", ExcelEx.ConvertToLetter(0), "EXCELの列番号は「1」始まりの為、列名は取得できない")
                Assert.AreEqual("A", ExcelEx.ConvertToLetter(1))
                Assert.AreEqual("Z", ExcelEx.ConvertToLetter(26))
                Assert.AreEqual("AA", ExcelEx.ConvertToLetter(26 * 1 + 1))
                Assert.AreEqual("ZZ", ExcelEx.ConvertToLetter(26 * 26 + 26))
                Assert.AreEqual("AAA", ExcelEx.ConvertToLetter(26 * 26 * 1 + 26 * 1 + 1))
                Assert.AreEqual("EG", ExcelEx.ConvertToLetter(26 * 5 + 7))
            End Sub
        End Class

        <Category("RequireExcel")> Public Class 読み取り専用で保存できるTest : Inherits ExcelExTest

            <Test()> Public Sub Save_読み取り専用で保存できる()
                xls.Save(isReadOnly:=True)
                Assert.IsTrue(True, "目視チェック：保存されたExcelファイルを開いた時に「読み取り専用で開きますか？」確認ダイアログが表示されること")
            End Sub

            <Test()> Public Sub SaveAs_読み取り専用で保存できる()
                xls.SaveAs(testingFileName, isReadOnly:=True)
                Assert.IsTrue(True, "目視チェック：保存されたExcelファイルを開いた時に「読み取り専用で開きますか？」確認ダイアログが表示されること")
            End Sub

        End Class

    End Class
End Namespace