Imports NUnit.Framework

Namespace Db.Impl.LinkServer
    <Category("RequireDB")> Public MustInherit Class Testpf01DaoImplTest

        Private Const KEY_FLD001 As String = "_/ABC/_"
        Private Const KEY_FLD002 As String = "_/ABD/_"
        Private Const KEY_FLD003 As String = "_/ABE/_"
        Private sut As Testpf01DaoImpl

        <SetUp()> Public Overridable Sub SetUp()
            sut = New Testpf01DaoImpl
        End Sub

        <TearDown()> Public Overridable Sub TearDown()
            ' 当テストで作成したデータが何らかの理由で重複すると、`DELETE OPENQUERY`で削除できなくなる
            ' (エラー内容)
            ' > TearDown : System.Data.SqlClient.SqlException : リンク サーバー "IBMAS13_CYJTESTF" の OLE DB プロバイダ "MSDASQL" でテーブル "SELECT * FROM TESTPF01 WHERE FLD001 = '_/ABC/_'" から削除できませんでした。更新はスキーマの要件を満たしていませんでした。
            ' > リンク サーバー "IBMAS13_CYJTESTF" の OLE DB プロバイダ "MSDASQL" から、メッセージ "キー列の情報が足りないか、正しくありません。更新の影響を受ける行が多すぎます。" が返されました。
            '
            ' 4部構成だと問題なく削除できるので、重複しても問題ないように4部構成のDelete文にする
            Using db As New TestLinkServerDbClient
                db.DenyOpenQuery = True
                db.BeginToKeepOpen()
                sut.DeleteByPk(KEY_FLD001)
                sut.DeleteByPk(KEY_FLD002)
                sut.DeleteByPk(KEY_FLD003)
            End Using
        End Sub

        Public Class DefaultTest : Inherits Testpf01DaoImplTest

            <Test()> Public Sub Insertできる()
                Dim actual As Integer = sut.InsertBy(New Testpf01Vo With {.Fld001 = KEY_FLD001, .Fld002 = "A"})
                Assert.That(actual, [Is].EqualTo(1))
            End Sub

            <Test()> Public Sub Insertできる_n件()
                Dim actual As Integer = sut.InsertBy(New Testpf01Vo With {.Fld001 = KEY_FLD002, .Fld002 = "A"},
                                                     New Testpf01Vo With {.Fld001 = KEY_FLD003, .Fld002 = "B"},
                                                     New Testpf01Vo With {.Fld001 = KEY_FLD001, .Fld002 = "C"})
                Assert.That(actual, [Is].EqualTo(3))
            End Sub

            <Test()> Public Sub Insertしたデータを取得できる()
                sut.InsertBy(New Testpf01Vo With {.Fld001 = KEY_FLD001, .Fld002 = "B"})
                Dim actuals As List(Of Testpf01Vo) = sut.FindBy(New Testpf01Vo With {.Fld001 = KEY_FLD001})
                Assert.That(actuals, [Is].Not.Empty)
                Assert.That(actuals.Count, [Is].EqualTo(1))
                Assert.That(actuals(0).Fld001, [Is].EqualTo(KEY_FLD001))
                Assert.That(actuals(0).Fld002, [Is].EqualTo("B"))
            End Sub

            <Test()> Public Sub Updateできる()
                sut.InsertBy(New Testpf01Vo With {.Fld001 = KEY_FLD001, .Fld002 = "C"})
                Dim actual As Integer = sut.UpdateByPk(New Testpf01Vo With {.Fld001 = KEY_FLD001, .Fld002 = "D"})
                Assert.That(actual, [Is].EqualTo(1))
                Dim actuals As List(Of Testpf01Vo) = sut.FindBy(New Testpf01Vo With {.Fld001 = KEY_FLD001})
                Assert.That(actuals, [Is].Not.Empty)
                Assert.That(actuals.Count, [Is].EqualTo(1))
                Assert.That(actuals(0).Fld001, [Is].EqualTo(KEY_FLD001))
                Assert.That(actuals(0).Fld002, [Is].EqualTo("D"))
            End Sub

            <Test(), ExpectedException(expectedmessage:="リンクサーバー IBMAS13_CYJTESTF に対して DeleteByPk 以外は使用禁止.")>
            Public Sub AS400のテーブルで_DeleteByは使用不可()
                sut.DeleteBy(New Testpf01Vo With {.Fld001 = KEY_FLD001})
            End Sub

        End Class

    End Class
End Namespace