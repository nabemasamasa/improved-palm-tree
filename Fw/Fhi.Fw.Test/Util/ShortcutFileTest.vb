Imports System.IO
Imports NUnit.Framework

Namespace Util
    Public MustInherit Class ShortcutFileTest

        Private ReadOnly temporaryPath As String = Path.Combine(Path.GetTempPath, GetType(ShortcutFileTest).Name)
        <SetUp()> Public Overridable Sub SetUp()
            FileUtil.MakeFolderIfNotExists(temporaryPath)
        End Sub

        <TearDown()> Public Overridable Sub TearDown()
            FileUtil.DeleteFolderIfExist(temporaryPath, recursive:=True)
        End Sub

        Public Class DefaultTest : Inherits ShortcutFileTest

            <Test()> Public Sub 拡張子未指定なら_ローカルファイルがリンク先なら_拡張子はlnkになる()
                Using sut As New ShortcutFile(Path.Combine(temporaryPath, "abc"), "c:\windows\system32\cmd.exe")
                    sut.Save()

                    Dim expected As String = Path.Combine(temporaryPath, "abc.lnk")
                    Dim actual As String = sut.FullName
                    Assert.That(actual, [Is].EqualTo(expected))
                    Assert.That(File.Exists(expected), [Is].True)
                End Using
            End Sub

            <Test()> Public Sub 拡張子未指定なら_httpがリンク先なら_拡張子はurlになる( _
                    <Values("http:", "https:", "file:", "javascript:")> ByVal prefix As String)
                Using sut As New ShortcutFile(Path.Combine(temporaryPath, "abc"), prefix & "//google.co.jp/")
                    sut.Save()

                    Dim expected As String = Path.Combine(temporaryPath, "abc.url")
                    Dim actual As String = sut.FullName
                    Assert.That(actual, [Is].EqualTo(expected))
                    Assert.That(File.Exists(expected), [Is].True)
                End Using
            End Sub

            <Test()> Public Sub 拡張子lnkで_ローカルファイルがリンク先なら_拡張子はそのまま()
                Using sut As New ShortcutFile(Path.Combine(temporaryPath, "abc.lnk"), "c:\windows\system32\cmd.exe")
                    sut.Save()

                    Dim expected As String = Path.Combine(temporaryPath, "abc.lnk")
                    Dim actual As String = sut.FullName
                    Assert.That(actual, [Is].EqualTo(expected))
                    Assert.That(File.Exists(expected), [Is].True)
                End Using
            End Sub

            <Test()> Public Sub 拡張子urlで_httpがリンク先なら_拡張子はそのまま( _
                    <Values("http:", "https:", "file:", "javascript:")> ByVal prefix As String)
                Using sut As New ShortcutFile(Path.Combine(temporaryPath, "abc.url"), prefix & "//google.co.jp/")
                    sut.Save()

                    Dim expected As String = Path.Combine(temporaryPath, "abc.url")
                    Dim actual As String = sut.FullName
                    Assert.That(actual, [Is].EqualTo(expected))
                    Assert.That(File.Exists(expected), [Is].True)
                End Using
            End Sub

            <Test()> Public Sub TargetPathが無くとも_ショートカットは作成できる()
                Using sut As New ShortcutFile(Path.Combine(temporaryPath, "abc.lnk"))
                    sut.Save()

                    Dim expected As String = Path.Combine(temporaryPath, "abc.lnk")
                    Assert.That(File.Exists(expected), [Is].True)
                End Using
            End Sub

        End Class

        Public Class ショートカットの読み込みTest : Inherits ShortcutFileTest

            Private testingLnk As String
            Private Const TESTING_TARGET_PATH As String = "c:\windows\system32\cmd.exe"

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                testingLnk = Path.Combine(temporaryPath, "abc.lnk")
            End Sub

            <Test()> Public Sub TargetPathの参照ができる()
                ShortcutFile.Create(testingLnk, TESTING_TARGET_PATH)
                Using sut As New ShortcutFile(testingLnk)
                    Assert.That(sut.TargetPath.ToLower, [Is].EqualTo(TESTING_TARGET_PATH))
                End Using
            End Sub

            <Test()> Public Sub TargetPathの参照ができる2()
                Using s As New ShortcutFile(testingLnk) With {.TargetPath = TESTING_TARGET_PATH}
                    s.Save()
                End Using
                Using sut As New ShortcutFile(testingLnk)
                    Assert.That(sut.TargetPath.ToLower, [Is].EqualTo(TESTING_TARGET_PATH))
                End Using
            End Sub

            <Test()> Public Sub TargetPathの参照ができる3()
                Using s As New ShortcutFile(testingLnk, TESTING_TARGET_PATH)
                    s.Save()
                End Using
                Using sut As New ShortcutFile(testingLnk)
                    Assert.That(sut.TargetPath.ToLower, [Is].EqualTo(TESTING_TARGET_PATH))
                End Using
            End Sub

            <Test()> Public Sub FullNameの参照ができる()
                ShortcutFile.Create(testingLnk, TESTING_TARGET_PATH)
                Using sut As New ShortcutFile(testingLnk)
                    Dim actual As String = sut.FullName
                    Assert.That(actual, [Is].EqualTo(testingLnk))
                End Using
            End Sub

            <Test()> Public Sub IconLocationの参照ができる()
                ShortcutFile.Create(testingLnk, TESTING_TARGET_PATH)
                Using sut As New ShortcutFile(testingLnk)
                    Dim actual As String = sut.IconLocation
                    Assert.That(actual, [Is].EqualTo(TESTING_TARGET_PATH & ",0"))
                End Using
            End Sub

            <Test()> Public Sub IconLocationの設定と参照ができる()
                Using s As New ShortcutFile(testingLnk, TESTING_TARGET_PATH) With {.IconLocation = "c:\windows\notepad.exe,0"}
                    s.Save()
                End Using
                Using sut As New ShortcutFile(testingLnk)
                    Dim actual As String = sut.IconLocation
                    Assert.That(actual, [Is].EqualTo("c:\windows\notepad.exe,0"))
                End Using
            End Sub

            <Test()> Public Sub Descriptionの設定と参照ができる()
                Using s As New ShortcutFile(testingLnk, TESTING_TARGET_PATH) With {.Description = "hoge"}
                    s.Save()
                End Using
                Using sut As New ShortcutFile(testingLnk)
                    Dim actual As String = sut.Description
                    Assert.That(actual, [Is].EqualTo("hoge"))
                End Using
            End Sub

            <Test()> Public Sub Argumentsの設定と参照ができる()
                Using s As New ShortcutFile(testingLnk, TESTING_TARGET_PATH) With {.Arguments = "fuga"}
                    s.Save()
                End Using
                Using sut As New ShortcutFile(testingLnk)
                    Dim actual As String = sut.Arguments
                    Assert.That(actual, [Is].EqualTo("fuga"))
                End Using
            End Sub

            <Test()> Public Sub WorkingDirectoryの設定と参照ができる()
                Dim tempPath As String = Path.GetTempPath
                Using s As New ShortcutFile(testingLnk, TESTING_TARGET_PATH) With {.WorkingDirectory = tempPath}
                    s.Save()
                End Using
                Using sut As New ShortcutFile(testingLnk)
                    Dim actual As String = sut.WorkingDirectory
                    Assert.That(actual, [Is].EqualTo(tempPath))
                End Using
            End Sub

            <Test()> Public Sub WorkingDirectoryの設定と参照ができる( _
                    <Values(ShortcutFile.Style.Normal, ShortcutFile.Style.Maximum, ShortcutFile.Style.Minimum)> ByVal style As ShortcutFile.Style)
                Using s As New ShortcutFile(testingLnk, TESTING_TARGET_PATH) With {.WindowStyle = style}
                    s.Save()
                End Using
                Using sut As New ShortcutFile(testingLnk)
                    Dim actual As ShortcutFile.Style = sut.WindowStyle
                    Assert.That(actual, [Is].EqualTo(style))
                End Using
            End Sub

            <Test()> Public Sub hotkeyの設定と参照ができる()
                Const TESTING_HOTKEY As String = "Ctrl+Shift+F"
                Using s As New ShortcutFile(testingLnk, TESTING_TARGET_PATH) With {.hotkey = TESTING_HOTKEY}
                    s.Save()
                End Using
                Using sut As New ShortcutFile(testingLnk)
                    Dim actual As String = sut.Hotkey
                    Assert.That(actual, [Is].EqualTo(TESTING_HOTKEY))
                End Using
            End Sub

        End Class

    End Class
End Namespace
