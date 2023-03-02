Imports NUnit.Framework

Namespace Sys

    Public MustInherit Class PrivateProfileHelperTest

        Public Class [Default] : Inherits PrivateProfileHelperTest

            <Test()> Public Sub ConvIniValueToDictionary_1組のみ()
                Dim actual As Dictionary(Of String, String) = PrivateProfileHelper.ConvIniValueToDictionary("a=b")
                Assert.IsTrue(actual.ContainsKey("a"))
                Assert.AreEqual("b", actual("a"))
            End Sub

            <Test()> Public Sub ConvIniValueToDictionary_複数組()
                Dim actual As Dictionary(Of String, String) = PrivateProfileHelper.ConvIniValueToDictionary("a=b,x=y,m=n")
                Assert.IsTrue(actual.ContainsKey("a"))
                Assert.IsTrue(actual.ContainsKey("x"))
                Assert.IsTrue(actual.ContainsKey("m"))
                Assert.AreEqual("b", actual("a"))
                Assert.AreEqual("y", actual("x"))
                Assert.AreEqual("n", actual("m"))
            End Sub

            <Test()> Public Sub ConvDictionaryToIniValue_1組のみ()
                Dim param As New Dictionary(Of String, String)
                param.Add("c", "d")
                Assert.AreEqual("c=d", PrivateProfileHelper.ConvDictionaryToIniValue(param))
            End Sub

            <Test()> Public Sub ConvDictionaryToIniValue_複数組()
                Dim param As New Dictionary(Of String, String)
                param.Add("c", "d")
                param.Add("z", "x")
                param.Add("q", "w")
                Dim actual As String = PrivateProfileHelper.ConvDictionaryToIniValue(param)
                Assert.IsTrue(0 <= actual.IndexOf("c=d"))
                Assert.IsTrue(0 <= actual.IndexOf("q=w"))
                Assert.IsTrue(0 <= actual.IndexOf("z=x"))
            End Sub

        End Class

        Public Class ConvDictionaryToIniValueとConvIniValueToDictionaryの可逆性 : Inherits PrivateProfileHelperTest

            <Test()> Public Sub NothingをIniValueにして戻したら_中身空のインスタンスで返る()
                Dim iniValue As String = PrivateProfileHelper.ConvDictionaryToIniValue(Nothing)
                Dim actuals As Dictionary(Of String, String) = PrivateProfileHelper.ConvIniValueToDictionary(iniValue)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(0, actuals.Count, "中身は空")
            End Sub

            <Test()> Public Sub 中身が空のKeyValueをIniValueにして戻したら_その状態で返る()
                Dim param As New Dictionary(Of String, String)
                Dim iniValue As String = PrivateProfileHelper.ConvDictionaryToIniValue(param)
                Dim actuals As Dictionary(Of String, String) = PrivateProfileHelper.ConvIniValueToDictionary(iniValue)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(0, actuals.Count, "中身は空")
            End Sub

            <Test()> Public Sub ひと組のKeyValueをIniValueにして戻したら_その状態で返る()
                Dim param As New Dictionary(Of String, String)
                param.Add("c", "d")
                Dim iniValue As String = PrivateProfileHelper.ConvDictionaryToIniValue(param)
                Dim actuals As Dictionary(Of String, String) = PrivateProfileHelper.ConvIniValueToDictionary(iniValue)

                Assert.IsNotNull(actuals)
                Assert.IsTrue(actuals.ContainsKey("c"))
                Assert.AreEqual("d", actuals("c"))
            End Sub

            <Test()> Public Sub 複数組のKeyValueをIniValueにして戻したら_その状態で返る()
                Dim param As New Dictionary(Of String, String)
                param.Add("a", "b")
                param.Add("c", "d")
                param.Add("e", "f")
                Dim iniValue As String = PrivateProfileHelper.ConvDictionaryToIniValue(param)
                Dim actuals As Dictionary(Of String, String) = PrivateProfileHelper.ConvIniValueToDictionary(iniValue)

                Assert.IsNotNull(actuals)
                Assert.IsTrue(actuals.ContainsKey("a"))
                Assert.IsTrue(actuals.ContainsKey("c"))
                Assert.IsTrue(actuals.ContainsKey("e"))
                Assert.AreEqual("b", actuals("a"))
                Assert.AreEqual("d", actuals("c"))
                Assert.AreEqual("f", actuals("e"))
            End Sub

        End Class

        Public Class ConvArrayToIniValueとConvIniValueToArrayの可逆性 : Inherits PrivateProfileHelperTest

            <Test()> Public Sub NothingをIniValueにして戻したら_長さ0の配列で返る()
                Dim iniValue As String = PrivateProfileHelper.ConvArrayToIniValue(Nothing)
                Dim actuals As String() = PrivateProfileHelper.ConvIniValueToArray(iniValue)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(0, actuals.Length, "長さ0の配列で返る")
            End Sub

            <Test()> Public Sub 長さ0の配列をIniValueにして戻したら_長さ0の配列で返る()
                Dim iniValue As String = PrivateProfileHelper.ConvArrayToIniValue(New String() {})
                Dim actuals As String() = PrivateProfileHelper.ConvIniValueToArray(iniValue)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(0, actuals.Length, "長さ0の配列で返る")
            End Sub

            <Test()> Public Sub 長さ1の配列で中身が空文字をIniValueにして戻したら_空文字だけを含む配列で返る()
                Dim iniValue As String = PrivateProfileHelper.ConvArrayToIniValue(New String() {""})
                Dim actuals As String() = PrivateProfileHelper.ConvIniValueToArray(iniValue)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Length)
                Assert.AreEqual("", actuals(0), "長さ1で中身空文字の配列で返る")
            End Sub

            <Test()> Public Sub 長さ1の配列で中身がNothingをIniValueにして戻したら_空文字だけを含む配列で返る()
                Dim iniValue As String = PrivateProfileHelper.ConvArrayToIniValue(New String() {Nothing})
                Dim actuals As String() = PrivateProfileHelper.ConvIniValueToArray(iniValue)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(1, actuals.Length)
                Assert.AreEqual("", actuals(0), "長さ1で中身空文字の配列で返る")
            End Sub

        End Class

        'Public Class Hoge : Inherits PrivateProfileHelperTest

        '    Private Const FILE_NAME As String = "c:\hoge.ini"

        '    Private Sub DeleteIniFileIfExists()

        '        If File.Exists(FILE_NAME) Then
        '            File.Delete(FILE_NAME)
        '        End If
        '    End Sub

        '    Private helper As PrivateProfileHelper

        '    <SetUp()> Public Sub Setup()
        '        DeleteIniFileIfExists()
        '        helper = New PrivateProfileHelper(New PrivateProfile(FILE_NAME))
        '    End Sub

        '    <TearDown()> Public Sub TearDown()
        '        DeleteIniFileIfExists()
        '    End Sub

        'End Class

    End Class
End Namespace