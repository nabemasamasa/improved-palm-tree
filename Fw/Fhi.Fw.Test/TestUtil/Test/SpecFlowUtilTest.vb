Imports BoDi
Imports TechTalk.SpecFlow.Configuration
Imports NUnit.Framework
Imports TechTalk.SpecFlow

Namespace TestUtil.Test
    Public MustInherit Class SpecFlowUtilTest

#Region "Fake classes..."
        Private Class TestRunContainerBuilder
            Public Shared Function CreateContainer(Optional ByVal configurationProvider As IRuntimeConfigurationProvider = Nothing) As IObjectContainer
                Dim aType As Type = NonPublicUtil.GetTypeOfNonPublic(Of ScenarioContext)("TechTalk.SpecFlow.Infrastructure.TestRunContainerBuilder")
                Dim result As Object = NonPublicUtil.InvokeMethodOfShared(aType, "CreateContainer", configurationProvider)
                Return DirectCast(result, IObjectContainer)
            End Function
        End Class

        Private Class TestTestRunnerFactory
            Public Shared Function CreateTestRunner(ByRef container As IObjectContainer, Optional ByVal registerMocks As Action(Of IObjectContainer) = Nothing) As TestRunner
                container = TestRunContainerBuilder.CreateContainer()
                If registerMocks IsNot Nothing Then
                    registerMocks(container)
                End If
                Return DirectCast(container.Resolve(Of ITestRunner)(), TestRunner)
            End Function
            Public Shared Function CreateTestRunner(Optional ByVal registerMocks As Action(Of IObjectContainer) = Nothing) As TestRunner
                Dim container As IObjectContainer = Nothing
                Return CreateTestRunner(container, registerMocks)
            End Function
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()

        End Sub

        <TearDown()> Public Overridable Sub TearDown()

        End Sub

        Public Class ScenarioContextTest : Inherits SpecFlowUtilTest

            Private testRunner As TestRunner

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                NonPublicUtil.SetPropertyOfShared(Of ScenarioContext)("Current", CreateScenarioContext())
            End Sub

            Private Function CreateScenarioContext(Optional ByVal registerMocks As Action(Of IObjectContainer) = Nothing) As ScenarioContext
                Dim container As IObjectContainer = Nothing
                testRunner = TestTestRunnerFactory.CreateTestRunner(container, registerMocks)
                Return NonPublicUtil.NewInstance(Of ScenarioContext)(New ScenarioInfo("sample scenario", New String() {}), testRunner, container)
            End Function

        End Class

        Public Class InitializerTest : Inherits SpecFlowUtilTest

            <TearDown()> Public Overrides Sub TearDown()
                If CollectionUtil.IsNotEmpty(sutScenarioContext) Then
                    SpecFlowUtil.FinalizeScenarioContext()
                    Assert.That(sutScenarioContext, [Is].Empty)
                End If

                If CollectionUtil.IsNotEmpty(sutFeatureContext) Then
                    SpecFlowUtil.FinalizeFeatureContext()
                    Assert.That(sutFeatureContext, [Is].Empty)
                End If
            End Sub

            Private Shared ReadOnly Property sutScenarioContext As ScenarioContext
                Get
                    Return ScenarioContext.Current
                End Get
            End Property

            Private Shared ReadOnly Property sutFeatureContext As FeatureContext
                Get
                    Return FeatureContext.Current
                End Get
            End Property

            <Test()> Public Sub ScenarioContextを初期化できる()
                SpecFlowUtil.InitializeScenarioContext()
                Assert.That(sutScenarioContext, [Is].Not.Null)
                Assert.That(sutScenarioContext.Count, [Is].EqualTo(0))

                Dim title As String = sutScenarioContext.ScenarioInfo.Title
                Assert.That(title, [Is].EqualTo("SpecFlowUtilScenarioTitle"))
            End Sub

            <Test()> Public Sub FeatureContextを初期化できる()
                SpecFlowUtil.InitializeFeatureContext("CustomSpecFlowUtilFeatureTitle", "CustomSpecFlowUtilFeatureDescription")
                Assert.That(sutFeatureContext, [Is].Not.Null)
                Assert.That(sutFeatureContext.Count, [Is].EqualTo(0))

                Dim title As String = sutFeatureContext.FeatureInfo.Title
                Dim description As String = sutFeatureContext.FeatureInfo.Description
                Assert.That(title, [Is].EqualTo("CustomSpecFlowUtilFeatureTitle"))
                Assert.That(description, [Is].EqualTo("CustomSpecFlowUtilFeatureDescription"))
            End Sub

        End Class

        Public Class EvaluateValueTest : Inherits SpecFlowUtilTest

            <Test(), Sequential()> Public Sub シンプルな文字列ならそのまま返す( _
                    <Values("aiueo", "あいうえお", "1 2 3")> ByVal arg As String, _
                    <Values("aiueo", "あいうえお", "1 2 3")> ByVal expected As String)
                Assert.That(SpecFlowUtil.EvaluateValue(arg), [Is].EqualTo(expected))
            End Sub

        End Class

        Public Class DateFunctionOnEvaluateValueTest : Inherits SpecFlowUtilTest

            <Test()> Public Sub 引数なしはスラッシュ付きのyyyymmdd()
                Assert.That(SpecFlowUtil.EvaluateValue("@Date"), [Is].EqualTo(DateTime.Now.ToString("yyyy/MM/dd")))
            End Sub

            <Test()> Public Sub 日付書式引数なら_それで出力する()
                Assert.That(SpecFlowUtil.EvaluateValue("@Date(yyyyMMdd)"), [Is].EqualTo(DateTime.Now.ToString("yyyyMMdd")))
            End Sub

            <Test()> Public Sub 数値引数なら_システム日付に日数を加算して出力する()
                Assert.That(SpecFlowUtil.EvaluateValue("@Date(+2)"), [Is].EqualTo(DateTime.Now.AddDays(2).ToString("yyyy/MM/dd")))
            End Sub

            <Test()> Public Sub Date関数_文字列と繋がってもOK()
                Assert.That(SpecFlowUtil.EvaluateValue("ABC@Date"), [Is].EqualTo("ABC" & DateTime.Now.ToString("yyyy/MM/dd")), "文字列と繋がっててもOK")
                Assert.That(SpecFlowUtil.EvaluateValue("@Datexyz"), [Is].EqualTo(DateTime.Now.ToString("yyyy/MM/dd") & "xyz"), "文字列と繋がっててもOK")
                Assert.That(SpecFlowUtil.EvaluateValue("ABC@Date(yyyyMMdd)"), [Is].EqualTo("ABC" & DateTime.Now.ToString("yyyyMMdd")), "文字列と繋がっててもOK")
                Assert.That(SpecFlowUtil.EvaluateValue("@Date(yyyyMMdd)xyz"), [Is].EqualTo(DateTime.Now.ToString("yyyyMMdd") & "xyz"), "文字列と繋がっててもOK")
            End Sub

        End Class

        Public Class ZenkakuFunctionOnEvaluateValueTest : Inherits SpecFlowUtilTest
            <Test()> Public Sub 引数文字列を全角にして返す()
                Assert.That(SpecFlowUtil.EvaluateValue("@Zenkaku(A1ｦ)"), [Is].EqualTo("Ａ１ヲ"))
            End Sub

            <Test()> Public Sub 全角が存在する記号は全角にして返す()
                Assert.That(SpecFlowUtil.EvaluateValue("@Zenkaku(&$)"), [Is].EqualTo("＆＄"))
            End Sub

            <Test()> Public Sub 引数の関数戻り値を全角にして返す()
                Assert.That(SpecFlowUtil.EvaluateValue("@Zenkaku(@Date(yyyyMMdd))"), [Is].EqualTo(StringUtil.ToZenkaku(DateTime.Now.ToString("yyyyMMdd"))))
            End Sub
        End Class

        Public Class HankakuFunctionOnEvaluateValueTest : Inherits SpecFlowUtilTest
            <Test()> Public Sub 引数文字列を半角にして返す()
                Assert.That(SpecFlowUtil.EvaluateValue("@Hankaku(Ａ１ヲ)"), [Is].EqualTo("A1ｦ"))
            End Sub

            <Test()> Public Sub 半角が存在する記号は半角にして返す()
                Assert.That(SpecFlowUtil.EvaluateValue("@Hankaku(＆＄)"), [Is].EqualTo("&$"))
            End Sub

            <Test()> Public Sub 半角の無い記号_ひらがな_漢字はそのまま返す()
                Assert.That(SpecFlowUtil.EvaluateValue("@Hankaku(■あ亜)"), [Is].EqualTo("■あ亜"))
            End Sub

            <Test()> Public Sub 引数の関数戻り値を半角にして返す()
                Assert.That(SpecFlowUtil.EvaluateValue("@Hankaku(@Zenkaku(AA))"), [Is].EqualTo("AA"))
            End Sub
        End Class

        Public Class GetFunctionOnEvaluateValueTest : Inherits ScenarioContextTest

            <Test()> Public Sub カレントシナリオの値を取得できる()

                SpecFlowUtil.SetCurrentScenarioValue("Key1", "10")

                Assert.That(SpecFlowUtil.EvaluateValue("@Get(Key1)"), [Is].EqualTo("10"))
            End Sub

            <Test()> Public Sub 存在しない変数名は例外になる()
                Try
                    SpecFlowUtil.EvaluateValue("@Get(Key1)")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("変数 ""Key1"" はみつからない"))
                End Try

            End Sub

        End Class

        Public Class ExpandFunctionOnEvaluateValueTest : Inherits SpecFlowUtilTest

            <Test()> Public Sub 環境変数を展開する()
                Dim actual As String = SpecFlowUtil.EvaluateValue("@Expand(%TEMP%\abc.txt)")
                Assert.That(actual, [Is].StringEnding("\abc.txt"))
                Assert.That(actual, [Is].Not.StringMatching("%TEMP%"))
            End Sub

        End Class

        Public Class ConvertTableToTwoJaggedArrayTest : Inherits ScenarioContextTest

            <Test()> Public Sub 環境変数を展開する()
                Dim aTable As New Table("a", "b", "c")
                aTable.AddRow("1", "a", "A")
                aTable.AddRow("2", "b", "B")

                Dim actuals As String()() = SpecFlowUtil.ConvTableToTwoJaggedArray(aTable)

                Assert.That(actuals, [Is].EqualTo(New String()() {New String() {"1", "a", "A"}, _
                                                                  New String() {"2", "b", "B"}}))
            End Sub

            <Test()> Public Sub 環境変数を展開する2()
                SpecFlowUtil.SetCurrentScenarioValue("Key1", "10")

                Dim aTable As New Table("a", "b", "c)")
                aTable.AddRow("1", "a", "A")
                aTable.AddRow("2", "b", "@Get(Key1)")

                Dim actuals As String()() = SpecFlowUtil.ConvTableToTwoJaggedArray(aTable)

                Assert.That(actuals, [Is].EqualTo(New String()() {New String() {"1", "a", "A"}, _
                                                                   New String() {"2", "b", "10"}}))
            End Sub

        End Class

        Public Class LeftTest : Inherits ScenarioContextTest

            <Test()> Public Sub カレントシナリオの値を取得できる()

                SpecFlowUtil.SetCurrentScenarioValue("Key1", "1234567890")
                Assert.That(SpecFlowUtil.EvaluateValue("@Left(Key1, 5)"), [Is].EqualTo("12345"))
            End Sub

            <Test()> Public Sub 引数は_変数名_長さ_が必要なので_それ以外はエラーになる(<Values("Key1", "Key1,1,2", "Key1, hoge", "")> invalidArgs As String)
                SpecFlowUtil.SetCurrentScenarioValue("Key1", "1234567890")
                Try
                    SpecFlowUtil.EvaluateValue("@Left(" & invalidArgs & ")")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].StringStarting("引数は @Left(変数, 長さ) にすべき"))
                End Try
            End Sub

            <Test()> Public Sub カレントシナリオに存在しない変数名はエラーになる()
                Try
                    SpecFlowUtil.EvaluateValue("@Left(Key1,2)")
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].EqualTo("変数 ""Key1"" はみつからない"))
                End Try

            End Sub

        End Class

    End Class
End Namespace
