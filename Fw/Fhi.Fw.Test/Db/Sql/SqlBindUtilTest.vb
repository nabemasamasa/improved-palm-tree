Imports NUnit.Framework
Imports Fhi.Fw.Db.Sql

Namespace Db.Sql
    Public Class SqlBindUtilTest

        <Test()> Public Sub ConvPropertyNameToInternalBindName_通常()
            Assert.AreEqual("@Hoge", SqlBindUtil.ConvPropertyNameToInternalBindName("Hoge"))
            Assert.AreEqual("@Hoge$Piyo", SqlBindUtil.ConvPropertyNameToInternalBindName("Hoge.Piyo"))
        End Sub

        <Test()> Public Sub ConvPropertyNameToInternalBindName_配列要素()
            Assert.AreEqual("@Hoge#2", SqlBindUtil.ConvPropertyNameToInternalBindName("Hoge(2)"))
        End Sub

        <Test()> Public Sub ConvPropertyNameToInternalBindName_Indexプロパティ()
            Assert.AreEqual("@Hoge#3", SqlBindUtil.ConvPropertyNameToInternalBindName("Hoge", 3))
            Assert.AreEqual("@Hoge$Piyo#6", SqlBindUtil.ConvPropertyNameToInternalBindName("Hoge.Piyo", 6))
        End Sub

        <Test()> Public Sub ConvBindNameToPropertyName_通常()
            Assert.AreEqual("Hoge", SqlBindUtil.ConvBindNameToPropertyName("@Hoge"))
            Assert.AreEqual("Fuga.Piyo", SqlBindUtil.ConvBindNameToPropertyName("@Fuga.Piyo"))
        End Sub

        <Test()> Public Sub ConvBindNameToPropertyName_Indexプロパティ()
            Assert.AreEqual("Hoge(0)", SqlBindUtil.ConvBindNameToPropertyName("@Hoge(0)"))
            Assert.AreEqual("Hoge(1)", SqlBindUtil.ConvBindNameToPropertyName("@Hoge#1"))
            Assert.AreEqual("Fuga.Piyo(2)", SqlBindUtil.ConvBindNameToPropertyName("@Fuga.Piyo(2)"))
            Assert.AreEqual("Fuga.Piyo(3)", SqlBindUtil.ConvBindNameToPropertyName("@Fuga$Piyo#3"))
            Assert.AreEqual("Fuga(4).Piyo", SqlBindUtil.ConvBindNameToPropertyName("@Fuga(4).Piyo"))
            Assert.AreEqual("Fuga(5).Piyo", SqlBindUtil.ConvBindNameToPropertyName("@Fuga#5$Piyo"))
        End Sub

        <Test()> Public Sub ConvBindNameUserToInternal_通常()
            Assert.AreEqual("@Hoge", SqlBindUtil.ConvBindNameUserToInternal("@Hoge"))
            Assert.AreEqual("@Hoge$Fuga", SqlBindUtil.ConvBindNameUserToInternal("@Hoge.Fuga"))
            Assert.AreEqual("@Hoge#2", SqlBindUtil.ConvBindNameUserToInternal("@Hoge(2)"))
            Assert.AreEqual("@Hoge$Fuga#3", SqlBindUtil.ConvBindNameUserToInternal("@Hoge.Fuga(3)"))
            Assert.AreEqual("@Hoge#5$Fuga#8", SqlBindUtil.ConvBindNameUserToInternal("@Hoge(5).Fuga(8)"))
        End Sub

        <Test()> Public Sub ConvBindNameInternalToUser_通常()
            Assert.AreEqual("@Hoge", SqlBindUtil.ConvBindNameInternalToUser("@Hoge"))
            Assert.AreEqual("@Hoge.Fuga", SqlBindUtil.ConvBindNameInternalToUser("@Hoge$Fuga"))
            Assert.AreEqual("@Hoge(2)", SqlBindUtil.ConvBindNameInternalToUser("@Hoge#2"))
            Assert.AreEqual("@Hoge.Fuga(3)", SqlBindUtil.ConvBindNameInternalToUser("@Hoge$Fuga#3"))
            Assert.AreEqual("@Hoge(5).Fuga(8)", SqlBindUtil.ConvBindNameInternalToUser("@Hoge#5$Fuga#8"))
        End Sub

        <Test()> Public Sub ConvInternalBindNameIfNecessary_直接の要素()
            Dim sql As String = " @Aaa(3)"
            Assert.AreEqual(" @Aaa#3", SqlBindUtil.ConvInternalBindNameIfNecessary(sql))
        End Sub

        <Test()> Public Sub ConvInternalBindNameIfNecessary_直接の要素2()
            Dim sql As String = " @Aaa(2).Bbb"
            Assert.AreEqual(" @Aaa#2$Bbb", SqlBindUtil.ConvInternalBindNameIfNecessary(sql))
        End Sub

        <Test()> Public Sub ConvInternalBindNameIfNecessary_サブプロパティの要素()
            Dim sql As String = " @Aaa.Bbb(8)"
            Assert.AreEqual(" @Aaa$Bbb#8", SqlBindUtil.ConvInternalBindNameIfNecessary(sql))
        End Sub

        <Test()> Public Sub ConvInternalBindNameIfNecessary_直接の要素でサブプロパティの要素()
            Dim sql As String = " @Aaa(1).Bbb(13)"
            Assert.AreEqual(" @Aaa#1$Bbb#13", SqlBindUtil.ConvInternalBindNameIfNecessary(sql))
        End Sub

        <Test()> Public Sub ConvInternalBindNameIfNecessary_複数()
            Dim sql As String = " @Aa(2) @Bbb(3) @Cccc(5)"
            Assert.AreEqual(" @Aa#2 @Bbb#3 @Cccc#5", SqlBindUtil.ConvInternalBindNameIfNecessary(sql))
        End Sub

    End Class
End Namespace