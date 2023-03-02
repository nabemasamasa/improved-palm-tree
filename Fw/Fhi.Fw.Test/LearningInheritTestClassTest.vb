Imports NUnit.Framework

Public MustInherit Class LearningInheritTestClassTest
    <Test()> Public Sub Hoge_全てのサブクラスで実行するテスト()
        Debug.Print(String.Format("{0}#Hoge() ... has BaseClass", Me.GetType.Name))
        Assert.IsTrue(True)
    End Sub
End Class

Public Class SubFugaLearningInheritTestClassTest : Inherits LearningInheritTestClassTest

    <Test()> Public Sub Fuga_サブクラス固有のテスト()
        Assert.AreEqual("SubFugaLearningInheritTestClassTest", Me.GetType.Name)
    End Sub
End Class

Public Class SubPiyoLearningInheritTestClassTest : Inherits LearningInheritTestClassTest

    <Test()> Public Sub Piyo_サブクラス固有のテスト()
        Assert.AreEqual("SubPiyoLearningInheritTestClassTest", Me.GetType.Name)
    End Sub
End Class
