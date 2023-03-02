Imports Fhi.Fw.Db
Imports NUnit.Framework

Namespace Util
    Public Class NameGetterTest

        Private getter As NameGetter

        <SetUp()> Public Sub SetUp()
            getter = New NameGetter
        End Sub

        <Test()> Public Sub 正常系()
            ' プロパティ名を取得したいVOのインスタンスを作成
            Dim vo As New SampleVo

            getter.Is(vo).By(vo.HogeDate, vo.HogeId, vo.HogeName, vo.HogeDecimal)
            Assert.AreEqual(4, getter.Result.Length)
            Assert.AreEqual("HOGE_DATE", getter.Result(0))
            Assert.AreEqual("HOGE_ID", getter.Result(1))
            Assert.AreEqual("HOGE_NAME", getter.Result(2))
            Assert.AreEqual("HOGE_DECIMAL", getter.Result(3))
        End Sub
    End Class
End Namespace