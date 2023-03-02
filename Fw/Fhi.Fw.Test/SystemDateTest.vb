Imports NUnit.Framework

<TestFixture()> Public Class SystemDateTest

    Private Class FakeBehavior : Implements SystemDate.IBehavior

        Private _date As DateTime
        Public Sub New(ByVal dateString As String)
            _date = CDate(dateString)
        End Sub

        Public Function GetDateTimeNow() As Date? Implements SystemDate.IBehavior.GetDateTimeNow
            Return _date
        End Function
    End Class

    <Test()> Public Sub TestCurrentDateTime()
        Dim d As New SystemDate(New FakeBehavior("2010/05/05 11:22:33"))
        Assert.AreEqual(CDate("2010/05/05 11:22:33"), d.CurrentDateTime)
    End Sub

    <Test()> Public Sub TestCurrentDateDbFormat()
        Dim d As New SystemDate(New FakeBehavior("2010/05/05 11:22:33"))
        Assert.AreEqual("2010-05-05", d.CurrentDateDbFormat)
    End Sub

    <Test()> Public Sub TestCurrentDateAsInteger()
        Dim d As New SystemDate(New FakeBehavior("2010/05/05 11:22:33"))
        Assert.AreEqual(20100505, d.CurrentDateAsInteger)
    End Sub

    <Test()> Public Sub TestCurrentTimeDbFormat()
        Dim d As New SystemDate(New FakeBehavior("2010/05/05 11:22:33"))
        Assert.AreEqual("11:22:33", d.CurrentTimeDbFormat)
    End Sub

    <Test()> Public Sub TestCurrentTimeAsInteger()
        Dim d As New SystemDate(New FakeBehavior("2010/05/05 11:22:33"))
        Assert.AreEqual(112233, d.CurrentTimeAsInteger)
    End Sub

End Class
