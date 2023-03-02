Imports NUnit.Framework
Imports Fhi.Fw.Util.Search

Namespace Util.Search
    Public Class BMSearchTest
        <Test()> Public Sub hoge()
            Assert.AreEqual(2, BMSearch.Search("abcabedcabcdcaabcdcabccc", "abcdc"))
        End Sub

        <Test()> Public Sub fuga()
            Assert.AreEqual(1, BMSearch.Search("abcdcdcaabc", "abcdc"))
        End Sub
    End Class
End Namespace