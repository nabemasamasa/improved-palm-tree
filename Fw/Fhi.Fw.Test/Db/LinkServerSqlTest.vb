Imports NUnit.Framework

Namespace Db
    Public MustInherit Class LinkServerSqlTest

        Private Const TESTING_LINK_SERVER As String = "ABC_XYZ"

        Private sut As LinkServerSql

        <SetUp()> Public Overridable Sub SetUp()
            sut = New LinkServerSql(TESTING_LINK_SERVER)
        End Sub

        Public Class SqlSelectTest : Inherits LinkServerSqlTest

            <Test(), Sequential()> Public Sub Select文用のLinkServer向けSQL構文変換できる( _
                    <Values("SELECT * FROM HOGE", "SELECT C FROM TABLE WHERE A='b'")> ByVal boundSql As String, _
                    <Values("SELECT * FROM OPENQUERY( ABC_XYZ, 'SELECT * FROM HOGE')", "SELECT * FROM OPENQUERY( ABC_XYZ, 'SELECT C FROM TABLE WHERE A=''b''')")> ByVal expected As String)
                Assert.That(sut.ConvertOpenqueryFrom(boundSql), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub ConvertFourPartNamesFrom_4部構成SQLにできる( _
                    <Values("select * from Hoge", "SELECT F.* FROM [FUGA] F")> ByVal boundSql As String, _
                    <Values("select * from [ABC_XYZ].[ABC].[XYZ].Hoge", "SELECT F.* FROM [ABC_XYZ].[ABC].[XYZ].[FUGA] F")> ByVal expected As String)

                Assert.That(sut.ConvertFourPartNamesFrom(boundSql), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub ConvertFourPartNamesFrom_特定のjoinだけ4部構成SQLにできる( _
                    <Values("select * from Hoge H inner join Fuga F on H.id = F.id", "select * from  PIYO H inner join  Fuga F on H.id = F.id")> ByVal boundSql As String, _
                    <Values("select * from [ABC_XYZ].[ABC].[XYZ].Hoge H inner join [ABC_XYZ].[ABC].[XYZ].Fuga F on H.id = F.id", "select * from  [ABC_XYZ].[ABC].[XYZ].PIYO H inner join  [ABC_XYZ].[ABC].[XYZ].Fuga F on H.id = F.id")> ByVal expected As String)

                Assert.That(sut.ConvertFourPartNamesFrom(boundSql), [Is].EqualTo(expected))
            End Sub

        End Class

        Public Class SqlInsertTest : Inherits LinkServerSqlTest

            <TestCase("INSERT INTO TABLE (ID) VALUES (1)", "INSERT OPENQUERY( ABC_XYZ, 'SELECT ID FROM TABLE WHERE 1=0') VALUES (1)")>
            <TestCase("INSERT INTO TABLE (ID) VALUES (1), (2), (3)", "INSERT OPENQUERY( ABC_XYZ, 'SELECT ID FROM TABLE WHERE 1=0') VALUES (1), (2), (3)")>
            <TestCase("INSERT  INTO  TABLE  VALUES  (  2  )  ", "INSERT OPENQUERY( ABC_XYZ, 'SELECT * FROM TABLE WHERE 1=0') VALUES  (  2  )")>
            <TestCase("INSERT INTO TABLE (ID, NAME) VALUES (3, 'AB')  ", "INSERT OPENQUERY( ABC_XYZ, 'SELECT ID, NAME FROM TABLE WHERE 1=0') VALUES (3, 'AB')")>
            Public Sub LinkServer向けSQLに構文変換できる(boundSql As String, expected As String)
                Assert.That(sut.ConvertOpenqueryFrom(boundSql), [Is].EqualTo(expected))
            End Sub

            <TestCase("INSERT INTO HOGE (FUGA) VALUES ('A')", "INSERT INTO [ABC_XYZ].[ABC].[XYZ].HOGE (FUGA) VALUES ('A')")>
            <TestCase("INSERT INTO HOGE (FUGA) VALUES ('A'), ('B'), ('C')", "INSERT INTO [ABC_XYZ].[ABC].[XYZ].HOGE (FUGA) VALUES ('A'), ('B'), ('C')")>
            <TestCase("insert into foo (bar, baz) values (123, 'xyz'", "insert into [ABC_XYZ].[ABC].[XYZ].foo (bar, baz) values (123, 'xyz'")>
            Public Sub ConvertFourPartNamesFrom_4部構成SQLにできる(boundSql As String, expected As String)

                Assert.That(sut.ConvertFourPartNamesFrom(boundSql), [Is].EqualTo(expected))
            End Sub

        End Class

        Public Class SqlUpdateTest : Inherits LinkServerSqlTest

            <Test(), Sequential()> Public Sub LinkServer向けSQLに構文変換できる( _
                    <Values("UPDATE TABLE SET A=1 WHERE B=2", _
                            "UPDATE  TABLE  SET  A=1  WHERE  B=2", _
                            "UPDATE TABLE SET C='c' WHERE D='d'")> ByVal boundSql As String, _
                    <Values("UPDATE OPENQUERY( ABC_XYZ, 'SELECT * FROM TABLE WHERE B=2') SET A=1", _
                            "UPDATE OPENQUERY( ABC_XYZ, 'SELECT * FROM TABLE WHERE  B=2') SET  A=1", _
                            "UPDATE OPENQUERY( ABC_XYZ, 'SELECT * FROM TABLE WHERE D=''d''') SET C='c'")> ByVal expected As String)
                Assert.That(sut.ConvertOpenqueryFrom(boundSql), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub ConvertFourPartNamesFrom_4部構成SQLにできる( _
                    <Values("UPDATE HOGE SET FUGA = 'A' WHERE FUGA = 'B'", "update foo set bar=123 where baz='xyz'")> ByVal boundSql As String, _
                    <Values("UPDATE [ABC_XYZ].[ABC].[XYZ].HOGE SET FUGA = 'A' WHERE FUGA = 'B'", "update [ABC_XYZ].[ABC].[XYZ].foo set bar=123 where baz='xyz'")> ByVal expected As String)

                Assert.That(sut.ConvertFourPartNamesFrom(boundSql), [Is].EqualTo(expected))
            End Sub

        End Class

        Public Class SqlDeleteTest : Inherits LinkServerSqlTest

            <Test(), Sequential()> Public Sub LinkServer向けSQLに構文変換できる( _
                    <Values("DELETE FROM TABLE WHERE A=1", _
                            "DELETE  FROM  TABLE  WHERE  A=1", _
                            "DELETE FROM TABLE WHERE B='CD'")> ByVal boundSql As String, _
                    <Values("DELETE OPENQUERY( ABC_XYZ, 'SELECT * FROM TABLE WHERE A=1')", _
                            "DELETE OPENQUERY( ABC_XYZ, 'SELECT * FROM TABLE WHERE  A=1')", _
                            "DELETE OPENQUERY( ABC_XYZ, 'SELECT * FROM TABLE WHERE B=''CD''')")> ByVal expected As String)
                Assert.That(sut.ConvertOpenqueryFrom(boundSql), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub ConvertFourPartNamesFrom_4部構成SQLにできる( _
                    <Values("DELETE FROM HOGE WHERE FUGA = 'A'", "delete from foo where baz='xyz'")> ByVal boundSql As String, _
                    <Values("DELETE FROM [ABC_XYZ].[ABC].[XYZ].HOGE WHERE FUGA = 'A'", "delete from [ABC_XYZ].[ABC].[XYZ].foo where baz='xyz'")> ByVal expected As String)

                Assert.That(sut.ConvertFourPartNamesFrom(boundSql), [Is].EqualTo(expected))
            End Sub

        End Class

    End Class
End Namespace
