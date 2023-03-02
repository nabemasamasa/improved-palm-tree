Imports NUnit.Framework
Imports System.Threading
Imports Fhi.Fw.Lang.Threading

Namespace Lang.Threading
    Public Class ThreadLocalTest

        Private Class ThreadTestHelper(Of T)

            Private storage As ThreadLocal(Of T)
            Public Result As T

            Public Sub New(ByVal storage As ThreadLocal(Of T))
                Me.storage = storage
            End Sub

            Public Sub Execute()
                Dim t As Thread = New Thread(New ThreadStart(AddressOf Run))
                t.Start()
                t.Join()
            End Sub

            Private Sub Run()
                Result = storage.Get()
            End Sub
        End Class

        Dim storageStr As ThreadLocal(Of String)
        Dim storageObj As ThreadLocal(Of Object)

        <SetUp()> Public Sub SetUp()
            storageStr = New ThreadLocal(Of String)
            storageObj = New ThreadLocal(Of Object)
        End Sub

        <TearDown()> Public Sub TearDown()
            storageStr.Remove()
            storageObj.Remove()
        End Sub

        <Test()> Public Sub インスタンス生成直後は空っぽ()

            Dim storage As New ThreadLocal(Of String)

            Assert.IsNull(storage.Get())
        End Sub

        <Test()> Public Sub Set_Setした値をGetできる()

            storageStr.Set("aiueo")

            Assert.AreEqual("aiueo", storageStr.Get)
        End Sub

        <Test()> Public Sub Set_既に値があっても上書き出来る()

            Dim value1 As New Object
            Dim value2 As New Object

            storageObj.Set(value1)
            Assert.AreSame(value1, storageObj.Get)
            storageObj.Set(value2)
            Assert.AreSame(value2, storageObj.Get)
            storageObj.Set(Nothing)
            Assert.AreSame(Nothing, storageObj.Get)
        End Sub

        <Test()> Public Sub 当スレッドでSetした値は別スレッドで見られない()
            Dim value1 As New Object
            storageObj.Set(value1)

            Dim helper As New ThreadTestHelper(Of Object)(storageObj)
            helper.Execute()

            Assert.AreNotSame(value1, helper.Result)
            Assert.AreSame(Nothing, helper.Result)
        End Sub

        <Test()> Public Sub 同一スレッドでもインスタンスが違うなら別インスタンスの値は見られない_当たり前()

            Dim local1 As New ThreadLocal(Of Object)
            Dim local2 As New ThreadLocal(Of Object)

            Try
                Dim value1 As New Object
                local1.Set(value1)

                Assert.AreNotSame(value1, local2.Get)
                Assert.AreSame(Nothing, local2.Get)
            Finally
                local1.Remove()
                local2.Remove()
            End Try
        End Sub

        <Test()> Public Sub DlgtInitialValue_初期値設定デリゲートを設定すれば初期値が返る()

            Dim local As New ThreadLocal(Of String)(Function() "aiueo")

            Try
                Assert.AreEqual("aiueo", local.Get, "初期値が返る")
            Finally
                local.Remove()
            End Try
        End Sub

    End Class
End Namespace