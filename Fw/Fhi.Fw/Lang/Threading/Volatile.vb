Namespace Lang.Threading
    ''' <summary>
    ''' スレッド間で変数の値に相違がないように変数アクセスをスレッドセーフにする
    ''' </summary>
    ''' <remarks>.NET4系でVolatileが実装されたので、委譲する</remarks>
    <Obsolete("System.Threading.Volatileを使ったほうが良さそう")> Public Class Volatile

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <typeparam name="T">変数の型</typeparam>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(Of T As Class)(ByRef address As T) As T
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <typeparam name="T">変数の型</typeparam>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(Of T As Class)(ByRef address As T, ByVal value As T)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As Boolean) As Boolean
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As Byte) As Byte
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As Double) As Double
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As Int16) As Int16
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As Int32) As Int32
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As Int64) As Int64
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As IntPtr) As IntPtr
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As SByte) As SByte
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As Single) As Single
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As UInt16) As UInt16
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As UInt32) As UInt32
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As UInt64) As UInt64
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数を読み込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Read(ByRef address As UIntPtr) As UIntPtr
            Return System.Threading.Volatile.Read(address)
        End Function

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As Boolean, ByVal value As Boolean)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As Byte, ByVal value As Byte)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As Double, ByVal value As Double)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As Int16, ByVal value As Int16)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As Int32, ByVal value As Int32)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As Int64, ByVal value As Int64)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As IntPtr, ByVal value As IntPtr)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As SByte, ByVal value As SByte)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As Single, ByVal value As Single)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As UInt16, ByVal value As UInt16)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As UInt32, ByVal value As UInt32)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As UInt64, ByVal value As UInt64)
            System.Threading.Volatile.Write(address, value)
        End Sub

        ''' <summary>
        ''' 変数に書き込む
        ''' </summary>
        ''' <param name="address">変数のアドレス</param>
        ''' <param name="value">書き込み値</param>
        ''' <remarks></remarks>
        Public Shared Sub Write(ByRef address As UIntPtr, ByVal value As UIntPtr)
            System.Threading.Volatile.Write(address, value)
        End Sub

    End Class
End Namespace