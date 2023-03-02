Imports System.Threading

Namespace Util
    ''' <summary>
    ''' タイムアウトキャッシュを担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class SimpleTimeoutCache(Of T)
        ''' <summary>キャッシュの生存時間規定値 15分</summary>
        Private Const DEFAULT_TIMEOUT_CLEAR_MILLIS As Long = 15 * 60 * 1000
        Private _timeoutMillis As Long = DEFAULT_TIMEOUT_CLEAR_MILLIS
        Private ReadOnly rwLock As New ReaderWriterLockSlim
        Private Const KEY As String = "KEY"
        Private cacheVosByKey1 As New TimeoutDictionary(Of String, T)
        ''' <summary>キャッシュの生存時間(ミリ秒)</summary>
        Protected Property TimeoutMillis() As Long
            Get
                Return _timeoutMillis
            End Get
            Set(ByVal value As Long)
                _timeoutMillis = value
                cacheVosByKey1.SetTimeout(value)
            End Set
        End Property

        Protected Sub New()
        End Sub

        ''' <summary>
        ''' キャッシュ内容をクリアする
        ''' </summary>
        ''' <remarks></remarks>
        Protected Sub ClearT()
            rwLock.EnterWriteLock()
            Try
                cacheVosByKey1.Clear()
            Finally
                rwLock.ExitWriteLock()
            End Try
        End Sub

        ''' <summary>
        ''' キャッシュ内容を取得する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Function GetT() As T
            rwLock.EnterUpgradeableReadLock()
            Try
                Dim instance As T
                Dim getsValue As Boolean = cacheVosByKey1.TryGetValue(KEY, instance)
                If Not getsValue OrElse instance Is Nothing Then
                    ' getsValue = true でも instance is Nothing の場合がある
                    rwLock.EnterWriteLock()
                    Try
                        cacheVosByKey1 = New TimeoutDictionary(Of String, T)
                        cacheVosByKey1.SetTimeout(_timeoutMillis)
                        instance = MakeT()
                        cacheVosByKey1.Add(KEY, instance)
                    Finally
                        rwLock.ExitWriteLock()
                    End Try
                End If
                Return instance
            Finally
                rwLock.ExitUpgradeableReadLock()
            End Try
        End Function

        ''' <summary>
        ''' キャッシュ内容を作成する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected MustOverride Function MakeT() As T

    End Class
End Namespace