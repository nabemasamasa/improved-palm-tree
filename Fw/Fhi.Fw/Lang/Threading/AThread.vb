Imports System.Threading
Imports System.Globalization

Namespace Lang.Threading
    ''' <summary>
    ''' Javaの（継承可能な）Threadクラス相当
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AThread : Implements Runnable

        Private ReadOnly _Thread As Thread
        Private target As Runnable
        Private _isInterrupted As Boolean

        Public Sub New()
            Me.New(Nothing, Nothing)
        End Sub

        Public Sub New(ByVal target As Runnable)
            Me.New(target, Nothing)
        End Sub

        Public Sub New(ByVal target As Runnable, ByVal name As String)
            _Thread = New Thread(AddressOf Run)
            Me.target = target

            _Thread.IsBackground = True
            If name IsNot Nothing Then
                _Thread.Name = name
            End If
        End Sub

        Public Sub New(ByVal name As String)
            Me.New(Nothing, name)
        End Sub

        ''' <summary>
        ''' スレッド起動時の実処理
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub Run() Implements Runnable.Run
            If target Is Nothing Then
                Throw New InvalidProgramException("Run()メソッドをオーバーライドするか、コンストラクタにRunnableインターフェースが必要")
            End If
            target.Run()
        End Sub

        ''' <summary>
        ''' 現在のスレッドが割り込まれているかどうかを返す（結果として割り込みステータスがクリアされる）
        ''' </summary>
        ''' <returns>現在のスレッドが割り込まれている場合、true</returns>
        ''' <remarks></remarks>
        Public Shared Function Interrupted() As Boolean
            Try
                Thread.Sleep(0)
            Catch ex As ThreadInterruptedException
                Return True
            End Try
            Return False
        End Function

        ''' <summary>
        ''' このスレッドが割り込まれているかどうかを返す（割り込みステータスは影響を受けない）
        ''' </summary>
        ''' <returns>このスレッドが割り込まれている場合、true</returns>
        ''' <remarks></remarks>
        Public Function IsInterrupted() As Boolean
            If Not IsCurrentThreadOwn() Then
                ' 別スレッドのインタラプト状態を再度精査する手段が無いので最新のフラグ返す
                Return _isInterrupted
            End If

            _isInterrupted = Me.IsInterrupted
            If _isInterrupted Then
                Me.Interrupt()
            End If
            Return _isInterrupted
        End Function

        ''' <summary>
        ''' 現在のスレッドがこのスレッドか？を返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Private Function IsCurrentThreadOwn() As Boolean

            Return Thread.CurrentThread Is _Thread
        End Function

        ' --- 以下は Threadクラスの委譲メソッド ---

        ''' <summary>
        ''' Causes the operating system to change the state of the current instance to <see cref="F:System.Threading.ThreadState.Running"/>.
        ''' </summary>
        ''' <exception cref="T:System.Threading.ThreadStateException">The thread has already been started. 
        '''                 </exception><exception cref="T:System.OutOfMemoryException">There is not enough memory available to start this thread. 
        '''                 </exception><filterpriority>1</filterpriority>
        Public Sub Start()
            _Thread.Start()
        End Sub

        ''' <summary>
        ''' Causes the operating system to change the state of the current instance to <see cref="F:System.Threading.ThreadState.Running"/>, and optionally supplies an object containing data to be used by the method the thread executes.
        ''' </summary>
        ''' <param name="parameter">An object that contains data to be used by the method the thread executes.
        '''                 </param><exception cref="T:System.Threading.ThreadStateException">The thread has already been started. 
        '''                 </exception><exception cref="T:System.OutOfMemoryException">There is not enough memory available to start this thread. 
        '''                 </exception><exception cref="T:System.InvalidOperationException">This thread was created using a <see cref="T:System.Threading.ThreadStart"/> delegate instead of a <see cref="T:System.Threading.ParameterizedThreadStart"/> delegate.
        '''                 </exception><filterpriority>1</filterpriority>
        Public Sub Start(ByVal parameter As Object)
            _Thread.Start(parameter)
        End Sub

        '''' <summary>
        '''' Applies a captured <see cref="T:System.Threading.CompressedStack"/> to the current thread.
        '''' </summary>
        '''' <param name="stack">The <see cref="T:System.Threading.CompressedStack"/> object to be applied to the current thread.
        ''''                 </param><filterpriority>2</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="UnmanagedCode"/><IPermission class="System.Security.Permissions.StrongNameIdentityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" PublicKeyBlob="00000000000000000400000000000000"/></PermissionSet>
        '<Obsolete("下位互換用. 使用すべきではない.")> Public Sub SetCompressedStack(ByVal stack As CompressedStack)
        '    _Thread.SetCompressedStack(stack)
        'End Sub

        '''' <summary>
        '''' Returns a <see cref="T:System.Threading.CompressedStack"/> object that can be used to capture the stack for the current thread.
        '''' </summary>
        '''' <returns>
        '''' A <see cref="T:System.Threading.CompressedStack"/> object that can be used to capture the stack for the current thread.
        '''' </returns>
        '''' <filterpriority>2</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="UnmanagedCode"/><IPermission class="System.Security.Permissions.StrongNameIdentityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" PublicKeyBlob="00000000000000000400000000000000"/></PermissionSet>
        '<Obsolete("下位互換用. 使用すべきではない.")> Public Function GetCompressedStack() As CompressedStack
        '    Return _Thread.GetCompressedStack()
        'End Function

        ''' <summary>
        ''' Raises a <see cref="T:System.Threading.ThreadAbortException"/> in the thread on which it is invoked, to begin the process of terminating the thread while also providing exception information about the thread termination. Calling this method usually terminates the thread.
        ''' </summary>
        ''' <param name="stateInfo">An object that contains application-specific information, such as state, which can be used by the thread being aborted. 
        '''                 </param><exception cref="T:System.Security.SecurityException">The caller does not have the required permission. 
        '''                 </exception><exception cref="T:System.Threading.ThreadStateException">The thread that is being aborted is currently suspended.
        '''                 </exception><filterpriority>1</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="ControlThread"/></PermissionSet>
        <Obsolete("安全性が保証できません.")> Public Sub Abort(ByVal stateInfo As Object)
            _Thread.Abort(stateInfo)
        End Sub

        ''' <summary>
        ''' Raises a <see cref="T:System.Threading.ThreadAbortException"/> in the thread on which it is invoked, to begin the process of terminating the thread. Calling this method usually terminates the thread.
        ''' </summary>
        ''' <exception cref="T:System.Security.SecurityException">The caller does not have the required permission. 
        '''                 </exception><exception cref="T:System.Threading.ThreadStateException">The thread that is being aborted is currently suspended.
        '''                 </exception><filterpriority>1</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="ControlThread"/></PermissionSet>
        <Obsolete("安全性が保証できません.")> Public Sub Abort()
            _Thread.Abort()
        End Sub

        '''' <summary>
        '''' Either suspends the thread, or if the thread is already suspended, has no effect.
        '''' </summary>
        '''' <exception cref="T:System.Threading.ThreadStateException">The thread has not been started or is dead. 
        ''''                 </exception><exception cref="T:System.Security.SecurityException">The caller does not have the appropriate <see cref="T:System.Security.Permissions.SecurityPermission"/>. 
        ''''                 </exception><filterpriority>1</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="ControlThread"/></PermissionSet>
        '<Obsolete("下位互換用. 使用すべきではない.")> Public Sub Suspend()
        '    _Thread.Suspend()
        'End Sub

        '''' <summary>
        '''' Resumes a thread that has been suspended.
        '''' </summary>
        '''' <exception cref="T:System.Threading.ThreadStateException">The thread has not been started, is dead, or is not in the suspended state. 
        ''''                 </exception><exception cref="T:System.Security.SecurityException">The caller does not have the appropriate <see cref="T:System.Security.Permissions.SecurityPermission"/>. 
        ''''                 </exception><filterpriority>1</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="ControlThread"/></PermissionSet>
        '<Obsolete("下位互換用. 使用すべきではない.")> <SuppressMessage("Microsoft.Warning", "BA40000")> Public Sub [Resume]()
        '    _Thread.[Resume]()
        'End Sub

        ''' <summary>
        ''' Interrupts a thread that is in the WaitSleepJoin thread state.
        ''' </summary>
        ''' <exception cref="T:System.Security.SecurityException">The caller does not have the appropriate <see cref="T:System.Security.Permissions.SecurityPermission"/>. 
        '''                 </exception><filterpriority>2</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="ControlThread"/></PermissionSet>
        Public Sub Interrupt()
            _Thread.Interrupt()
            _isInterrupted = True
        End Sub

        ''' <summary>
        ''' Blocks the calling thread until a thread terminates, while continuing to perform standard COM and SendMessage pumping.
        ''' </summary>
        ''' <exception cref="T:System.Threading.ThreadStateException">The caller attempted to join a thread that is in the <see cref="F:System.Threading.ThreadState.Unstarted"/> state. 
        '''                 </exception><exception cref="T:System.Threading.ThreadInterruptedException">The thread is interrupted while waiting. 
        '''                 </exception><filterpriority>1</filterpriority>
        Public Sub Join()
            _Thread.Join()
        End Sub

        ''' <summary>
        ''' Blocks the calling thread until a thread terminates or the specified time elapses, while continuing to perform standard COM and SendMessage pumping.
        ''' </summary>
        ''' <returns>
        ''' true if the thread has terminated; false if the thread has not terminated after the amount of time specified by the <paramref name="millisecondsTimeout"/> parameter has elapsed.
        ''' </returns>
        ''' <param name="millisecondsTimeout">The number of milliseconds to wait for the thread to terminate. 
        '''                 </param><exception cref="T:System.ArgumentOutOfRangeException">The value of <paramref name="millisecondsTimeout"/> is negative and is not equal to <see cref="F:System.Threading.Timeout.Infinite"/> in milliseconds. 
        '''                 </exception><exception cref="T:System.Threading.ThreadStateException">The thread has not been started. 
        '''                 </exception><filterpriority>1</filterpriority>
        Public Function Join(ByVal millisecondsTimeout As Integer) As Boolean
            Return _Thread.Join(millisecondsTimeout)
        End Function

        ''' <summary>
        ''' Blocks the calling thread until a thread terminates or the specified time elapses, while continuing to perform standard COM and SendMessage pumping.
        ''' </summary>
        ''' <returns>
        ''' true if the thread terminated; false if the thread has not terminated after the amount of time specified by the <paramref name="timeout"/> parameter has elapsed.
        ''' </returns>
        ''' <param name="timeout">A <see cref="T:System.TimeSpan"/> set to the amount of time to wait for the thread to terminate. 
        '''                 </param><exception cref="T:System.ArgumentOutOfRangeException">The value of <paramref name="timeout"/> is negative and is not equal to <see cref="F:System.Threading.Timeout.Infinite"/> in milliseconds, or is greater than <see cref="F:System.Int32.MaxValue"/> milliseconds. 
        '''                 </exception><exception cref="T:System.Threading.ThreadStateException">The caller attempted to join a thread that is in the <see cref="F:System.Threading.ThreadState.Unstarted"/> state. 
        '''                 </exception><filterpriority>1</filterpriority>
        Public Function Join(ByVal timeout As TimeSpan) As Boolean
            Return _Thread.Join(timeout)
        End Function

        ''' <summary>
        ''' Returns an <see cref="T:System.Threading.ApartmentState"/> value indicating the apartment state.
        ''' </summary>
        ''' <returns>
        ''' One of the <see cref="T:System.Threading.ApartmentState"/> values indicating the apartment state of the managed thread. The default is <see cref="F:System.Threading.ApartmentState.Unknown"/>.
        ''' </returns>
        ''' <filterpriority>1</filterpriority>
        Public Function GetApartmentState() As ApartmentState
            Return _Thread.GetApartmentState()
        End Function

        ''' <summary>
        ''' Sets the apartment state of a thread before it is started.
        ''' </summary>
        ''' <returns>
        ''' true if the apartment state is set; otherwise, false.
        ''' </returns>
        ''' <param name="state">The new apartment state.
        '''                 </param><exception cref="T:System.ArgumentException"><paramref name="state"/> is not a valid apartment state.
        '''                 </exception><exception cref="T:System.Threading.ThreadStateException">The thread has already been started.
        '''                 </exception><filterpriority>1</filterpriority>
        Public Function TrySetApartmentState(ByVal state As ApartmentState) As Boolean
            Return _Thread.TrySetApartmentState(state)
        End Function

        ''' <summary>
        ''' Sets the apartment state of a thread before it is started.
        ''' </summary>
        ''' <param name="state">The new apartment state.
        '''                 </param><exception cref="T:System.ArgumentException"><paramref name="state"/> is not a valid apartment state.
        '''                 </exception><exception cref="T:System.Threading.ThreadStateException">The thread has already been started.
        '''                 </exception><exception cref="T:System.InvalidOperationException">The apartment state has already been initialized.
        '''                 </exception><filterpriority>1</filterpriority>
        Public Sub SetApartmentState(ByVal state As ApartmentState)
            _Thread.SetApartmentState(state)
        End Sub

        ''' <summary>
        ''' Gets a unique identifier for the current managed thread.
        ''' </summary>
        ''' <returns>
        ''' An integer that represents a unique identifier for this managed thread.
        ''' </returns>
        ''' <filterpriority>1</filterpriority>
        Public ReadOnly Property ManagedThreadId() As Integer
            Get
                Return _Thread.ManagedThreadId
            End Get
        End Property

        ''' <summary>
        ''' Gets an <see cref="T:System.Threading.ExecutionContext"/> object that contains information about the various contexts of the current thread. 
        ''' </summary>
        ''' <returns>
        ''' An <see cref="T:System.Threading.ExecutionContext"/> object that consolidates context information for the current thread.
        ''' </returns>
        ''' <filterpriority>2</filterpriority>
        Public ReadOnly Property ExecutionContext() As ExecutionContext
            Get
                Return _Thread.ExecutionContext
            End Get
        End Property

        ''' <summary>
        ''' Gets or sets a value indicating the scheduling priority of a thread.
        ''' </summary>
        ''' <returns>
        ''' One of the <see cref="T:System.Threading.ThreadPriority"/> values. The default value is Normal.
        ''' </returns>
        ''' <exception cref="T:System.Threading.ThreadStateException">The thread has reached a final state, such as <see cref="F:System.Threading.ThreadState.Aborted"/>. 
        '''                 </exception><exception cref="T:System.ArgumentException">The value specified for a set operation is not a valid ThreadPriority value. 
        '''                 </exception><filterpriority>1</filterpriority>
        Public Property Priority() As ThreadPriority
            Get
                Return _Thread.Priority
            End Get
            Set(ByVal value As ThreadPriority)
                _Thread.Priority = value
            End Set
        End Property

        ''' <summary>
        ''' Gets a value indicating the execution status of the current thread.
        ''' </summary>
        ''' <returns>
        ''' true if this thread has been started and has not terminated normally or aborted; otherwise, false.
        ''' </returns>
        ''' <filterpriority>1</filterpriority>
        Public ReadOnly Property IsAlive() As Boolean
            Get
                Return _Thread.IsAlive
            End Get
        End Property

        ''' <summary>
        ''' Gets a value indicating whether or not a thread belongs to the managed thread pool.
        ''' </summary>
        ''' <returns>
        ''' true if this thread belongs to the managed thread pool; otherwise, false.
        ''' </returns>
        ''' <filterpriority>2</filterpriority>
        Public ReadOnly Property IsThreadPoolThread() As Boolean
            Get
                Return _Thread.IsThreadPoolThread
            End Get
        End Property

        ''' <summary>
        ''' Gets or sets a value indicating whether or not a thread is a background thread.
        ''' </summary>
        ''' <returns>
        ''' true if this thread is or is to become a background thread; otherwise, false.
        ''' </returns>
        ''' <exception cref="T:System.Threading.ThreadStateException">The thread is dead. 
        '''                 </exception><filterpriority>1</filterpriority>
        Public Property IsBackground() As Boolean
            Get
                Return _Thread.IsBackground
            End Get
            Set(ByVal value As Boolean)
                _Thread.IsBackground = value
            End Set
        End Property

        ''' <summary>
        ''' Gets a value containing the states of the current thread.
        ''' </summary>
        ''' <returns>
        ''' One of the <see cref="T:System.Threading.ThreadState"/> values indicating the state of the current thread. The initial value is Unstarted.
        ''' </returns>
        ''' <filterpriority>2</filterpriority>
        Public ReadOnly Property ThreadState() As ThreadState
            Get
                Return _Thread.ThreadState
            End Get
        End Property

        '''' <summary>
        '''' Gets or sets the apartment state of this thread.
        '''' </summary>
        '''' <returns>
        '''' One of the <see cref="T:System.Threading.ApartmentState"/> values. The initial value is Unknown.
        '''' </returns>
        '''' <exception cref="T:System.ArgumentException">An attempt is made to set this property to a state that is not a valid apartment state (a state other than single-threaded apartment (STA) or multithreaded apartment (MTA)). 
        ''''                 </exception><filterpriority>2</filterpriority>
        '<Obsolete("下位互換用. 使用すべきではない.")> Public Property ApartmentState() As ApartmentState
        '    Get
        '        Return _Thread.ApartmentState
        '    End Get
        '    Set(ByVal value As ApartmentState)
        '        _Thread.ApartmentState = value
        '    End Set
        'End Property

        ''' <summary>
        ''' Gets or sets the current culture used by the Resource Manager to look up culture-specific resources at run time.
        ''' </summary>
        ''' <returns>
        ''' A <see cref="T:System.Globalization.CultureInfo"/> representing the current culture.
        ''' </returns>
        ''' <exception cref="T:System.ArgumentNullException">The property is set to null. 
        '''                 </exception><exception cref="T:System.ArgumentException">The property is set to a culture name that cannot be used to locate a resource file. Resource filenames must include only letters, numbers, hyphens or underscores.
        '''                 </exception><filterpriority>2</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="UnmanagedCode"/></PermissionSet>
        Public Property CurrentUICulture() As CultureInfo
            Get
                Return _Thread.CurrentUICulture
            End Get
            Set(ByVal value As CultureInfo)
                _Thread.CurrentUICulture = value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the culture for the current thread.
        ''' </summary>
        ''' <returns>
        ''' A <see cref="T:System.Globalization.CultureInfo"/> representing the culture for the current thread.
        ''' </returns>
        ''' <exception cref="T:System.NotSupportedException">The property is set to a neutral culture. Neutral cultures cannot be used in formatting and parsing and therefore cannot be set as the thread's current culture.
        '''                 </exception><exception cref="T:System.ArgumentNullException">The property is set to null.
        '''                 </exception><filterpriority>2</filterpriority><PermissionSet><IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="ControlThread"/></PermissionSet>
        Public Property CurrentCulture() As CultureInfo
            Get
                Return _Thread.CurrentCulture
            End Get
            Set(ByVal value As CultureInfo)
                _Thread.CurrentCulture = value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the name of the thread.
        ''' </summary>
        ''' <returns>
        ''' A string containing the name of the thread, or null if no name was set.
        ''' </returns>
        ''' <exception cref="T:System.InvalidOperationException">A set operation was requested, and the Name property has already been set. 
        '''                 </exception><filterpriority>1</filterpriority>
        Public Property Name() As String
            Get
                Return _Thread.Name
            End Get
            Set(ByVal value As String)
                _Thread.Name = value
            End Set
        End Property

    End Class
End Namespace