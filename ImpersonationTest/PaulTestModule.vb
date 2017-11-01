Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports System.Web

Public Class PaulTestModule
    Implements IHttpModule

    Private _objContext As System.Security.Principal.WindowsImpersonationContext
    Public Shared Property ThreadID As New Dictionary(Of Integer, Integer)
    Private Shared _lockObject As New Object()

    Public Sub Init(context As HttpApplication) Implements IHttpModule.Init
        AddHandler System.Web.HttpContext.Current.ApplicationInstance.PreRequestHandlerExecute, AddressOf PreRequestHandlerExecute
        AddHandler System.Web.HttpContext.Current.ApplicationInstance.PostRequestHandlerExecute, AddressOf PostRequestHandlerExecute
    End Sub

    Private Sub PreRequestHandlerExecute(ByVal sender As Object, args As EventArgs)
        SyncLock ThreadID
            _objContext = Impersonate("ImpersonatedUser", "", "****", False)
            ThreadID(CInt(System.Web.HttpContext.Current.Request.QueryString("Count"))) = System.Threading.Thread.CurrentThread.ManagedThreadId
        End SyncLock
    End Sub

    Private Sub PostRequestHandlerExecute(ByVal sender As Object, args As EventArgs)
        If Not _objContext Is Nothing Then
            _objContext.Undo()
        End If
    End Sub

    Public Sub Dispose() Implements IHttpModule.Dispose
    End Sub

    ''' <summary>Impersonates the current request as the specified user</summary>
    ''' <param name="strUserName">The name of the user to impersonate</param>
    ''' <param name="strDomainName">The domain name of the user to impersonate</param>
    ''' <param name="strPassword">The password of the user to impersonate</param>
    ''' <param name="blnForUNC">Indicates if you want to impersonate for UNC (so it works for sources outside your domain)</param>
    ''' <returns>The impersonation context if the impersonation was successful, else Nothing</returns>
    ''' <remarks></remarks>
    ''' <modification date="01-11-2006" developer="b. molsbeck">Created</modification>
    ''' <modification date="05-02-2007" developer="s. v. loon">Moved function to SecurityClass</modification>
    Friend Shared Function Impersonate(ByVal strUserName As String,
                                       ByVal strDomainName As String,
                                       ByVal strPassword As String,
                                       ByVal blnForUNC As Boolean) As WindowsImpersonationContext

        'create variable to store token
        Dim objToken As New SafeHandledToken()
        Dim objDuplicatedToken As New SafeHandledToken()
        Dim objIdentity As WindowsIdentity = Nothing

        Try

            Dim intLogonType As Integer

            'setup right logon type
            If blnForUNC Then
                intLogonType = LogonType.LOGON32_LOGON_NEW_CREDENTIALS
            Else
                intLogonType = LogonType.LOGON32_LOGON_INTERACTIVE
            End If

            Dim intLogonProvider As Integer = LogonProvider.LOGON32_PROVIDER_DEFAULT

            'try to logon on user to domain
            Dim blnSuccess As Boolean = LogonUser(strUserName,
                                                           strDomainName,
                                                           strPassword,
                                                           intLogonType,
                                                           intLogonProvider,
                                                           objToken)

            'logon unsuccessful?
            If Not blnSuccess Then
                Return Nothing
            End If

            DuplicateToken(objToken.DangerousGetHandle(), ImpersonationLevel.SecurityImpersonation, objDuplicatedToken)

            'impersonate user
            Using objDuplicatedToken
                objIdentity = objDuplicatedToken.GetWindowsIdentity()
            End Using

            Using objIdentity
                Return objIdentity.Impersonate()
            End Using

        Finally
            If Not objIdentity Is Nothing Then
                objIdentity.Dispose()
            End If
        End Try

    End Function

    Friend Enum ImpersonationLevel

        SecurityAnonymous = 0
        SecurityIdentification
        SecurityImpersonation
        SecurityDelegation
    End Enum

    ''' <summary>
    ''' Safe wrapper for a handle to a user token.
    ''' </summary>
    Public NotInheritable Class SafeHandledToken
        Inherits Microsoft.Win32.SafeHandles.SafeHandleZeroOrMinusOneIsInvalid

        ''' <summary>
        ''' Creates a New SafeTokenHandle. This constructor should only be called by P/Invoke.
        ''' </summary>
        Public Sub New()
            MyBase.New(True)
        End Sub

        ''' <summary>
        ''' Creates a New SafeTokenHandle to wrap the specified user token.
        ''' </summary>
        ''' <param name="tokenHandle">The user token to wrap.</param>
        Public Sub New(tokenHandle As IntPtr, ownsHandle As Boolean)
            MyBase.New(ownsHandle)
            MyBase.SetHandle(tokenHandle)
        End Sub

        Public Function GetWindowsIdentity() As WindowsIdentity

            If MyBase.IsClosed Then
                Throw New ObjectDisposedException("The user token has been released.")
            End If

            If MyBase.IsInvalid Then
                Throw New InvalidOperationException("The user token is invalid.")
            End If

            Return New WindowsIdentity(MyBase.handle)

        End Function

        ''' <summary>
        ''' </summary>
        ''' <returns>
        ''' <c>TRUE</c> if the function succeeds, <c>FALSE otherwise</c>.
        '''
        ''' <para>
        ''' To get extended error information, call <see cref="System.Runtime.InteropServices.Marshal.GetLastWin32Error"/>.
        ''' </para>
        ''' </returns>
        <System.Runtime.ConstrainedExecution.ReliabilityContract(System.Runtime.ConstrainedExecution.Consistency.WillNotCorruptState, System.Runtime.ConstrainedExecution.Cer.MayFail)>
        Protected Overrides Function ReleaseHandle() As Boolean
            Return CloseHandle(MyBase.handle)
        End Function

    End Class

    ''' <summary>Logs on a user to a domain with a password using a logontype and logon provider</summary>
    ''' <param name="lpszUsername">The name of the user to log on</param>
    ''' <param name="lpszDomain">The domain of the user to log on</param>
    ''' <param name="lpszPassword">The password of the use to log on</param>
    ''' <param name="dwLogonType">The logon type</param>
    ''' <param name="dwLogonProvider">The logon provider</param>
    ''' <param name="phToken">(output) the logon-token</param>
    ''' <returns>True if the logon succeeded, else False</returns>
    ''' <remarks></remarks>
    ''' <modification date="01-11-2006 " developer="b. molsbeck">Created</modification>
    ''' <modification date="05-02-2007" developer="s. v. loon">Moved function to SecurityClass</modification>
    <DllImport("advapi32.dll")>
    Private Shared Function LogonUser(ByVal lpszUsername As String,
                                      ByVal lpszDomain As String,
                                      ByVal lpszPassword As String,
                                      ByVal dwLogonType As Integer,
                                      ByVal dwLogonProvider As Integer,
                                      ByRef phToken As SafeHandledToken) As Boolean
    End Function

    ''' <summary>Logs on a user to a domain with a password using a logontype and logon provider</summary>
    ''' <returns>True if the logon succeeded, else False</returns>
    ''' <remarks></remarks>
    ''' <modification date="01-11-2006 " developer="b. molsbeck">Created</modification>
    ''' <modification date="05-02-2007" developer="s. v. loon">Moved function to SecurityClass</modification>
    <DllImport("advapi32.dll")>
    Private Shared Function DuplicateToken(ByVal existingTokenHandle As IntPtr,
                                           ByVal impersionationlevel As Integer,
                                           ByRef newTokenHandle As SafeHandledToken) As Boolean
    End Function

    <DllImport("kernel32.dll")>
    Private Shared Function CloseHandle(ByVal handle As IntPtr) As Boolean
    End Function

    Private Enum LogonType As Integer

        'This logon type is intended for users who will be interactively using the computer, such as a user being logged on 
        'by a terminal server, remote shell, or similar process.
        'This logon type has the additional expense of caching logon information for disconnected operations; 
        'therefore, it is inappropriate for some client/server applications,
        'such as a mail server.
        LOGON32_LOGON_INTERACTIVE = 2

        'This logon type is intended for high performance servers to authenticate plaintext passwords.
        'The LogonUser function does not cache credentials for this logon type.
        LOGON32_LOGON_NETWORK = 3

        'This logon type is intended for batch servers, where processes may be executing on behalf of a user without 
        'their direct intervention. This type is also for higher performance servers that process many plaintext
        'authentication attempts at a time, such as mail or Web servers. 
        'The LogonUser function does not cache credentials for this logon type.
        LOGON32_LOGON_BATCH = 4

        'Indicates a service-type logon. The account provided must have the service privilege enabled. 
        LOGON32_LOGON_SERVICE = 5

        'This logon type is for GINA DLLs that log on users who will be interactively using the computer. 
        'This logon type can generate a unique audit record that shows when the workstation was unlocked. 
        LOGON32_LOGON_UNLOCK = 7

        'This logon type preserves the name and password in the authentication package, which allows the server to make 
        'connections to other network servers while impersonating the client. A server can accept plaintext credentials 
        'from a client, call LogonUser, verify that the user can access the system across the network, and still 
        'communicate with other servers.
        'NOTE: Windows NT:  This value is not supported. 
        LOGON32_LOGON_NETWORK_CLEARTEXT = 8

        'This logon type allows the caller to clone its current token and specify new credentials for outbound connections.
        'The new logon session has the same local identifier but uses different credentials for other network connections. 
        'NOTE: This logon type is supported only by the LOGON32_PROVIDER_WINNT50 logon provider.
        'NOTE: Windows NT:  This value is not supported. 
        LOGON32_LOGON_NEW_CREDENTIALS = 9

    End Enum

    Private Enum LogonProvider As Integer

        'Use the standard logon provider for the system. 
        'The default security provider is negotiate, unless you pass NULL for the domain name and the user name 
        'is not in UPN format. In this case, the default provider is NTLM. 
        'NOTE: Windows 2000/NT:   The default security provider is NTLM.
        LOGON32_PROVIDER_DEFAULT = 0

    End Enum

End Class
