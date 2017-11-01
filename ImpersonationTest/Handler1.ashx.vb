Imports System.Web
Imports System.Web.Services

Public Class Handler1
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        context.Response.ContentType = "text/plain"
        context.Response.Write(System.Security.Principal.WindowsIdentity.GetCurrent().Name + " Module threadid: " + PaulTestModule.ThreadID(CInt(context.Request.QueryString("Count"))) + " Current thread id: " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString())

    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class