Public Class Page
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Context.Response.Clear()
        Context.Response.Write(System.Security.Principal.WindowsIdentity.GetCurrent().Name + " Module threadid: " + PaulTestModule.ThreadID(CInt(Request.QueryString("Count"))).ToString() + " Current thread id: " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString())
        Context.Response.End()

    End Sub

End Class