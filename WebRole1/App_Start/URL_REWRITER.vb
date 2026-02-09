Imports Microsoft.VisualBasic
Imports System.Web

Public Class URL_REWRITER
    Implements IHttpModule


    Public Sub Dispose() Implements System.Web.IHttpModule.Dispose

    End Sub

    Public Sub Init(ByVal context As System.Web.HttpApplication) Implements System.Web.IHttpModule.Init
        AddHandler context.AuthorizeRequest, AddressOf Realiza
    End Sub

    Private Sub Realiza(ByVal sender As Object, ByVal e As EventArgs)
        Dim app As HttpApplication
        Dim cadena As String
        Dim partes() As String
        Dim m As Match

        app = DirectCast(sender, HttpApplication)

        cadena = app.Request.Path.ToLower
        If Regex.IsMatch(cadena, "/dbv.+\.pdf$") Then

            m = Regex.Match(cadena, "dbv[0-9]+-[0-9]+-[0-9]+-[01]+-[01]\.pdf$")

            partes = m.Value.Substring(3, m.Value.Length - 7).Split("-")
            Dim id, bd, serie, temp, borr As Long

            id = partes(0)
            bd = partes(1)
            serie = partes(2)
            temp = partes(3)
            borr = partes(4)
            app.Context.RewritePath("~/generapdf.ashx?ID=" & id & "&Extension=PDF" & "&BD=" & bd & "&Serie=" & serie & "&temp=" & temp & "&borr=" & borr)

        End If
    End Sub
End Class
