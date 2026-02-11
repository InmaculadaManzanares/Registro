Imports System
Imports System.Data
Imports System.Text

Partial Public Class ResetPasswordRequest
    Inherits System.Web.UI.Page

    Private Function Q(name As String) As String
        Dim vq = Request.QueryString(name)
        If vq Is Nothing Then Return ""
        Return vq.Trim()
    End Function

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If IsPostBack Then Return

        ' Si viene desde la app: ResetPasswordRequest.aspx?u=email
        Dim loginUsuario As String = Q("u")
        If loginUsuario <> "" Then
            txtusuario.Text = loginUsuario

            ' opcional: auto-enviar
            Try
                EnviarResetSiExiste(loginUsuario)
                ShowOk("Si el correo existe, recibirá un email con instrucciones para restablecer la contraseña.")
            Catch
                ' no revelar errores
                ShowOk("Si el correo existe, recibirá un email con instrucciones para restablecer la contraseña.")
            End Try
        End If
    End Sub

    Protected Sub btnEnviar_Click(sender As Object, e As EventArgs) Handles btnEnviar.Click
        Dim loginUsuario As String = If(txtusuario.Text, "").Trim()

        If loginUsuario = "" Then
            ShowError("Debe indicar un correo.")
            Exit Sub
        End If

        Try
            EnviarResetSiExiste(loginUsuario)
            ShowOk("Si el correo existe, recibirá un email con instrucciones para restablecer la contraseña.")
        Catch
            ' por seguridad: mensaje genérico
            ShowOk("Si el correo existe, recibirá un email con instrucciones para restablecer la contraseña.")
        End Try
    End Sub

    Private Sub EnviarResetSiExiste(loginUsuario As String)
        Dim emailEnvio As String = ""

        ' 1) Buscar EmailEnvio del usuario
        Dim con As New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Try
            Dim dt As DataTable = con.SelectSQL("SELECT TOP 1 EmailEnvio FROM Usuarios WHERE Email=" & v(loginUsuario))
            If dt.Rows.Count > 0 AndAlso dt.Rows(0)("EmailEnvio") IsNot DBNull.Value Then
                emailEnvio = dt.Rows(0)("EmailEnvio").ToString().Trim()
            End If
        Finally
            If con IsNot Nothing Then con.CerrarConexion()
        End Try

        If emailEnvio = "" Then emailEnvio = loginUsuario

        ' 2) Lógica EXACTA de envío (tu misma función)
        EnviarResetPassword(loginUsuario, emailEnvio)
    End Sub

    ' === COPIA de tu función (misma lógica) ===
    Private Function EnviarResetPassword(loginUsuario As String, emailEnvio As String) As Boolean
        Dim con As New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))

        Try
            If String.IsNullOrWhiteSpace(loginUsuario) Then Return False
            If String.IsNullOrWhiteSpace(emailEnvio) Then Return False

            ' 1) Generar token y hash
            Dim token As String = ResetToken.GenerateTokenBase64Url(32) ' 256-bit
            Dim tokenHash As Byte() = ResetToken.Sha256Bytes(token)

            ' 2) Guardar token en BD (15 min)
            Dim expires As Date = DateAdd(DateInterval.Minute, 15, Now)

            Dim tokenHex As String = "0x" & BitConverter.ToString(tokenHash).Replace("-", "")

            Dim sql As String =
                "INSERT INTO PasswordResetTokens (Email, TokenHash, ExpiresAt, Used) VALUES (" &
                v(loginUsuario) & "," & tokenHex & "," & Util.fechasql(expires, tipofecha.yyyyddMMmmss) & ",0)"

            con.Ejecuta(sql)

            ' 3) Enlace (u = loginUsuario)
            Dim baseUrl As String = System.Configuration.ConfigurationManager.AppSettings("url").ToString()
            Dim link As String = baseUrl.TrimEnd("/"c) & "/ResetPassword.aspx?u=" &
                                 Server.UrlEncode(loginUsuario) & "&t=" & Server.UrlEncode(token)

            ' 4) Email (se envía al email REAL)
            Dim correo As New EnviodeEmail
            Dim asunto As String = "Restablecer contraseña"

            Dim urlTemplate As String = Server.MapPath("~/Plantillas/RecuperarContraseña.html")
            Dim urlimagen As String = System.Configuration.ConfigurationManager.AppSettings("url").ToString()

            Dim Template As New StringBuilder
            Template.Append(GetHTMLFromAddress(urlTemplate))
            Template.Replace("$USER$", loginUsuario)
            Template.Replace("$CONTRASE$", link)
            Template.Replace("$URLIMAGEN$", urlimagen)

            Dim fromEmail As String = VariablesGlobales.Email_WebCampus(Application)
            Return correo.EnviarEmail_html(fromEmail, emailEnvio, asunto, Template.ToString(), "", Application)

        Catch
            Return False
        Finally
            If con IsNot Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try
    End Function

    ' === Helpers UI ===
    Private Sub ShowError(msg As String)
        mensaje.Text = msg
        Dim script As String = "document.getElementById('alert').style.display='block';"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "alerta", script, True)
    End Sub
    Public Shared Function GetHTMLFromAddress(ByVal Address As String) As String
        Dim ASCII As New System.Text.UTF8Encoding
        Dim netWeb As New System.Net.WebClient
        Dim lsWeb As String
        Dim laWeb As Byte()

        Try
            laWeb = netWeb.DownloadData(Address)

            lsWeb = ASCII.GetString(laWeb)
        Catch ex As Exception
            Throw New Exception(ex.Message.ToString + ex.ToString)
        End Try
        Return lsWeb
    End Function
    Private Sub ShowOk(msg As String)
        mensajeOk.Text = msg

        panel1.Visible = False
        panelVolver.Visible = True

        Dim script As String = "document.getElementById('ok').style.display='block';"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "okmsg", script, True)
    End Sub

End Class
