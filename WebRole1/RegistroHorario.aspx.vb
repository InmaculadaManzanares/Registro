Imports System
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class RegistroHorario
    Inherits System.Web.UI.Page
    Public USUARIO As String
    Private Const PASSWORD_EXPIRE_DAYS As Integer = 90
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            panel1.Visible = True
            panel2.Visible = False
            txtcontraseña.Attributes.Add("onkeypress", "javascript:if(event.keyCode==13){return true;}")
            ImageButton1.Attributes.Add("onkeypress", "javascript:if(event.keyCode==13){return true;}")
            Dim con As ConexionSQL = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
            Dim colorprimario As String = ""
            Dim colorsecundario As String = ""
            Try 'lee la configuracion para cambiar los colores
                colorprimario = con.ejecuta1v_string("select Colorprimario from config")
                colorsecundario = con.ejecuta1v_string("select colorsecundario from config")
            Catch ex As Exception
            End Try
            If colorprimario = "" Then colorprimario = "#252149" 'AZUL OSCURO
            If colorsecundario = "" Then colorsecundario = "#27b9c6" 'TURQUESA
            Session("colorprimario") = colorprimario
            Session("colorsecundario") = colorsecundario
            list.Visible = False
            panelReset.Visible = False
        End If
        ' Inyectar el script en la página
        Dim script1 As String = Util.CambiarColorControlesScript(Session("colorprimario"), Session("colorsecundario"))
        ClientScript.RegisterStartupScript(Me.GetType(), "ReplaceColorsScript1", script1)
    End Sub
    Private Function IsPasswordExpired(lastChangedUtcObj As Object) As Boolean
        If lastChangedUtcObj Is Nothing OrElse lastChangedUtcObj Is DBNull.Value Then
            ' Si no hay fecha, trátalo como caducado para forzar reset (o decide tú)
            Return True
        End If

        Dim lastChanged As DateTime = Convert.ToDateTime(lastChangedUtcObj)
        Dim expiresAt As DateTime = lastChanged.AddDays(PASSWORD_EXPIRE_DAYS)

        ' Si guardas UTC usa UtcNow; si guardas local usa Now
        Return DateTime.UtcNow > expiresAt
    End Function
    Protected Sub lnkReset_Click(sender As Object, e As EventArgs) Handles lnkReset.Click
        ' Muestra panel reset y oculta login
        panelReset.Visible = True
        panel1.Visible = False

        ' Puedes rellenar el email por si el usuario ya lo escribió
        txtResetEmail.Text = txtusuario.Text

        ' Oculta los campos de login para que desaparezcan como pides
        txtusuario.Visible = False
        txtcontraseña.Visible = False
        ImageButton1.Visible = False
        lnkReset.Visible = False
    End Sub

    Protected Sub lnkVolverAcceso_Click(sender As Object, e As EventArgs) Handles lnkVolverAcceso.Click
        ' Vuelve al login normal
        'panelReset.Visible = False

        'txtusuario.Visible = True
        'txtcontraseña.Visible = True
        'ImageButton1.Visible = True
        'txtResetEmail.Visible = True
        'btnResetEnviar.Visible = True
        'mensaje.Text = ""
        Response.Redirect(Request.RawUrl)
    End Sub

    Protected Sub btnResetEnviar_Click(sender As Object, e As EventArgs) Handles btnResetEnviar.Click
        Dim loginUsuario As String = txtResetEmail.Text.Trim()

        If String.IsNullOrWhiteSpace(loginUsuario) Then
            mensaje.Text = "Debe indicar un correo electrónico."
            MostrarAlerta()
            Exit Sub
        End If

        ' 1) Buscar email real de envío (EmailEnvio), si existe
        Dim con As ConexionSQL = Nothing
        Try
            con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))

            Dim emailEnvio As String = con.ejecuta1v_string("SELECT ISNULL(NULLIF(EmailEnvio,''), Email) FROM Usuarios WHERE Email=" & v(loginUsuario))
            If String.IsNullOrWhiteSpace(emailEnvio) Then emailEnvio = loginUsuario

            ' 2) Si existe usuario => enviamos reset
            ' (Si NO existe, respondemos igual por seguridad)
            Dim existe As String = con.ejecuta1v_string("SELECT Email FROM Usuarios WHERE Email=" & v(loginUsuario))

            If Not String.IsNullOrWhiteSpace(existe) Then
                EnviarResetPassword(loginUsuario, emailEnvio)
            End If

            ' 3) Mensaje genérico (no revela si existe)
            mensaje.Text = "Si el correo existe, recibirá un email con instrucciones para restablecer la contraseña."
            MostrarAlerta()

            ' 4) Oculta todo y deja solo volver al acceso

            btnResetEnviar.Visible = False
            divResetEmail.Visible = False
        Catch
            mensaje.Text = "No se pudo solicitar el restablecimiento. Inténtelo de nuevo."
            MostrarAlerta()
        Finally
            If con IsNot Nothing Then con.CerrarConexion()
        End Try
    End Sub

    Protected Function cod_usuario() As String
        Dim codigo As String
        Dim posicion As Integer
        Dim login As String
        login = Request.ServerVariables("LOGON_USER").ToString()
        posicion = InStr(login, "\")
        If posicion <> 0 Then
            codigo = Mid(login, posicion + 1, Len(login) - posicion)
        Else
            codigo = login
        End If

        Return codigo
    End Function


    ''' <summary>
    ''' Procedimiento para mostrar el mensaje de alerta.
    ''' </summary>
    Private Sub MostrarAlerta()
        Dim script As String = "MostrarAlerta();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "MostrarAlerta()", script, True)
    End Sub
    Protected Sub ImageButton1_Click(sender As Object, e As EventArgs) Handles ImageButton1.Click
        Dim bien As Boolean = False
        Dim con As ConexionSQL = Nothing
        Dim SEN As New SentenciaSQL

        Try
            con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))

            ' 1) Buscar usuario por login (email = usuario/login)
            With SEN
                .Limpia()
                .sql_select = "email, password, PasswordLastChangedUtc, EmailEnvio"
                .sql_from = "Usuarios"
                .add_condicion("email", txtusuario.Text)
            End With

            Dim dt As DataTable = con.SelectSQL(SEN.texto_sql)

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                Me.mensaje.Text = "El usuario no es correcto o no está registrado"
                MostrarAlerta()
                Exit Sub
            End If

            Dim login As String = dt.Rows(0)("email").ToString()
            Dim storedHash As String = dt.Rows(0)("password").ToString()

            If String.IsNullOrWhiteSpace(storedHash) Then
                Me.mensaje.Text = "El usuario no es correcto o no está registrado"
                MostrarAlerta()
                Exit Sub
            End If

            ' Email real de envío (si no está informado, fallback al login)
            Dim emailEnvio As String = ""
            If dt.Columns.Contains("EmailEnvio") AndAlso dt.Rows(0)("EmailEnvio") IsNot DBNull.Value Then
                emailEnvio = dt.Rows(0)("EmailEnvio").ToString().Trim()
            End If
            If String.IsNullOrWhiteSpace(emailEnvio) Then
                emailEnvio = login
            End If

            ' Caducidad
            If IsPasswordExpired(dt.Rows(0)("PasswordLastChangedUtc")) Then

                ' Envía correo de reset y bloquea acceso
                If EnviarResetPassword(login, emailEnvio) Then
                    Me.mensaje.Text = "Tu contraseña ha caducado. Te hemos enviado un correo con el enlace para restablecerla."
                Else
                    Me.mensaje.Text = "Tu contraseña ha caducado, pero no se pudo enviar el correo. Contacta con soporte."
                End If

                MostrarAlerta()
                panel1.Visible = False
                panelReset.Visible = True
                divResetEmail.Visible = False
                btnResetEnviar.Visible = False
                lnkReset.Visible = False
                panel2.Visible = False
                list.Visible = False
                Session("usuario") = Nothing
                Exit Sub
            End If




            ' 2) Verificar contraseña (texto plano) contra hash almacenado
            If PasswordHasher.VerifyPassword(txtcontraseña.Text, storedHash) Then
                bien = True
                Session("usuario") = login
            Else
                Me.mensaje.Text = "El usuario no es correcto o no está registrado"
                MostrarAlerta()
                Exit Sub
            End If

            ' 3) Si login OK, mostrar panel2 como ya hacías
            If bien Then
                list.Visible = True
                panel1.Visible = False
                lnkReset.Visible = False
                ' Nombre del usuario
                With SEN
                    .Limpia()
                    .sql_select = "nombre"
                    .sql_from = "Usuarios"
                    .add_condicion("email", txtusuario.Text)
                End With
                txtnombre.Text = con.ejecuta1v_string(SEN.texto_sql)

                ' Entrada/Salida del día
                Dim inout As String = con.ejecuta1v_string(
                "select top 1 ISNULL(inout, '') as inout from RegistroHorario " &
                "where email=" & v(txtusuario.Text) & " and FORMAT(fecha, 'yyyyMMdd') = FORMAT(getdate(), 'yyyyMMdd') " &
                "order by fecha desc"
            )

                If inout = "Entrada" Then
                    EntradaSalida.Text = "Salida"
                Else
                    EntradaSalida.Text = "Entrada"
                End If

                panel2.Visible = True
            End If

        Catch ex As Exception
            Me.mensaje.Text = "Imposible conectar con la BD"
            MostrarAlerta()
        Finally
            If con IsNot Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try
    End Sub
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

    Private Sub EntradaSalida_Click(sender As Object, e As EventArgs) Handles EntradaSalida.Click
        If EntradaSalida.Text = "Salir" Then
            txtnombre.Text = ""
            txtubicacion.Visible = False
            txtlugar.Text = "Oficina"
            txtlugar.Visible = True
            EntradaSalida.Text = "Entrada"
            panel2.Visible = False
            panel1.Visible = True
        Else
            Dim con As ConexionSQL
            Dim upd = New Update_sql
            con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
            Try
                con.Ejecuta("insert into RegistroHorario(email, fecha, inout, oficina, obs, ubicacion) values (" &
                          v(Session("usuario")) & ", getdate(), " & v(EntradaSalida.Text) & ", " & v(txtlugar.Text) & "," &
                          v(txtlugar.Text) & ",'Web')")
                txtnombre.Text = DateTime.Now.ToLongDateString() & " " & DateTime.Now.ToLongTimeString

                con.Ejecuta("update Usuarios set  notificado =0 where email = " & v(Session("usuario")))
                txtubicacion.Visible = True
                EntradaSalida.Text = "Salir"
                lblubicacion.Visible = False
                txtlugar.Visible = False
            Catch ex As Exception
                Me.mensaje.Text = "No se ha podido realizar la " & EntradaSalida.Text & " intentelo de nuevo"
                MostrarAlerta()
            Finally
                If Not con Is Nothing Then
                    con.CerrarConexion()
                    con = Nothing
                End If
            End Try
        End If
    End Sub
    Public Function encripta_valor(ByVal claveinstalacion As String, ByVal claveespaceland As String, ByVal texto As String) As String
        Dim gen As New Crypto(claveinstalacion & "X" & claveespaceland)
        Dim ta As String
        ta = texto
        Return gen.Encriptar(ta)
    End Function
    Private Sub list_Click(sender As Object, e As EventArgs) Handles list.Click
        Dim login As String = CStr(Session("usuario"))

        Dim masterKeyB64 As String = System.Configuration.ConfigurationManager.AppSettings("UrlCryptoMasterKey")
        Dim token As String = CryptoAes.EncryptToBase64Url(login, masterKeyB64)

        Response.Redirect("~/ListRegistroEmail?Email=" & token)
    End Sub



End Class