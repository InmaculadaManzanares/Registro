Imports TSHAK
Partial Class InicioGestion
    Inherits System.Web.UI.Page
    Public USUARIO As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then

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
            ' Inyectar el script en la página
        End If
        Dim script1 As String = Util.CambiarColorControlesScript(Session("colorprimario"), Session("colorsecundario"))
        ClientScript.RegisterStartupScript(Me.GetType(), "ReplaceColorsScript1", script1)



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
        Dim conexion_bd As Boolean = True

        Try
            con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))

            ' 1) Buscar usuario admin por login (email = usuario/login)
            ' Traemos el password almacenado (hash) y el email (login)
            With SEN
                .Limpia()
                .sql_select = "email, password"
                .sql_from = "Usuarios"
                .add_condicion("email", txtusuario.Text)
                .add_condicion("administrador", "1")
            End With

            Dim dt As DataTable = con.SelectSQL(SEN.texto_sql)

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                Me.mensaje.Text = "El usuario no es correcto o no tiene permiso"
                MostrarAlerta()
                Exit Sub
            End If

            Dim login As String = dt.Rows(0)("email").ToString()
            Dim storedHash As String = dt.Rows(0)("password").ToString()

            ' 2) Verificar contraseña introducida contra el hash almacenado
            If String.IsNullOrWhiteSpace(storedHash) Then
                Me.mensaje.Text = "El usuario no es correcto o no tiene permiso"
                MostrarAlerta()
                Exit Sub
            End If

            If PasswordHasher.VerifyPassword(txtcontraseña.Text, storedHash) Then
                bien = True
                Session("usuario") = login
            Else
                Me.mensaje.Text = "El usuario no es correcto o no tiene permiso"
                MostrarAlerta()
                Exit Sub
            End If

        Catch ex As Exception
            conexion_bd = False
            Me.mensaje.Text = "Imposible conectar con la BD"
            MostrarAlerta()
        Finally
            If con IsNot Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try

        If bien Then
            Grabar_Accion("Login = '" & Session("usuario") & "'", "login", Now())
            Response.Redirect("~/ListRegistro.aspx")
        End If
    End Sub

    Private Sub Grabar_Accion(ByVal detail As String, ByVal accion As String, ByVal fechaahora As Date)

        Dim con As ConexionSQL
        Dim ins As insert_sql
        Dim upd = New Update_sql

        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        If fechaahora = Nothing Then fechaahora = Now()
        Try

            ins = New insert_sql
            With ins
                .Tabla = "auditLog"
                .add_valor("detail", detail)
                .add_valor("action", accion)
                .add_valor_expr("date", Util.fechasql(fechaahora, tipofecha.yyyyddMMmmss))
                .add_valor("email", Session("usuario"))
            End With
            con.Ejecuta(ins.texto_sql)

        Catch ex As Exception


        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try


    End Sub


End Class
