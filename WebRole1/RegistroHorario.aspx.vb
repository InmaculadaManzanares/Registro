Imports System
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class RegistroHorario
    Inherits System.Web.UI.Page
    Public USUARIO As String
    Private Shared ReadOnly Key As Byte() = Encoding.UTF8.GetBytes("1234567890123456") ' 16 bytes para AES-128
    Private Shared ReadOnly IV As Byte() = Encoding.UTF8.GetBytes("6543210987654321") ' 16 bytes para AES-128

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

        End If
        ' Inyectar el script en la página
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
        Dim con As ConexionSQL
        Dim SEN As New SentenciaSQL
        Dim conexion_bd As Boolean = True
        Dim serr As String = String.Empty
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Try

            With SEN
                .Limpia()
                .sql_select = "email"
                .sql_from = "Usuarios"
                .add_condicion("password", txtcontraseña.Text)
                .add_condicion("email", txtusuario.Text)
            End With
            USUARIO = con.ejecuta1v_string(SEN.texto_sql)
            If USUARIO.Trim <> "" Then
                bien = True
                Session("usuario") = USUARIO
            Else
                Me.mensaje.Text = "El usuario no es correcto o no está registrado"
                MostrarAlerta()
            End If

            If bien Then
                list.Visible = True
                panel1.Visible = False
                With SEN
                    .Limpia()
                    .sql_select = "nombre"
                    .sql_from = "Usuarios"
                    .add_condicion("email", txtusuario.Text)
                End With
                txtnombre.Text = con.ejecuta1v_string(SEN.texto_sql)
                Dim inout As String = con.ejecuta1v_string("select top 1 ISNULL(inout, '') as inout from RegistroHorario 
                where email=" & v(txtusuario.Text) & " and FORMAT(fecha, 'yyyyMMdd') = FORMAT(getdate(), 'yyyyMMdd') order by fecha desc")
                If inout = "Entrada" Then EntradaSalida.Text = "Salida" Else EntradaSalida.Text = "Entrada"

                panel2.Visible = True

            End If


        Catch ex As Exception
            conexion_bd = False
            Me.mensaje.Text = "Imposible conectar con la BD"
            MostrarAlerta()
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
            End If
        End Try
    End Sub



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
        Dim valor As String = Session("usuario")
        Dim valor1 As String = Encrypt(valor)

        Dim valor2 As String = Decrypt(valor1)
        Dim valornuevo As String = Uri.EscapeDataString(valor1)
        Response.Redirect("~/ListRegistroEmail?Email=" & valornuevo)
    End Sub

    Public Shared Function Encrypt(plainText As String) As String
        If String.IsNullOrEmpty(plainText) Then
            Throw New ArgumentNullException(NameOf(plainText))
        End If

        Using aes As Aes = aes.Create()
            aes.Key = Key
            aes.IV = IV

            Dim encryptor As ICryptoTransform = aes.CreateEncryptor(aes.Key, aes.IV)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor, CryptoStreamMode.Write)
                    Using sw As New StreamWriter(cs)
                        sw.Write(plainText)
                    End Using
                End Using
                Return Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
    End Function

    Public Shared Function Decrypt(cipherText As String) As String
        Try
            If String.IsNullOrEmpty(cipherText) Then
                Throw New ArgumentNullException(NameOf(cipherText))
            End If

            Dim cipherBytes = Convert.FromBase64String(cipherText) ' Verificar si es Base64 válido
            Using aes As Aes = Aes.Create()
                aes.Key = Key
                aes.IV = IV

                Dim decryptor As ICryptoTransform = aes.CreateDecryptor(aes.Key, aes.IV)
                Using ms As New MemoryStream(cipherBytes)
                    Using cs As New CryptoStream(ms, decryptor, CryptoStreamMode.Read)
                        Using sr As New StreamReader(cs)
                            Return sr.ReadToEnd()
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As FormatException
            Throw New FormatException("La cadena proporcionada no es válida como Base64.", ex)
        End Try
    End Function

End Class