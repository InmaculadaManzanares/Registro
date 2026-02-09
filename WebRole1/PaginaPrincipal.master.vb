#Region "Librerias"

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Web
Imports System.IO

#End Region

Partial Class PaginaPrincipal
    Inherits System.Web.UI.MasterPage



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If SessionTimeOut.IsSessionTimedOut Then
            Response.Redirect("~/inicioGestion.aspx")
        End If

        If Session("usuario") Is Nothing Then
            Response.Redirect("~/inicioGestion.aspx")
        Else
            usu.Text = "(" & (Session("usuario")).ToString.Trim & ")"

        End If

        ' Inyectar el script en la página 
        Dim script1 As String = Util.CambiarColorControlesScript(Session("colorprimario"), Session("colorsecundario"))
            Page.ClientScript.RegisterStartupScript(Me.GetType(), "ReplaceColorsScript1", script1)

    End Sub



    Private ReadOnly Property UsuarioActual() As String
        Get
            Return VariablesGlobales.Usuario(Session, Application)
        End Get
    End Property

    Private Property ControlesActuales() As coleccion_control_visor
        Get
            Return VariablesGlobales.ControlesActualesVisor(Me.ID, Session)
        End Get
        Set(ByVal value As coleccion_control_visor)
            VariablesGlobales.ControlesActualesVisor(Me.ID, Session) = value
        End Set
    End Property




    Private Sub ocultar(menu As String)
        Dim script As String = String.Empty
        script = "desactiva_" + menu + " ();"
        If menu = "intro" Then script = "desactiva_intro();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
    End Sub

    Private Property ExtensionesAFullText() As String
        Get
            Return VariablesGlobales.ExtensionesAFullText(Session)
        End Get
        Set(ByVal value As String)
            VariablesGlobales.ExtensionesAFullText(Session) = value
        End Set
    End Property

    Private Property ContenedorFicheros() As String
        Get
            Return VariablesGlobales.ContenedorFicheros(Session)
        End Get
        Set(ByVal value As String)
            VariablesGlobales.ContenedorFicheros(Session) = value
        End Set
    End Property

    Private Property ContenedorDescarga() As String
        Get
            Return VariablesGlobales.ContenedorDescarga(Session)
        End Get
        Set(ByVal value As String)
            VariablesGlobales.ContenedorDescarga(Session) = value
        End Set
    End Property

    Private Property CadenaConexionDocBox() As String
        Get
            Return VariablesGlobales.CadenaConexionDocBox(Session)
        End Get
        Set(ByVal value As String)
            VariablesGlobales.CadenaConexionDocBox(Session) = value
        End Set
    End Property

    Private Function TienePermisos(ByVal id As Long, ByVal codusuario As String, ByVal permiso As String) As Boolean
        Dim sen As New SentenciaSQL
        Dim con As ConexionSQL = Nothing

        Try
            con = New ConexionSQL(Me.CadenaConexionDocBox)
            sen.sql_select = Util.func(UDF_PERMISOS, id.ToString, v(codusuario), v(permiso))
            Return con.ejecuta1v_long(sen.texto_sql) > 0
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try
    End Function
    Private Function TienePermisos(ByVal id As Long) As Boolean
        Dim sen As New SentenciaSQL
        Dim con As ConexionSQL = Nothing

        Try
            con = New ConexionSQL(Me.CadenaConexionDocBox)
            With sen
                .sql_select = "COUNT(*) as c"
                .sql_from = TABLA_PERMISOS
                .add_condicion(CAMPO_IDCUADRO, id)

            End With
            Return con.ejecuta1v_long(sen.texto_sql) > 0
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try
    End Function



    Protected Sub BCerrarsesion_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BCerrarsesion.Click

        logout()

    End Sub

    ''' <summary>
    ''' Metodo encargado del Cierre de Sesión y limpieza de cookies.
    ''' </summary>
    Private Sub logout()

        Try
            ' Desconectamos la sesión y limpiamos las cookies.
            Session.Abandon()
            Response.Cookies.Add(New HttpCookie("ASP.NET_SessionId", ""))

            ' Redireccionamos a la pantalla de login.
            Response.Redirect("~/inicioGestion.aspx")

        Catch ex As Exception

        End Try

        ocultarcargando()

    End Sub

    'Ocultar Cargando
    Private Sub ocultarcargando()
        Dim script As String = String.Empty
        script = "ocultarcargando();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
    End Sub



End Class

''' <summary>
''' Clase que contiene un método para comprobar si la sesión actual está o no caducada.
''' </summary>
Public Class SessionTimeOut
    Public Shared Function IsSessionTimedOut() As Boolean
        Dim ctx As HttpContext = HttpContext.Current
        If ctx Is Nothing Then
            Throw New Exception("Este método sólo se puede usar en una aplicación Web")
        End If

        'Comprobamos que haya sesión en primer lugar 
        '(por ejemplo si por ejemplo EnableSessionState=false)
        If ctx.Session Is Nothing Then
            Return False
        End If
        'Si no hay sesión, no puede caducar
        'Se comprueba si se ha generado una nueva sesión en esta petición
        If Not ctx.Session.IsNewSession Then
            Return False
        End If
        'Si no es una nueva sesión es que no ha caducado
        Dim objCookie As HttpCookie = ctx.Request.Cookies("ASP.NET_SessionId")
        'Esto en teoría es imposible que pase porque si hay una 
        'nueva sesión debería existir la cookie, pero lo compruebo porque
        'IsNewSession puede dar True sin ser cierto (más en el post)
        If objCookie Is Nothing Then
            Return False
        End If

        'Si hay un valor en la cookie es que hay un valor de sesión previo, pero como la sesión 
        'es nueva no debería estar, por lo que deducimos que la sesión anterior ha caducado
        If Not String.IsNullOrEmpty(objCookie.Value) Then
            Return True
        Else
            Return False
        End If
    End Function
End Class