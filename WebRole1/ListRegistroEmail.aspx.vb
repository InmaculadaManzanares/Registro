Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Partial Class ListRegistroEmail
    Inherits System.Web.UI.Page

    Dim msj_alerta As String = String.Empty
    Dim cadenasql_vacia As String = " union all SELECT '' AS Expr1, '' AS Expr2"
    Dim cadenasql_cero As String = " union all SELECT '0' AS Expr1, '' AS Expr2"
    Dim valor_cliente As String
    Dim valor_expediente As String
    Dim CAMPOS_ARCHIVO As String
    Dim cadena_sinmaximo As String
    Dim CadenaExcel As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        If Not Page.IsPostBack Then
            Dim con As ConexionSQL = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
            Dim fini As Date = CDate(con.ejecuta1v_string("select  GETDATE()"))
            Dim ffin As Date = CDate(con.ejecuta1v_string("select GETDATE()"))
            fecha1.Text = fini.Year & "-" & fini.Month.ToString.PadLeft(2, "0") & "-" & "01"
            fecha2.Text = ffin.Year & "-" & ffin.Month.ToString.PadLeft(2, "0") & "-" & ffin.Day.ToString.PadLeft(2, "0")

            'ListRegistroEmail?Email=elopez@spaceland.es
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
            Session("usuario") = ""
            Dim email As String = ""
            email = Request.QueryString("Email")

            If String.IsNullOrWhiteSpace(email) Then
                panelresultado.Visible = False
                mostraralerta("Enlace no válido o incompleto. Vuelve a acceder desde el link recibido.", 2)
                Exit Sub
            End If

            Dim nombre As String = Decrypt(email)

            If String.IsNullOrWhiteSpace(nombre) Then
                panelresultado.Visible = False
                mostraralerta("El enlace ha caducado o no es válido. Solicita uno nuevo.", 2)
                Exit Sub
            End If

            Session("usuario") = nombre
            txtnombre.Value = nombre
            lblnombre.Text = nombre


            filtro()
        End If
        ' Inyectar el script en la página
        Dim script1 As String = Util.CambiarColorControlesScript(Session("colorprimario"), Session("colorsecundario"))
        ClientScript.RegisterStartupScript(Me.GetType(), "ReplaceColorsScript1", script1)



    End Sub
    Public Function Decrypt(cipherText As String) As String
        Try
            If String.IsNullOrWhiteSpace(cipherText) Then
                Throw New ArgumentNullException(NameOf(cipherText))
            End If

            Dim masterKeyB64 As String = System.Configuration.ConfigurationManager.AppSettings("UrlCryptoMasterKey")
            If String.IsNullOrWhiteSpace(masterKeyB64) Then
                Throw New Exception("Falta la clave UrlCryptoMasterKey en web.config.")
            End If

            ' cipherText viene ya en Base64URL (no hace falta Replace(" ","+"))
            Return CryptoAes.DecryptFromBase64Url(cipherText, masterKeyB64)

        Catch ex As Exception
            msj_alerta = ex.Message
            mostraralerta(msj_alerta, 2)
            Return ""
        End Try
    End Function


    Private Sub activarcalendario(ByVal txt As String)
        Dim script As String = "activar_calendario('" & "#" & "ContentPlaceHolder1_" & txt & "');"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
    End Sub


    Private Sub Grabar_Accion(ByVal detail As String, ByVal accion As String, ByVal fechaahora As Date)
        Dim resul As Long
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

            msj_alerta = ex.Message
            mostraralerta(msj_alerta, 2)
            resul = -1
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try


    End Sub
    Private Sub desactivar(combo As DropDownList)
        combo.Attributes.Add("class", "form-control no-arrow")
        combo.Attributes.Add("disabled", "disabled")
    End Sub
    Private Sub activar(combo As DropDownList)
        combo.Attributes.Add("class", "form-control")
        combo.Attributes.Remove("disabled")
    End Sub

    Protected Sub CargaMatriz(ByVal cadena As String)
        Dim con As ConexionSQL
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Try

            Me.SqlDS.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Me.SqlDS.SelectCommand = cadena
            Me.matriz.DataSourceID = "SqlDS"
            Me.matriz.DataBind()
            If matriz.Rows.Count > 0 Then
                matriz.PagerSettings.LastPageText = matriz.PageCount
                matriz.PagerSettings.FirstPageText = "1"
                panelresultado.Visible = True

            Else
                panelresultado.Visible = False
            End If

        Catch ex As Exception
            msj_alerta = "Error al cargar matriz: " & ex.Message
            mostraralerta(msj_alerta, 2)
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try
    End Sub
    Protected Sub Matriz_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles matriz.PageIndexChanging
        matriz.PageIndex = e.NewPageIndex
        Me.filtro()
    End Sub

    Private Sub matriz_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles matriz.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then

            If e.Row.Cells.Count - 1 = matriz.HeaderRow.Cells.Count - 1 Then



            End If
        End If
    End Sub

    Protected Sub Matriz_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles matriz.Sorted
        Me.filtro()
    End Sub
    Protected Sub Filtrar_Click(sender As Object, e As EventArgs) Handles Filtrar.Click
        ocultar_mensaje()
        filtro()
        activar_panel_lista()
        Ocultar_Panel()
    End Sub
    Protected Sub Salir_Click(sender As Object, e As EventArgs) Handles Salir.Click
        Response.Redirect("~/RegistroHorario.aspx")
    End Sub
    Public Sub filtro()
        CargaMatriz(monte_sql("matriz"))
    End Sub
    Private Function monte_sql(tipo As String) As String
        Dim cadena As String = ""
        Dim filtrosql As String = ""
        Dim detailreg As String = ""
        Dim con As ConexionSQL
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Try
            ocultar_mensaje()
            If fecha1.Text <> "" AndAlso IsDate(fecha1.Text) = False Then
                msj_alerta = "La fecha inicio no es correcta"
                mostraralerta(msj_alerta, 2)
                Return ""
            End If
            If fecha2.Text <> "" AndAlso IsDate(fecha2.Text) = False Then
                msj_alerta = "La fecha fin no es correcta"
                mostraralerta(msj_alerta, 2)
                Return ""
            End If
            If fecha1.Text <> "" And fecha2.Text <> "" AndAlso CDate(fecha1.Text) > CDate(fecha2.Text) = True Then
                msj_alerta = "La fecha inicio es superior a la fecha fin"
                mostraralerta(msj_alerta, 2)
                Return ""
            End If
            Dim f1 As DateTime
            Dim f2 As DateTime

            If IsDate(fecha1.Text) Then f1 = CDate(fecha1.Text & " 00:00:00")

            If IsDate(fecha2.Text) Then f2 = CDate(fecha2.Text & " 23:59:59")

            cadena = "with registro  as(SELECT        Usuarios.Nombre AS nombre, Usuarios.email, MyDates.mydate AS fecha, Usuarios.Empresa AS empresa, empresas.Nombre AS Nempresa, centros.Nombre AS Centro
             FROM Usuarios LEFT OUTER JOIN
             centros ON Usuarios.Centro = centros.Id LEFT OUTER JOIN
             empresas ON Usuarios.Empresa = empresas.Id CROSS JOIN
             MyDates WHERE
             (MyDates.mydate >= CONVERT(DATETIME, '" & CDate(f1) & "', 103)) 
             AND (MyDates.mydate <= CONVERT(DATETIME, '" & CDate(f2) & "', 103))),   	
 registrohoras as  ( SELECT        registro_1.Nempresa, registro_1.Centro, registro_1.empresa, registro_1.email, 		 RegistroHorario.oficina,
	         registro_1.nombre, CONVERT(date, registro_1.fecha, 103) AS fecha, CONVERT(VARCHAR(5), RegistroHorario.Fecha, 108) 
             AS Entrada,
             (SELECT        TOP (1) CONVERT(VARCHAR(5), Fecha, 108) AS Expr1
             FROM            RegistroHorario AS rh
             WHERE        (RegistroHorario.email = email) AND (inOut = 'salida') AND (CONVERT(datetime, Fecha, 103) > CONVERT(datetime, RegistroHorario.Fecha, 103))
             and (CONVERT(date, Fecha, 103) = CONVERT(date, RegistroHorario.Fecha, 103))
             ORDER BY Fecha) AS Salida, CONVERT(varchar(5), DATEADD(second, DATEDIFF(SS, RegistroHorario.Fecha,
             (SELECT        TOP (1) Fecha
             FROM            RegistroHorario AS rh
             WHERE        (RegistroHorario.email = email) AND (inOut = 'salida') AND (CONVERT(datetime, Fecha, 103) > CONVERT(datetime, RegistroHorario.Fecha, 103))
             and (CONVERT(date, Fecha, 103) = CONVERT(date, RegistroHorario.Fecha, 103))             
             ORDER BY Fecha)), 0), 108)
			 AS horas
             FROM  RegistroHorario RIGHT OUTER JOIN
             registro AS registro_1 ON RegistroHorario.email = registro_1.email AND CONVERT(date, RegistroHorario.Fecha, 103) = CONVERT(date, registro_1.fecha, 103)
             GROUP BY registro_1.Nempresa, registro_1.Centro, registro_1.empresa, registro_1.email, RegistroHorario.email, registro_1.nombre, registro_1.fecha, RegistroHorario.Fecha, RegistroHorario.inOut
             		 ,RegistroHorario.oficina HAVING   (RegistroHorario.inOut = N'entrada')) ,				 					 
 registrovaca as  ( SELECT        registro_1.Nempresa, registro_1.Centro, registro_1.empresa, registro_1.email, 		 RegistroHorario.oficina,
	         registro_1.nombre, CONVERT(date, registro_1.fecha, 103) AS fecha, ''
             AS Entrada, '' AS Salida, '' as Horas
			  FROM  RegistroHorario  RIGHT OUTER JOIN
             registro AS registro_1 ON RegistroHorario.email = registro_1.email AND CONVERT(date, RegistroHorario.Fecha, 103) = CONVERT(date, registro_1.fecha, 103)
             GROUP BY registro_1.Nempresa, registro_1.Centro, registro_1.empresa, registro_1.email, RegistroHorario.email, registro_1.nombre, registro_1.fecha,
			 RegistroHorario.Fecha,RegistroHorario.oficina HAVING   (RegistroHorario.oficina = N'VACACIONES') or  (RegistroHorario.oficina = N'BAJA') or (RegistroHorario.oficina = N'OTROS')
			 ) ,
             registrohor as ( select * from registrohoras union all select * from registrovaca)
             select   registro_2.Nempresa, registro_2.Centro,  registro_2.email,   registro_2.nombre, 				 case   DATEPART(dw, registro_2.fecha)  			 
			 when 1 then 'Lunes' when 2 then 'Martes' when 3 then 'Miercoles' when 4 then 'Jueves' 
			  when 5 then 'Viernes' when 6 then 'Sabado' when 7 then 'Domingo' end as Dia,		 RegistroHor.oficina as Lugar,
	         CONVERT(date, registro_2.fecha, 103)  as fecha,registrohor.entrada,registrohor.salida,registrohor.horas from
	         RegistroHor RIGHT OUTER JOIN
             registro AS registro_2 ON RegistroHor.email = registro_2.email  
			 AND CONVERT(date, RegistroHor.Fecha, 103) = CONVERT(date, registro_2.fecha, 103)  where
             registro_2.email = " & v(Session("usuario"))

            cadena = cadena & " ORDER BY registro_2.Nempresa, registro_2.Centro, fecha, registro_2.email "

            Grabar_Accion("busqueda = '" & filtrosql, "busqueda", Now())
            Return cadena

        Catch ex As Exception
            cadena = ""
            Return cadena
            msj_alerta = ex.Message
            mostraralerta(msj_alerta, 2)
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try


    End Function
    Private Function monte_sqldetail(email As String, fecha As String) As String
        Dim cadena As String = ""
        Dim filtrosql As String = ""
        Try

            cadena = "select fecha,inout as 'Entrada/Salida', oficina as lugar, obs, Ubicacion
            ,Dispositivo  from RegistroHorario where  CONVERT(DATE, fecha, 103) =
            CONVERT(DATE, '" & fecha & "', 103) 
			and email = '" & email & "' order by  CONVERT(DATEtime, fecha, 103) "





            Return cadena
        Catch ex As Exception

            Return cadena
            msj_alerta = ex.Message
            mostraralerta(msj_alerta, 2)
        End Try

    End Function

    Private Sub activar_panel()
        Dim script As String = "activapanel();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "activapanel()", script, True)
    End Sub

    Private Sub activar_panel_lista()

        Dim script As String = "activapanellista();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "activapanellista()", script, True)
    End Sub

    'Ocultar Panel Detalle
    Private Sub Ocultar_Panel()
        Dim script As String = "ocultarpanel();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "ocultarpanel()", script, True)
    End Sub
    '----- funciones que deben ser globales MOVER
    Protected Function Scrub(ByVal text As String) As String
        Return text.Replace("&nbsp;", "")
    End Function


    Protected Sub Limpiar_Click(sender As Object, e As EventArgs) Handles Limpiar.Click
        ocultar_mensaje()
        matriz.DataSource = Nothing
        panelresultado.Visible = False
        activar_panel_lista()
        Ocultar_Panel()
    End Sub


    Private Sub logout()

        Try
            ' Desconectamos la sesión y limpiamos las cookies.
            Session.Abandon()
            Response.Cookies.Add(New HttpCookie("ASP.NET_SessionId", ""))
            ' Redireccionamos a la pantalla de login.
            Response.Redirect("~/LoginCoriel.aspx")

        Catch ex As Exception

        End Try


    End Sub
    ' Mostrar Alerta 
    'Private Sub mostraralerta(cadena As String, ByVal tipo As Integer)

    '    mensajepopup.Visible = True
    '    msgSolicitud.Text = cadena
    '    If tipo = 1 Then
    '        mensajepopup.CssClass = "alertmsg alert-success"
    '    Else
    '        mensajepopup.CssClass = "alertmsg alert-danger"
    '    End If

    'End Sub
    Private Sub mostraralerta(cadena As String, ByVal tipo As Integer)
        mensajepopup.Visible = True
        msgSolicitud.Text = cadena

        If tipo = 1 Then
            mensajepopup.CssClass = "alert alert-success"
        Else
            mensajepopup.CssClass = "alert alert-danger"
        End If
    End Sub

    ' Limpiar Campos 
    Private Sub limpiar_campo_email()
        Dim script As String = "limpiar_campos();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
    End Sub

    Private Sub ocultar_mensaje()

        mensajepopup.Visible = False
    End Sub


    Public Function ceros(ByVal Nro As String, ByVal Cantidad As Integer) As String
        Dim numero As String, cuantos As String, i As Integer
        numero = Trim(Nro) 'Trim quita los espacion en blanco 
        cuantos = "0"
        For i = 1 To Cantidad
            cuantos = cuantos & "0"
        Next i
        ceros = Mid(cuantos, 1, Cantidad - Len(numero)) & numero
    End Function








End Class
