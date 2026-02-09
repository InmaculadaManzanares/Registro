Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.IO

Partial Class Completar
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

        If SessionTimeOut.IsSessionTimedOut Then
            Response.Redirect("~/InicioGestion.aspx")
        End If

        If Session("usuario") Is Nothing Then
            Response.Redirect("~/InicioGestion.aspx")
        Else
            activarcalendario("fecha1")
            activarcalendario("fecha2")

            If Not Page.IsPostBack Then
                Dim con As ConexionSQL
                con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
                'fecha1.Text = Format(DateTime.Now.ToString("dd/MM/yyyy"), "yyyy-MM-ddTHH:mm") ' CDate(Now.Day & "/" & Now.Month & "/" & Now.Year)
                'fecha2.Text = String.Format("{0:yyyy-MM-dd}", DateTime.Now.ToString("dd/MM/yyyy")) 'CDate(Now.Day & "/" & Now.Month & "/" & Now.Year) 
                Dim fini As Date = CDate(con.ejecuta1v_string("select DATEADD(month, DATEDIFF(month, - 1, GETDATE()) - 2, 0)"))
                Dim ffin As Date = CDate(con.ejecuta1v_string("select DATEADD(ss, - 1, DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0))"))
                fecha1.Text = fini.Year & "-" & fini.Month.ToString.PadLeft(2, "0") & "-" & fini.Day.ToString.PadLeft(2, "0")
                fecha2.Text = ffin.Year & "-" & ffin.Month.ToString.PadLeft(2, "0") & "-" & ffin.Day.ToString.PadLeft(2, "0")
                txthora1.Text = "07:00"
                txthora2.Text = "15:00"

                cargacombo()
                Ocultar_Panel()
                Dim texttitulo As Label = DirectCast(Page.Master.FindControl("titulo"), Label)
                texttitulo.Text = "Completar Registros"
                panelresultado.Visible = False
            Else

            End If

        End If

    End Sub
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
    Private Sub cargacombo()



        Dim con As ConexionSQL
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Try
            comboempresa.Items.Clear()
            Me.Sqlempresa.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Me.Sqlempresa.SelectCommand = "select   0 as id ,''  as nombre union all select id, nombre from empresas"
            Me.comboempresa.DataSourceID = "Sqlempresa"
            Me.comboempresa.DataTextField = "nombre"
            Me.comboempresa.DataValueField = "id"
            comboempresa.DataBind()

            combonombre.Items.Clear()
            Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios"
            Me.combonombre.DataSourceID = "Sqlemail"
            Me.combonombre.DataTextField = "nombre"
            Me.combonombre.DataValueField = "email"
            comboempresa.DataBind()

            combocentro.Items.Clear()
            Me.Sqlcentro.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Me.Sqlcentro.SelectCommand = "select   0 as id ,''  as nombre union all select id, nombre from centros"
            Me.combocentro.DataSourceID = "Sqlcentro"
            Me.combocentro.DataTextField = "nombre"
            Me.combocentro.DataValueField = "id"
            combocentro.DataBind()
        Catch ex As Exception
            msj_alerta = ex.Message
            mostraralerta(msj_alerta, 2)
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try

    End Sub
    Private Function filtrocampo(campo As String, datos As String) As String
        Dim dato As String() = datos.ToString.Split(";"c)
        Dim x As Integer
        Dim filtro As String = ""
        For x = 0 To dato.Length - 1
            If dato(x).Trim <> VACIO Then filtro = filtro & campo & " = '" & dato(x).Trim & "' or "
        Next
        filtro = " ( " & Left(filtro, filtro.Length - 4) & " ) "
        Return filtro
    End Function
    Private Sub rellena_combo(combo As DropDownList, cadsql As String)
        Dim con As ConexionSQL
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Try
            combo.Items.Clear()
            Dim vc As New ValoresCombo
            Dim dt As DataTable
            Dim dr As DataRow
            dt = con.SelectSQL(cadsql)
            For Each dr In dt.Rows
                combo.Items.Add(dr(0).ToString)
            Next
        Catch ex As Exception
            msj_alerta = ex.Message
            mostraralerta(msj_alerta, 2)
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try
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
                Completar.Visible = True
            Else
                panelresultado.Visible = False
                Completar.Visible = False
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
             MyDates WHERE  DATEPART(dw, mydate)  < 6 and (festivos = '' or festivos is null) and (MyDates.mydate >= 
             CONVERT(DATETIME, '" & CDate(f1) & "', 103)) 
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
			 RegistroHorario.Fecha,RegistroHorario.oficina HAVING   (RegistroHorario.oficina = N'VACACIONES') or  (RegistroHorario.oficina = N'BAJA')
			 ) ,
registrohor as ( select * from registrohoras union all select * from registrovaca)
             select   registro_2.Nempresa, registro_2.Centro,  registro_2.email,   registro_2.nombre, 		 RegistroHor.oficina as Lugar,
	         CONVERT(date, registro_2.fecha, 103)  as fecha,registrohor.entrada,registrohor.salida,registrohor.horas from
	         RegistroHor RIGHT OUTER JOIN
             registro AS registro_2 ON RegistroHor.email = registro_2.email  
			 AND CONVERT(date, RegistroHor.Fecha, 103) = CONVERT(date, registro_2.fecha, 103)  where 
            (Entrada is null or Salida is null and  RegistroHor.oficina <>'baja'  and  RegistroHor.oficina <> 'vacaciones' 
            and  RegistroHor.oficina <>'otros') "
            If comboempresa.Text <> "0" Then If filtrosql = "" Then filtrosql = " registro_2.empresa = " & v(comboempresa.Text) Else filtrosql = filtrosql & " and registro_2.empresa = " & v(comboempresa.Text)
            If combocentro.Text <> "0" Then If filtrosql = "" Then filtrosql = " registro_2.centro = " & v(combocentro.SelectedItem.Text) Else filtrosql = filtrosql & " and registro_2.centro = " & v(combocentro.SelectedItem.Text)
            If combonombre.Text <> "" Then If filtrosql = "" Then filtrosql = " registro_2.email = " & v(combonombre.Text) Else filtrosql = filtrosql & " and registro_2.email = " & v(combonombre.Text)

            If filtrosql <> "" Then filtrosql = " and " & filtrosql
            cadena = cadena & filtrosql & " ORDER BY registro_2.Nempresa, registro_2.Centro, fecha, registro_2.email "

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
        Completar.Visible = False
        panelresultado.Visible = False
        cargacombo()
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
    Private Sub mostraralerta(cadena As String, ByVal tipo As Integer)

        mensajepopup.Visible = True
        msgSolicitud.Text = cadena
        If tipo = 1 Then
            mensajepopup.CssClass = "alertmsg alert-success"
        Else
            mensajepopup.CssClass = "alertmsg alert-danger"
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



    Protected Sub Excel_Click(sender As Object, e As EventArgs) Handles Excel.Click

        Dim sb As StringBuilder = New StringBuilder()
        Dim sw As IO.StringWriter = New IO.StringWriter(sb)
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
        Dim pagina As Page = New Page
        Dim form = New HtmlForm
        Dim cadena As String = ""
        Dim filtrosql As String = ""
        Dim detailreg As String = ""
        ocultar_mensaje()
        If fecha1.Text <> "" AndAlso IsDate(fecha1.Text) = False Then
            msj_alerta = "La fecha inicio no es correcta"
            mostraralerta(msj_alerta, 2)
            Exit Sub
        End If
        If fecha2.Text <> "" AndAlso IsDate(fecha2.Text) = False Then
            msj_alerta = "La fecha fin no es correcta"
            mostraralerta(msj_alerta, 2)
            Exit Sub
        End If
        If fecha1.Text <> "" And fecha2.Text <> "" AndAlso CDate(fecha1.Text) > CDate(fecha2.Text) = True Then
            msj_alerta = "La fecha inicio es superior a la fecha fin"
            mostraralerta(msj_alerta, 2)
            Exit Sub
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
			 RegistroHorario.Fecha,RegistroHorario.oficina HAVING   (RegistroHorario.oficina = N'VACACIONES') or  (RegistroHorario.oficina = N'BAJA')
			 ) ,
registrohor as ( select * from registrohoras union all select * from registrovaca)
             select   registro_2.Nempresa, registro_2.Centro,  registro_2.email, registro_2.nombre,	 
	          left(registro_2.fecha,10) as fecha,RegistroHor.oficina as Lugar, registrohor.entrada,registrohor.salida,registrohor.horas from
	         RegistroHor RIGHT OUTER JOIN
             registro AS registro_2 ON RegistroHor.email = registro_2.email  AND CONVERT(date, RegistroHor.Fecha, 103) = CONVERT(date, registro_2.fecha, 103) "
        If comboempresa.Text <> "0" Then If filtrosql = "" Then filtrosql = " registro_2.empresa = " & v(comboempresa.Text) Else filtrosql = filtrosql & " and registro_2.empresa = " & v(comboempresa.Text)
        If combocentro.Text <> "0" Then If filtrosql = "" Then filtrosql = " registro_2.centro = " & v(combocentro.SelectedItem.Text) Else filtrosql = filtrosql & " and registro_2.centro = " & v(combocentro.SelectedItem.Text)
        If combonombre.Text <> "" Then If filtrosql = "" Then filtrosql = " registro_2.email = " & v(combonombre.Text) Else filtrosql = filtrosql & " and registro_2.email = " & v(combonombre.Text)
        If filtrosql <> "" Then filtrosql = " where " & filtrosql
        cadena = cadena & filtrosql & " ORDER BY registro_2.Nempresa, registro_2.Centro,  registro_2.email,fecha "



        Dim strFile As String = Server.MapPath("~/descargas/Registro_") + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv"

        Dim sw1 As StreamWriter

        sw1 = New StreamWriter(New FileStream(strFile, FileMode.Create, FileAccess.ReadWrite), Encoding.Default)

        Dim NColummna As Integer = 0
        Using con As New SqlConnection(VariablesGlobales.CadenadeConexion(Application))
            Using cmd As New SqlCommand(cadena)
                Using sda As New SqlDataAdapter()
                    cmd.Connection = con
                    sda.SelectCommand = cmd
                    Using dt As New DataTable()
                        sda.Fill(dt)

                        Dim csv As String = String.Empty
                        NColummna = 1
                        For Each column As DataColumn In dt.Columns

                            If NColummna < dt.Columns.Count Then
                                csv += column.ColumnName + ";"c
                            Else
                                csv += column.ColumnName
                            End If

                            NColummna = NColummna + 1
                        Next
                        csv += vbCr & vbLf
                        sw1.Write(csv)
                        csv = ""
                        Dim mlinea As Long
                        For Each row As DataRow In dt.Rows
                            mlinea = mlinea + 1
                            NColummna = 1
                            For Each column As DataColumn In dt.Columns
                                If NColummna < dt.Columns.Count Then
                                    csv += row(column.ColumnName).ToString().Replace(",", ",") + ";"c
                                Else
                                    csv += row(column.ColumnName).ToString().Replace(",", ",")
                                End If
                                NColummna = NColummna + 1
                            Next
                            csv += vbCr & vbLf
                            sw1.Write(csv)
                            csv = ""
                        Next
                        sw1.Close()

                        Dim file As System.IO.FileInfo = New System.IO.FileInfo(strFile)
                        If (file.Exists) Then

                            Response.ContentType = "text/csv"
                            Response.AppendHeader("Content-Disposition", "Attachment; Filename=" + file.Name + "")
                            Response.TransmitFile(strFile)
                            Response.End()
                        End If
                    End Using
                End Using
            End Using
        End Using

    End Sub

    Private Sub comboempresa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboempresa.SelectedIndexChanged
        If comboempresa.Text = "0" Then
            combonombre.Items.Clear()
            If combocentro.Text = "0" Then
                Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios"
                Me.combonombre.DataSourceID = "Sqlemail"
                Me.combonombre.DataTextField = "nombre"
                Me.combonombre.DataValueField = "email"
            Else
                Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios where centro = " & combocentro.Text
                Me.combonombre.DataSourceID = "Sqlemail"
                Me.combonombre.DataTextField = "nombre"
                Me.combonombre.DataValueField = "email"
            End If
            combonombre.DataBind()
        Else
            combonombre.Items.Clear()
            If combocentro.Text = "0" Then
                Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios where empresa = " & comboempresa.Text
                Me.combonombre.DataSourceID = "Sqlemail"
                Me.combonombre.DataTextField = "nombre"
                Me.combonombre.DataValueField = "email"
                combonombre.DataBind()
            Else
                Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios where empresa = " & comboempresa.Text & " and centro = " & combocentro.Text
                Me.combonombre.DataSourceID = "Sqlemail"
                Me.combonombre.DataTextField = "nombre"
                Me.combonombre.DataValueField = "email"
                combonombre.DataBind()
            End If

        End If
        resultado.Text = ""
        Completar.Visible = False
        matriz.DataSource = Nothing
        activar_panel_lista()
        Ocultar_Panel()
    End Sub

    Private Sub combonombre_SelectedIndexChanged(sender As Object, e As EventArgs) Handles combonombre.SelectedIndexChanged
        resultado.Text = ""
        Completar.Visible = False
        matriz.DataSource = Nothing
        activar_panel_lista()
        Ocultar_Panel()
    End Sub

    Private Sub combocentro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles combocentro.SelectedIndexChanged
        If comboempresa.Text = "0" Then
            combonombre.Items.Clear()
            If combocentro.Text = "0" Then
                Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios"
                Me.combonombre.DataSourceID = "Sqlemail"
                Me.combonombre.DataTextField = "nombre"
                Me.combonombre.DataValueField = "email"
            Else
                Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios where centro = " & combocentro.Text
                Me.combonombre.DataSourceID = "Sqlemail"
                Me.combonombre.DataTextField = "nombre"
                Me.combonombre.DataValueField = "email"
            End If
            combonombre.DataBind()
        Else
            combonombre.Items.Clear()
            If combocentro.Text = "0" Then
                Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios where empresa = " & comboempresa.Text
                Me.combonombre.DataSourceID = "Sqlemail"
                Me.combonombre.DataTextField = "nombre"
                Me.combonombre.DataValueField = "email"
                combonombre.DataBind()
            Else
                Me.Sqlemail.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.Sqlemail.SelectCommand = "select '' as email ,''  as nombre union all select email, nombre from usuarios where empresa = " & comboempresa.Text & " and centro = " & combocentro.Text
                Me.combonombre.DataSourceID = "Sqlemail"
                Me.combonombre.DataTextField = "nombre"
                Me.combonombre.DataValueField = "email"
                combonombre.DataBind()
            End If

        End If
        resultado.Text = ""
        Completar.Visible = False
        matriz.DataSource = Nothing
        activar_panel_lista()
        Ocultar_Panel()
    End Sub

    Private Sub Completar_Click(sender As Object, e As EventArgs) Handles Completar.Click
        If txthora1.Text <> "" And txthora2.Text <> "" And lugar.Text <> "" Then




            Dim con As ConexionSQL
            con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
            Try
                Dim cadsql As String = monte_sql("matrix")
                Dim dt As DataTable
                Dim dr As DataRow
                Dim ins As insert_sql
                Dim fecha1 As String
                Dim fecha2 As String
                dt = con.SelectSQL(cadsql)
                For Each dr In dt.Rows

                    If dr(2).ToString <> "" Then
                        'dia de la semana
                        Dim valorfecha As DateTime = Left(dr(5).ToString, 10)
                        Dim diaSemana As DayOfWeek = valorfecha.DayOfWeek
                        Dim diaSemanaEspanol As String = ObtenerDiaSemanaEspanol(diaSemana)
                        Dim hora1 As String = con.ejecuta1v_string("select horainicio from usuarioshorarios where email = '" & dr(2).ToString &
                                                                    "' and dia = '" & diaSemanaEspanol & "'")
                        Dim hora2 As String = con.ejecuta1v_string("select horafin from usuarioshorarios where email = '" & dr(2).ToString &
                                                                    "' and dia = '" & diaSemanaEspanol & "'")
                        If hora1 <> "" And hora2 <> "" Then

                            If dr(6).ToString = "" Then

                                Dim horainicial As String = hora1
                                Dim time As TimeSpan = TimeSpan.Parse(horainicial)
                                Dim twoMinutes As TimeSpan = TimeSpan.FromMinutes(3)
                                Dim hora1menos2 As TimeSpan = time.Subtract(twoMinutes)
                                Dim hora1mas2 As TimeSpan = time.Add(twoMinutes)
                                Dim startTime As String = hora1menos2.ToString
                                Dim endTime As String = hora1mas2.ToString

                                ' Genera una hora aleatoria entre las dos horas dadas
                                Dim randomTimeinicial As String = ""
                                randomTimeinicial = GenerateRandomTime(startTime, endTime)

                                ins = New insert_sql
                                With ins
                                    .Tabla = "RegistroHorario"
                                    .add_valor(CAMPO_EMAIL, dr(2).ToString)
                                    fecha1 = ""
                                    fecha1 = Left(dr(5).ToString, 10) & " " & randomTimeinicial

                                    .add_valor_expr("fecha", Util.fechasql(fecha1, tipofecha.yyyyddMMmmss))
                                    .add_valor("oficina", lugar.Text)
                                    .add_valor("inOut", "Entrada")
                                    .add_valor("obs", lugar.Text)
                                    .add_valor("ubicacion", "Web automatica")
                                End With
                                con.Ejecuta(ins.texto_sql)
                            End If


                            If dr(7).ToString = "" Then

                                Dim horafinal As String = hora2
                                Dim finaltime As TimeSpan = TimeSpan.Parse(horafinal)
                                Dim finaltwoMinutes As TimeSpan = TimeSpan.FromMinutes(3)
                                Dim finalmenos2 As TimeSpan = finaltime.Subtract(finaltwoMinutes)
                                Dim finalmas2 As TimeSpan = finaltime.Add(finaltwoMinutes)
                                Dim finalstartTime As String = finalmenos2.ToString
                                Dim finalendTime As String = finalmas2.ToString

                                ' Genera una hora aleatoria entre las dos horas dadas
                                Dim randomTimefinal As String = ""
                                randomTimefinal = GenerateRandomTime(finalstartTime, finalendTime)
                                ins = New insert_sql
                                With ins
                                    .Tabla = "RegistroHorario"
                                    .add_valor(CAMPO_EMAIL, dr(2).ToString)
                                    fecha2 = ""
                                    fecha2 = Left(dr(5).ToString, 10) & " " & randomTimefinal

                                    .add_valor_expr("fecha", Util.fechasql(fecha2, tipofecha.yyyyddMMmmss))
                                    .add_valor("oficina", lugar.Text)
                                    .add_valor("inOut", "Salida")
                                    .add_valor("obs", lugar.Text)
                                    .add_valor("ubicacion", "Web automatica")
                                End With
                                con.Ejecuta(ins.texto_sql)
                            End If
                        End If

                        'turno tarde

                        Dim tardehora1 As String = con.ejecuta1v_string("select horainicio from usuarioshorarios where email = '" & dr(2).ToString &
                                              "' and dia = '" & diaSemanaEspanol & " tarde'")
                        Dim tardehora2 As String = con.ejecuta1v_string("select horafin from usuarioshorarios where email = '" & dr(2).ToString &
                                                                    "' and dia = '" & diaSemanaEspanol & " tarde'")
                        If tardehora1 <> "" And tardehora2 <> "" Then

                            Dim horainicialtarde As String = tardehora1
                            Dim timetarde As TimeSpan = TimeSpan.Parse(horainicialtarde)
                            Dim twoMinutestarde As TimeSpan = TimeSpan.FromMinutes(3)
                            Dim hora1menos2tarde As TimeSpan = timetarde.Subtract(twoMinutestarde)
                            Dim hora1mas2tarde As TimeSpan = timetarde.Add(twoMinutestarde)
                            Dim startTimetarde As String = hora1menos2tarde.ToString
                            Dim endTimetarde As String = hora1mas2tarde.ToString

                            ' Genera una hora aleatoria entre las dos horas dadas
                            Dim randomTimeinicialtarde As String = ""
                            randomTimeinicialtarde = GenerateRandomTime(startTimetarde, endTimetarde)

                            ins = New insert_sql
                            With ins
                                .Tabla = "RegistroHorario"
                                .add_valor(CAMPO_EMAIL, dr(2).ToString)
                                fecha1 = ""
                                fecha1 = Left(dr(5).ToString, 10) & " " & randomTimeinicialtarde
                                .add_valor_expr("fecha", Util.fechasql(fecha1, tipofecha.yyyyddMMmmss))
                                .add_valor("oficina", lugar.Text)
                                .add_valor("inOut", "Entrada")
                                .add_valor("obs", lugar.Text)
                                .add_valor("ubicacion", "Web automatica")
                            End With
                            con.Ejecuta(ins.texto_sql)

                            Dim horafinaltarde As String = tardehora2
                            Dim finaltimetarde As TimeSpan = TimeSpan.Parse(horafinaltarde)
                            Dim finaltwoMinutestarde As TimeSpan = TimeSpan.FromMinutes(3)
                            Dim finalmenos2tarde As TimeSpan = finaltimetarde.Subtract(finaltwoMinutestarde)
                            Dim finalmas2tarde As TimeSpan = finaltimetarde.Add(finaltwoMinutestarde)
                            Dim finalstartTimetarde As String = finalmenos2tarde.ToString
                            Dim finalendTimetarde As String = finalmas2tarde.ToString

                            ' Genera una hora aleatoria entre las dos horas dadas
                            Dim randomTimefinaltarde As String = ""
                            randomTimefinaltarde = GenerateRandomTime(finalstartTimetarde, finalendTimetarde)
                            ins = New insert_sql
                            With ins
                                .Tabla = "RegistroHorario"
                                .add_valor(CAMPO_EMAIL, dr(2).ToString)
                                fecha2 = ""
                                fecha2 = Left(dr(5).ToString, 10) & " " & randomTimefinaltarde

                                .add_valor_expr("fecha", Util.fechasql(fecha2, tipofecha.yyyyddMMmmss))
                                .add_valor("oficina", lugar.Text)
                                .add_valor("inOut", "Salida")
                                .add_valor("obs", lugar.Text)
                                .add_valor("ubicacion", "Web automatica")
                            End With
                            con.Ejecuta(ins.texto_sql)
                        End If


                        'fin turno tarde




                    End If
                Next
                msj_alerta = "Operación finalizada"
                mostraralerta(msj_alerta, 1)
                matriz.DataSource = Nothing
                Completar.Visible = False
                panelresultado.Visible = False
                cargacombo()
                activar_panel_lista()
                Ocultar_Panel()
            Catch ex As Exception
                msj_alerta = ex.Message
                mostraralerta(msj_alerta, 2)
            Finally
                If Not con Is Nothing Then
                    con.CerrarConexion()
                    con = Nothing
                End If
            End Try
        Else
            msj_alerta = "Revise las horas y el lugar"
            mostraralerta(msj_alerta, 2)
        End If

    End Sub
    Function ObtenerDiaSemanaEspanol(dia As DayOfWeek) As String
        Select Case dia
            Case DayOfWeek.Sunday
                Return "domingo"
            Case DayOfWeek.Monday
                Return "lunes"
            Case DayOfWeek.Tuesday
                Return "martes"
            Case DayOfWeek.Wednesday
                Return "miércoles"
            Case DayOfWeek.Thursday
                Return "jueves"
            Case DayOfWeek.Friday
                Return "viernes"
            Case DayOfWeek.Saturday
                Return "sábado"
            Case Else
                Return "Desconocido"
        End Select
    End Function
    Function GenerateRandomTime(ByVal startTime As String, ByVal endTime As String) As String
        ' Convierte las horas a TimeSpan
        System.Threading.Thread.Sleep(200)
        Dim start As TimeSpan
        start = TimeSpan.Parse(startTime)
        Dim [end] As TimeSpan
        [end] = TimeSpan.Parse(endTime)

        ' Convierte TimeSpan a segundos
        Dim startSeconds As Integer = 0
        Dim endSeconds As Integer = 0
        startSeconds = CInt(start.TotalSeconds)
        endSeconds = CInt([end].TotalSeconds)

        ' Genera un número aleatorio de segundos entre el rango
        Dim rand As New Random()
        Dim randomSeconds As Integer = 0
        randomSeconds = rand.Next(startSeconds, endSeconds + 1)

        ' Convierte los segundos aleatorios de vuelta a TimeSpan
        Dim randomTimeSpan As TimeSpan
        randomTimeSpan = TimeSpan.FromSeconds(randomSeconds)

        ' Devuelve la hora en formato "HH:mm:ss"
        Dim val As String = ""
        val = randomTimeSpan.ToString("hh\:mm\:ss")
        Return val
    End Function

End Class
