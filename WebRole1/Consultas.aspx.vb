Imports System
Imports System.Data
Imports System.Data.SqlClient
Partial Class Consultas
    Inherits System.Web.UI.Page
    Dim msj_alerta As String = String.Empty

    Protected Sub BBuscar_Click(ByVal sender As Object, ByVal e As EventArgs)

        If consulta.Text <> "" Then
            filtro()
        End If
    End Sub

    Protected Sub bLimpiar_Click(ByVal sender As Object, ByVal e As EventArgs)
        limpia()
        Me.combo_consulta.SelectedValue = 0
    End Sub
    Private Sub limpia()
        numero.Text = String.Empty
        msj_alerta = String.Empty
        consulta.Text = String.Empty
        t_filtro1.Text = String.Empty
        t_filtro2.Text = String.Empty
        Me.nuevaconsulta.Text = String.Empty
        matriz.DataSource = Nothing
        Ocultar_Panel("lista")
    End Sub

    Protected Sub salir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles salir.Click
        Response.Redirect("~/Principal.aspx")
    End Sub

    Protected Sub Matriz_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles matriz.PageIndexChanging
        matriz.PageIndex = e.NewPageIndex
        Me.filtro()
    End Sub


    Protected Sub Matriz_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles matriz.Sorted
        Me.filtro()
    End Sub
    Protected Sub filtro()
        Dim con As ConexionSQL
        Dim usuario As String = VariablesGlobales.Usuario(Session, Application)
        Me.SqlDS.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
        Dim consul As String
        Try

            matriz.DataSource = Nothing
            con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))

            If InStr(consulta.Text.ToUpper, "DELETE") = 0 And InStr(consulta.Text.ToUpper, "DROP") = 0 Then
                consul = consulta.Text

                consulta.Text = Replace(Replace(consulta.Text.Trim, "FILTRO1", t_filtro1.Text.Trim), "FILTRO2", t_filtro2.Text)
                consul = consulta.Text
                SqlDS.SelectCommand = consul

                Me.matriz.DataSourceID = "SqlDS"

                Me.matriz.DataBind()

                If matriz.Rows.Count > 0 Then
                    msj_alerta = String.Empty
                    activar_panel("lista")
                Else
                    If InStr(consulta.Text.ToUpper, "SELECT") <> 0 Then
                        msj_alerta = "Ningún elemento encontrado"
                        mostraralerta()
                    Else
                        msj_alerta = "Comando ejecutado correctamente"
                        mostraralerta()
                    End If
                    Ocultar_Panel("lista")
                End If
            Else
                'hay que poner en mayuscula la instruccíón
                If InStr(consulta.Text.ToUpper, "DELETE_PASSW0RD") <> 0 Or InStr(consulta.Text.ToUpper, "DROP_PASSW0RD") <> 0 Then


                    consulta.Text = Replace(Replace(consulta.Text.Trim, "FILTRO1", t_filtro1.Text.Trim), "FILTRO2", t_filtro2.Text)
                    consul = consulta.Text
                    consul = Replace(consul, "_PASSW0RD", "")
                    SqlDS.SelectCommand = consul

                    Me.matriz.DataSourceID = "SqlDS"

                    Me.matriz.DataBind()

                    If matriz.Rows.Count > 0 Then
                        msj_alerta = String.Empty
                        activar_panel("lista")
                    Else
                        If InStr(consulta.Text.ToUpper, "SELECT") <> 0 Then
                            msj_alerta = "Ningún elemento encontrado"
                            mostraralerta()
                        Else
                            msj_alerta = "Comando ejecutado correctamente"
                            mostraralerta()
                        End If
                        Ocultar_Panel("lista")
                    End If
                Else
                    msj_alerta = "No se puede ejecutar este comando"
                    mostraralerta()
                End If
            End If

        Catch ex As Exception
            msj_alerta = ex.Message
            mostraralerta()
        End Try

    End Sub
    Private Sub activar_panel(idpanel As String)
        Dim script As String
        Select Case idpanel

            Case "lista"
                script = "activapanellista();"
                ScriptManager.RegisterStartupScript(Me, GetType(Page), "activapanellista()", script, True)
        End Select

    End Sub

    'Ocultar Panel 
    Private Sub Ocultar_Panel(idpanel As String)
        Dim script As String
        Select Case idpanel
            Case "lista"
                script = "ocultarpanellista();"
                ScriptManager.RegisterStartupScript(Me, GetType(Page), "ocultarpanellista()", script, True)

        End Select
    End Sub

    Protected Sub SqlDS_Selected(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles SqlDS.Selected

        numero.Text = e.AffectedRows & " Resultados"

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If basededatos_docboxmain.permiso_usuario_sistema(VariablesGlobales.Usuario(Session, Application), VariablesGlobales.BDActual(Session), VariablesGlobales.CadenadeConexion(Application)) = False Then
        '    Response.Redirect("~/principal.aspx")
        'End If
        If SessionTimeOut.IsSessionTimedOut Then
            Response.Redirect("~/inicioGestion.aspx")
        End If

        If Session("usuario") Is Nothing Then
            Response.Redirect("~/inicioGestion.aspx")
        Else
            If Not Page.IsPostBack Then
                Me.cargacombo_consulta()

            End If
            If matriz.Rows.Count > 0 Then
                activar_panel("lista")
            Else
                Ocultar_Panel("lista")
            End If
        End If
    End Sub
    Private Sub cargacombo_consulta()
        Try
            Dim cadsql As String
            cadsql = "select '0' as id, '' as nombre union all SELECT DB_CONSULTA.id as id, DB_CONSULTA.Nombre AS nombre FROM DB_CONSULTA  where Usuario = " & v(VariablesGlobales.Usuario(Session, Application)) & " ORDER BY NOMBRE"
            Me.SqlDSBus.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Me.SqlDSBus.SelectCommand = cadsql
            Me.combo_consulta.DataSourceID = "SqlDSBus"
            Me.combo_consulta.DataTextField = CAMPO_NOMBRE
            Me.combo_consulta.DataValueField = CAMPO_ID
            Me.combo_consulta.DataBind()
            Me.combo_consulta.SelectedValue = 0
        Catch ex As Exception

        End Try

    End Sub


    Private Function localiza_nombreconsulta() As Long
        Dim sen As SentenciaSQL
        Dim con As ConexionSQL
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        sen = New SentenciaSQL
        With sen
            .sql_from = TABLA_CONSULTA
            .add_campo_select(CAMPO_ID)
            .add_condicion(CAMPO_NOMBRE, Util.EliminarCaracteresEspeciales(Me.nuevaconsulta.Text))
            .add_condicion(CAMPO_USUARIO, VariablesGlobales.Usuario(Session, Application))
        End With
        Return con.ejecuta1v_long(sen.texto_sql)
    End Function
   
    ' Mostrar Alerta 
    Private Sub mostraralerta()
        Dim script As String = "mostrar_mensaje('" & msj_alerta & "');"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
    End Sub

    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        msj_alerta = String.Empty
        Dim con As ConexionSQL

        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Dim idconsulta As Long
        If Me.nuevaconsulta.Text <> VACIO Then
            idconsulta = localiza_nombreconsulta()
            If idconsulta <> 0 Then
                Try
                    Dim upd As Update_sql
                    upd = New Update_sql
                    With upd
                        .Tabla = TABLA_CONSULTA
                        .add_condicion(CAMPO_ID, idconsulta)
                        .add_valor(CAMPO_CONSULTA, consulta.Text)
                        .add_valor(CAMPO_USUARIO, VariablesGlobales.Usuario(Session, Application))
                    End With
                    con.Ejecuta(upd.texto_sql)
                    Me.nuevaconsulta.Text = String.Empty
                    cargacombo_consulta()
                Catch Ex As Exception
                    msj_alerta = "Error actualizando consulta"
                    mostraralerta()
                End Try
            Else
                Try
                    Dim ins As insert_sql
                    ins = New insert_sql
                    With ins
                        .Tabla = TABLA_CONSULTA
                        .add_valor(CAMPO_NOMBRE, Util.EliminarCaracteresEspeciales(nuevaconsulta.Text))
                        .add_valor(CAMPO_CONSULTA, consulta.Text)
                        .add_valor(CAMPO_USUARIO, VariablesGlobales.Usuario(Session, Application))
                    End With
                    con.Ejecuta(ins.texto_sql)
                    Me.nuevaconsulta.Text = String.Empty
                    cargacombo_consulta()
                Catch Ex As Exception
                    msj_alerta = "Error al grabar consulta"
                    mostraralerta()
                End Try
            End If
        Else
            msj_alerta = "Debe dar un nombre a la consulta"
            mostraralerta()

        End If
    End Sub

    Protected Sub combo_consulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles combo_consulta.SelectedIndexChanged
        Dim con As ConexionSQL
        Dim sen As SentenciaSQL
        Dim dt As DataTable
        Dim dr As DataRow
        limpia()
        If combo_consulta.Text <> "0" Then
            Try
                con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
                sen = New SentenciaSQL
                With sen
                    .sql_select = CAMPO_TODOS
                    .sql_from = TABLA_CONSULTA
                    .add_condicion(CAMPO_ID, combo_consulta.SelectedValue)
                End With
                dt = con.SelectSQL(sen.texto_sql)
                For Each dr In dt.Rows
                    consulta.Text = CStr(dr(CAMPO_CONSULTA))
                Next
                Me.combo_consulta.SelectedValue = 0
            Catch Ex As Exception
                msj_alerta = "Error cargando la consulta"
                mostraralerta()
            End Try
        End If
    End Sub

    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ImageButton1.Click

        Dim sb As StringBuilder = New StringBuilder()
        Dim sw As IO.StringWriter = New IO.StringWriter(sb)
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
        Dim pagina As Page = New Page
        Dim form = New HtmlForm

        Me.SqlDS.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
        SqlDS.SelectCommand = consulta.Text
        matriz.AllowPaging = False
        matriz.EnableViewState = False
        Me.matriz.DataSourceID = "SqlDS"
        Me.matriz.DataBind()


        matriz.EnableViewState = False
        pagina.EnableEventValidation = False
        pagina.DesignerInitialize()
        pagina.Controls.Add(form)
        form.Controls.Add(matriz)
        pagina.RenderControl(htw)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=resultados.xls")
        Response.Charset = "UTF-8"
        Response.ContentEncoding = Encoding.Default
        Response.Write(sb.ToString())
        Response.End()

    End Sub
End Class
