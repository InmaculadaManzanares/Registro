Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Partial Class MtoEmpresas
    Inherits System.Web.UI.Page

    ' Variable global.
    Dim msj_alerta As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        If SessionTimeOut.IsSessionTimedOut Then
            Response.Redirect("~/InicioGestion.aspx")
        End If

        If Session("usuario") Is Nothing Then
            Response.Redirect("~/InicioGestion.aspx")
        Else
            If Not Page.IsPostBack Then
                borra_campos()
                filtro()
                cargacombo_campos("SELECT  Id,Nombre  FROM  empresas", combo_campo)
                Dim texttitulo As Label = DirectCast(Page.Master.FindControl("titulo"), Label)
                texttitulo.Text = "Empresas"

                Ocultar_Panel()
            Else

            End If
        End If




    End Sub


    Protected Sub CargaMatriz(ByVal cond_filtro As String)
        Dim cadena As String
        Me.SqlDS.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
        cadena = "SELECT  Id,  Nombre FROM  Empresas"
        If cond_filtro <> "" Then
            Me.SqlDS.SelectCommand = cadena & SQL_WHERE & cond_filtro & sql_ORDERBY & CAMPO_NOMBRE
        Else
            Me.SqlDS.SelectCommand = cadena & sql_ORDERBY & CAMPO_NOMBRE
        End If
        Me.matriz.DataSourceID = "SqlDS"
        Me.matriz.DataBind()
    End Sub
    Protected Sub Matriz_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles matriz.PageIndexChanging
        matriz.PageIndex = e.NewPageIndex
        Me.filtro()
    End Sub
    Protected Sub Matriz_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles matriz.Sorted
        Me.filtro()
    End Sub
    Public Sub filtro()
        Dim valor As String = VACIO
        Dim cad_filtro As String = VACIO
        Dim operador As String
        valor = EliminarCaracteresEspeciales(txt_valorescampo.Text)

        operador = combo_operador.SelectedValue
        If valor <> VACIO Then
            If operador.Trim = SQL_LIKE.Trim Then
                cad_filtro = combo_campo.SelectedValue & SQL_LIKE & v_like(valor)
            Else
                cad_filtro = combo_campo.SelectedValue & operador & v(valor)
            End If
        End If
        CargaMatriz(cad_filtro)
    End Sub
    Private Sub cargacombo_campos(ByVal sentencia As String, ByVal combo As DropDownList)
        Dim cadsql As String
        Dim posini, posfin As Integer
        Dim p1(), p2() As String
        Dim valor As String

        posini = Len(SQL_SELECT)
        posfin = InStr(sentencia, SQL_FROM)
        cadsql = Mid(sentencia, posini, posfin - posini)
        p1 = cadsql.Split(COMA)
        For Each s As String In p1
            If s.ToUpper.Contains(sql_AS) Then
                valor = s.ToUpper.Trim
                p2 = valor.Split(sql_AS)
                combo.Items.Add(p2(0).Trim)
            Else
                combo.Items.Add(s.Trim)
            End If
        Next
    End Sub


    Protected Sub Borrar(ByVal clave As String)
        Dim del As delete_sql
        Dim con As ConexionSQL
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Try
            'usuario
            del = New delete_sql
            With del
                .Tabla = "empresas"
                .add_condicion("ID", clave)
            End With
            con.Ejecuta(del.texto_sql)
            del = New delete_sql
        Catch ex As Exception
        Finally
            If Not con Is Nothing Then
                con.CerrarConexion()
                con = Nothing
            End If
        End Try
    End Sub
    Private Sub ocultar_mensaje()
        Dim script As String = "ocultar_mensaje();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
    End Sub

    Private Sub combo_campo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles combo_campo.SelectedIndexChanged
        activar_panel_lista()
        Ocultar_Panel()

    End Sub

    Private Sub combo_operador_SelectedIndexChanged(sender As Object, e As EventArgs) Handles combo_operador.SelectedIndexChanged
        activar_panel_lista()
        Ocultar_Panel()
    End Sub
    Protected Function Guardar(ByVal clave As String) As Boolean
        Dim resul As Boolean

        Dim sen As SentenciaSQL
        Dim con As ConexionSQL
        Dim nombre As String
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        ocultar_mensaje()
        If txtid.Text = "" Then
            Try
                sen = New SentenciaSQL
                With sen
                    .sql_from = "empresas"
                    .add_campo_select("nombre")
                    .add_condicion("nombre", txtnombre.Text)
                End With
                nombre = con.ejecuta1v_string(sen.texto_sql)

                If nombre = VACIO Then
                    Dim ins As insert_sql
                    ins = New insert_sql
                    With ins
                        .Tabla = "empresas"
                        .add_valor(CAMPO_NOMBRE, txtnombre.Text)
                    End With
                    con.Ejecuta(ins.texto_sql)
                    msj_alerta = "La empresa se ha guardado correctamente."
                    mostraralerta()

                    activar_panel()
                    activar_panel_lista()
                    resul = True
                Else
                    msj_alerta = "La empresa ya existe"
                    mostraralerta()

                    activar_panel_lista()
                    activar_panel()

                    resul = False
                End If
            Catch ex As Exception

                msj_alerta = "Error insertando empresa: " & ex.Message
                mostraralerta()
                resul = False
            Finally
                If Not con Is Nothing Then
                    con.CerrarConexion()
                    con = Nothing
                End If
            End Try

        Else
            Dim upd As Update_sql
            Try
                upd = New Update_sql
                With upd
                    .Tabla = "empresas"
                    .add_valor(CAMPO_NOMBRE, txtnombre.Text)
                    .add_condicion("id", txtid.Text)
                End With

                con.Ejecuta(upd.texto_sql)
                resul = True
                msj_alerta = "La empresa se ha guardado correctamente."
                mostraralerta()
                activar_panel()
                activar_panel_lista()
            Catch ex As Exception

                msj_alerta = "Error insertando empresa: " & ex.Message
                mostraralerta()
                activar_panel_lista()
                activar_panel()
                resul = False
            End Try
        End If
        If resul = True Then borra_campos()

        Return resul
    End Function
    ' Mostrar Alerta 
    Private Sub mostraralerta()
        'Dim script As String = "mostrar_mensaje();"

        Dim script As String = "mostrar_mensaje('" & msj_alerta & "');"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
        msj_alerta = String.Empty
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
                .Tabla = "Registro"
                .add_valor("detalle", detail)
                .add_valor("accion", accion)
                .add_valor_expr("fecha", Util.fechasql(fechaahora, tipofecha.yyyyddMMmmss))
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
    '----- funcion activar panel
    'Apertura Panel-Detalle
    Private Sub activar_panel()
        Dim script As String = "activapanel();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "activapanel()", script, True)
    End Sub
    'Apertura Panel-Detalle
    Private Sub activar_panel_lista()
        Dim script As String = "activapanellista();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "activapanellista()", script, True)
    End Sub

    'Ocultar Panel Detalle
    Private Sub Ocultar_Panel()
        Dim script As String = "ocultarpanel();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "ocultarpanel()", script, True)
    End Sub
    Protected Sub btnEdit_Click(sender As Object, e As EventArgs)
        Dim btnEdit As LinkButton
        Dim row As GridViewRow
        Dim sen As SentenciaSQL
        Dim con As ConexionSQL
        Dim dt As DataTable
        Dim dr As DataRow
        borra_campos()
        ocultar_mensaje()
        btnEdit = DirectCast(sender, LinkButton)
        row = DirectCast(btnEdit.NamingContainer, GridViewRow)

        txtid.Text = Scrub(row.Cells(1).Text)

        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))

        sen = New SentenciaSQL
        With sen
            .sql_from = "empresas"
            .add_campo_select("*")
            .add_condicion("id", txtid.Text)
        End With
        dt = con.SelectSQL(sen.texto_sql)
        For Each dr In dt.Rows
            txtnombre.Text = CStr(dr(CAMPO_NOMBRE))

        Next

        txtid.Enabled = False

        UpdatePanel1.Update()
        activar_panel()

        ' Pintar la fila seleccionada.
        Dim row1 As GridViewRow
        Dim actual As String = row.RowIndex

        For a = 0 To matriz.Rows.Count - 1
            row1 = Me.matriz.Rows(a)

            If a = actual Then
                row1.ControlStyle.BackColor = ColorTranslator.FromHtml("#F0E6EF")
            Else
                row1.ControlStyle.BackColor = ColorTranslator.FromHtml("#FFFFFF")
            End If

        Next
        If Not con Is Nothing Then
            con.CerrarConexion()
            con = Nothing
        End If
    End Sub
    Protected Sub btnDelete_Click(sender As Object, e As EventArgs)
        ocultar_mensaje()
        Dim btnEdit As LinkButton
        Dim row As GridViewRow
        btnEdit = DirectCast(sender, LinkButton)
        row = DirectCast(btnEdit.NamingContainer, GridViewRow)
        Dim nmaqui As String = Scrub(row.Cells(1).Text)
        Borrar(nmaqui)
        filtro()
        UpdatePanel1.Update()
        activar_panel()
        activar_panel_lista()
    End Sub


    '----- funciones que deben ser globales MOVER

    Protected Function Scrub(ByVal text As String) As String
        Return text.Replace("&nbsp;", "")
    End Function


    Private Sub borra_campos()
        txtid.Text = String.Empty
        txtnombre.Text = String.Empty

    End Sub

    Protected Sub Filtrar_Click(sender As Object, e As EventArgs) Handles Filtrar.Click
        ocultar_mensaje()
        borra_campos()
        filtro()
        activar_panel_lista()
        Ocultar_Panel()
    End Sub



    Protected Sub Limpiar_Click(sender As Object, e As EventArgs) Handles Limpiar.Click
        ocultar_mensaje()
        txt_valorescampo.Text = ""
        borra_campos()
        filtro()
        activar_panel_lista()
        Ocultar_Panel()

    End Sub


    ' Limpiar Campos 
    Private Sub limpiar_campo_email()
        Dim script As String = "limpiar_campos();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
        'ScriptManager.RegisterClientScriptBlock(Me, GetType(Page), "script", script, True)
    End Sub


    Protected Sub nuevo_Click(sender As Object, e As EventArgs) Handles nuevo.Click
        ocultar_mensaje()
        borra_campos()
        UpdatePanel1.Update()
        limpiar_campo_email()
        activar_panel()
        txtid.Enabled = False
    End Sub


    Protected Sub btnsave_Click(sender As Object, e As EventArgs)
        If txtnombre.Text <> VACIO Then
            If compruebausuario() Then
                If Guardar(txtid.Text) Then
                    filtro()
                    UpdatePanel1.Update()
                    activar_panel_lista()
                    Ocultar_Panel()
                Else
                    activar_panel()

                End If
            Else
                activar_panel()
            End If
        Else

        End If
    End Sub

    Public Function compruebausuario() As Boolean

        Return True
    End Function

    Protected Sub btncancel_Click(sender As Object, e As EventArgs) Handles btncancel.Click
        activar_panel_lista()
        Ocultar_Panel()
    End Sub




End Class
