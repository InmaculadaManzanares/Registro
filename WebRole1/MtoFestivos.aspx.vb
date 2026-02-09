Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Partial Class MtoFestivos
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
            Dim script1 As String = Util.CambiarColorControlesScript(Session("colorprimario"), Session("colorsecundario"))
            Page.ClientScript.RegisterStartupScript(Me.GetType(), "ReplaceColorsScript1", script1)
            activarcalendario("txtfecha")
            If Not Page.IsPostBack Then
                borra_campos()


                Me.SqlDSestado.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.SqlDSestado.SelectCommand = "SELECT  year(Mydate) as Año from MyDates group by  year(Mydate) order by year(Mydate)"
                Me.valor.DataSourceID = "SqlDSestado"
                Me.valor.DataTextField = "Año"
                Me.valor.DataValueField = "Año"
                valor.DataBind()
                valor.Text = Year(Now)
                Dim texttitulo As Label = DirectCast(Page.Master.FindControl("titulo"), Label)
                texttitulo.Text = "Festivos"
                filtro()
                Ocultar_Panel()
            Else

            End If
        End If




    End Sub

    Private Sub activarcalendario(ByVal txt As String)
        Dim script As String = "activar_calendario('" & "#" & "ContentPlaceHolder1_" & txt & "');"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), script, script, True)
    End Sub
    Protected Sub CargaMatriz(ByVal cond_filtro As String)
        Dim cadena As String
        Me.SqlDS.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
        cadena = "SELECT  year(Mydate) as Año, format(Mydate,'dd/MM/yyyy')  as Dia  FROM  MyDates where festivos = 'festivo' and  year(Mydate) = " & valor.Text & " order by Mydate"
        Me.SqlDS.SelectCommand = cadena
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

        CargaMatriz("")
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
    Private Sub cargavalor(ByVal sentencia As String, ByVal combo As DropDownList)
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
                .Tabla = "MyDates"
                .add_condicion("Mydate", clave)
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
                    .sql_from = "MyDates"
                    .add_campo_select("festivos")
                    .add_condicion("mydate", txtfecha.Text)
                    .add_condicion("festivos", "Festivo")
                End With
                nombre = con.ejecuta1v_string(sen.texto_sql)

                If nombre = VACIO Then
                    Dim ins As insert_sql
                    ins = New insert_sql
                    With ins
                        .Tabla = "MyDates"
                        .add_valor("mydate", txtfecha.Text)
                        .add_valor("festivos", "Festivo")
                    End With
                    con.Ejecuta(ins.texto_sql)
                    msj_alerta = "El festivo se ha guardado correctamente."
                    mostraralerta()

                    activar_panel()
                    activar_panel_lista()
                    resul = True
                Else
                    msj_alerta = "El festivo ya existe"
                    mostraralerta()

                    activar_panel_lista()
                    activar_panel()

                    resul = False
                End If
            Catch ex As Exception

                msj_alerta = "Error insertando festivo: " & ex.Message
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
                    .Tabla = "MyDates"
                    .add_valor(CAMPO_NOMBRE, txtfecha.Text)
                    .add_condicion("id", txtid.Text)
                End With

                con.Ejecuta(upd.texto_sql)
                resul = True
                msj_alerta = "El festivo se ha guardado correctamente."
                mostraralerta()
                activar_panel()
                activar_panel_lista()
            Catch ex As Exception

                msj_alerta = "Error insertando festivo: " & ex.Message
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

        borra_campos()
        ocultar_mensaje()
        btnEdit = DirectCast(sender, LinkButton)
        row = DirectCast(btnEdit.NamingContainer, GridViewRow)

        txtid.Text = Scrub(row.Cells(1).Text)
        txtfecha.Text = Scrub(row.Cells(2).Text)
        txtid.Enabled = False
        txtfecha.Enabled = False
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

    End Sub
    Protected Sub btnDelete_Click(sender As Object, e As EventArgs)
        ocultar_mensaje()
        Dim btnEdit As LinkButton
        Dim row As GridViewRow
        btnEdit = DirectCast(sender, LinkButton)
        row = DirectCast(btnEdit.NamingContainer, GridViewRow)
        Dim nmaqui As String = Scrub(row.Cells(2).Text)
        Borrar(nmaqui)
        filtro()
        UpdatePanel1.Update()


        activar_panel_lista()
        Ocultar_Panel()
    End Sub


    '----- funciones que deben ser globales MOVER

    Protected Function Scrub(ByVal text As String) As String
        Return text.Replace("&nbsp;", "")
    End Function


    Private Sub borra_campos()
        txtid.Text = String.Empty
        'txtfecha.Text = String.Empty

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
        txtfecha.Enabled = True
    End Sub


    Protected Sub btnsave_Click(sender As Object, e As EventArgs)
        If txtfecha.Enabled = False Then
            activar_panel_lista()
            Ocultar_Panel()
            Exit Sub
        End If
        If txtfecha.Text <> VACIO Then
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
