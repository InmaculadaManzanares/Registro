Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Public Class Modificaciones
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
            activarcalendario("txtfecha")
            If Not Page.IsPostBack Then
                borra_campos()
                filtro()
                cargacombo_campos("SELECT   Email, Nombre, inOut AS Entrada/Salida, oficina AS Lugar, Obs FROM  RegistroHorario", combo_campo)
                Me.SqlCombousuario.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.SqlCombousuario.SelectCommand = "select '' as email, '' as nombre union all select email,nombre from usuarios"
                Me.cmbusuario.DataSourceID = "SqlCombousuario"
                Me.cmbusuario.DataTextField = "nombre"
                Me.cmbusuario.DataValueField = "email"
                cmbusuario.DataBind()
                Dim texttitulo As Label = DirectCast(Page.Master.FindControl("titulo"), Label)
                texttitulo.Text = "Registro Altas/Bajas/Modificaciones"

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
        cadena = "SELECT   RegistroHorario.id ,     RegistroHorario.email, Usuarios.Nombre, empresas.Nombre AS empresa,  
                         centros.Nombre AS centro,  RegistroHorario.Fecha,inOut , RegistroHorario.oficina , RegistroHorario.Obs
FROM            Usuarios INNER JOIN
                         RegistroHorario ON Usuarios.email = RegistroHorario.email LEFT OUTER JOIN
                         centros ON Usuarios.Centro = centros.Id LEFT OUTER JOIN
                         empresas ON Usuarios.Empresa = empresas.Id  "
        If cond_filtro <> "" Then
            Me.SqlDS.SelectCommand = cadena & " where " & cond_filtro & sql_ORDERBY & " RegistroHorario.id "
        Else
            Me.SqlDS.SelectCommand = cadena & sql_ORDERBY & " RegistroHorario.FECHA DESC "
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
        Dim campo As String
        valor = EliminarCaracteresEspeciales(txt_valorescampo.Text)

        operador = combo_operador.SelectedValue
        If valor <> VACIO Then
            campo = combo_campo.SelectedValue
            If campo = "Nombre" Then campo = " usuarios.nombre "
            If campo = "Email" Then campo = " RegistroHorario.email "
            If campo = "Id" Then campo = " RegistroHorario.id "
            If campo = "Entrada/Salida" Then campo = "RegistroHorario.inOut"
            If campo = "Lugar" Then campo = "RegistroHorario.oficina"

            If operador.Trim = SQL_LIKE.Trim Then
                cad_filtro = campo & SQL_LIKE & v_like(valor)
            Else
                cad_filtro = campo & operador & v(valor)
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
            If s.Contains(sql_AS) Then
                valor = s.Trim
                p2 = valor.Split(sql_AS)
                combo.Items.Add(p2(2).Trim)
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
                .Tabla = "RegistroHorario"
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
        Dim con As ConexionSQL
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        ocultar_mensaje()
        If txtid.Text = "" Then
            Try

                Dim ins As insert_sql
                ins = New insert_sql
                With ins
                    .Tabla = "RegistroHorario"
                    .add_valor(CAMPO_EMAIL, cmbusuario.Text)
                    .add_valor_expr("fecha", Util.fechasql(txtfecha.Text & " " & txthora.Text, tipofecha.yyyyddMMmmss))
                    .add_valor("inOut", EntradaSalida.Text)
                    .add_valor("oficina", txtlugar.Text)
                    .add_valor("obs", txtobs.Text)
                    .add_valor("ubicacion", "Manual")
                End With
                con.Ejecuta(ins.texto_sql)


                msj_alerta = "El registro se ha guardado correctamente."
                mostraralerta()

                activar_panel()
                activar_panel_lista()
                resul = True

            Catch ex As Exception

                msj_alerta = "Error insertando registro: " & ex.Message
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
                    .Tabla = "RegistroHorario"
                    .add_valor(CAMPO_EMAIL, cmbusuario.Text)
                    .add_valor_expr("fecha", Util.fechasql(txtfecha.Text & " " & txthora.Text, tipofecha.yyyyddMMmmss))
                    .add_valor("inOut", EntradaSalida.Text)
                    .add_valor("oficina", txtlugar.Text)
                    .add_valor("obs", txtobs.Text)
                    .add_condicion("id", txtid.Text)
                End With

                con.Ejecuta(upd.texto_sql)
                resul = True
                msj_alerta = "El registro se ha guardado correctamente."
                mostraralerta()
                activar_panel()
                activar_panel_lista()
            Catch ex As Exception

                msj_alerta = "Error insertando registro: " & ex.Message
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
            .sql_from = "RegistroHorario"
            .add_campo_select("*")
            .add_condicion("id", txtid.Text)
        End With
        dt = con.SelectSQL(sen.texto_sql)
        For Each dr In dt.Rows

            If dr(CAMPO_EMAIL).ToString <> "" Then cmbusuario.Text = CStr(dr(CAMPO_EMAIL))
            If dr("inOut").ToString <> "" Then EntradaSalida.Text = CStr(dr("inOut"))
            If dr("fecha").ToString <> "" Then

                Dim hora As String() = CStr(dr("fecha")).ToString.Split(" ")
                txtfecha.Text = hora(0)
                If hora.Length = 2 Then txthora.Text = hora(1)
            End If
            If dr("oficina").ToString <> "" Then txtlugar.Text = CStr(dr("oficina"))
            If dr("obs").ToString <> "" Then txtobs.Text = CStr(dr("obs"))

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
        cmbusuario.Text = String.Empty
        txtlugar.Text = String.Empty
        txtfecha.Text = String.Empty
        txtobs.Text = String.Empty
        txthora.Text = String.Empty
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
        ' limpiar_campo_email()
        ' borra_campos()

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
        txtlugar.Text = "Oficina"
        txtid.Enabled = False
    End Sub


    Protected Sub btnsave_Click(sender As Object, e As EventArgs)
        If txtlugar.Text <> VACIO And cmbusuario.Text <> VACIO And IsDate(txtfecha.Text) And IsDate(txthora.Text) Then
            If IsDate(txthora.Text) = False Then
                msj_alerta = "hora incorrecta."
                mostraralerta()
                activar_panel()
            End If
            If Guardar(txtid.Text) Then
                filtro()
                UpdatePanel1.Update()
                activar_panel_lista()
                Ocultar_Panel()
            Else
                activar_panel()
            End If
        Else
            msj_alerta = "Compruebe los datos."
            mostraralerta()
            activar_panel()
        End If


    End Sub



    Protected Sub btncancel_Click(sender As Object, e As EventArgs) Handles btncancel.Click
        activar_panel_lista()
        Ocultar_Panel()
    End Sub

    Private Sub cmbusuario_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbusuario.SelectedIndexChanged
        activar_panel_lista()
        activar_panel()
    End Sub

    Private Sub EntradaSalida_SelectedIndexChanged(sender As Object, e As EventArgs) Handles EntradaSalida.SelectedIndexChanged
        activar_panel_lista()
        activar_panel()
    End Sub
End Class