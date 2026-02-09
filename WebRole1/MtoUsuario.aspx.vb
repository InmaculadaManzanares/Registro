Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Partial Class MtoUsuario
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
                cargacombo_campos("SELECT  Usuarios.Id,  Usuarios.Email,   Usuarios.Nombre, Empresa , Centro FROM  Usuarios", combo_campo)
                Me.SqlComboempresas.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.SqlComboempresas.SelectCommand = "select id,nombre from empresas"
                Me.cmbempresa.DataSourceID = "SqlComboempresas"
                Me.cmbempresa.DataTextField = "nombre"
                Me.cmbempresa.DataValueField = "id"
                cmbempresa.DataBind()
                Me.SqlCombocentros.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
                Me.SqlCombocentros.SelectCommand = "select id,nombre from centros"
                Me.cmbcentro.DataSourceID = "SqlCombocentros"
                Me.cmbcentro.DataTextField = "nombre"
                Me.cmbcentro.DataValueField = "id"
                cmbcentro.DataBind()
                Dim texttitulo As Label = DirectCast(Page.Master.FindControl("titulo"), Label)
                texttitulo.Text = "Usuarios"

                Ocultar_Panel()
            Else

            End If
        End If




    End Sub


    Protected Sub CargaMatriz(ByVal cond_filtro As String)
        Dim cadena As String
        Me.SqlDS.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
        cadena = "SELECT        Usuarios.Id, Usuarios.email, Usuarios.password, Usuarios.Nombre, empresas.Nombre AS empresa, CASE WHEN verificado = 0 THEN 'No' ELSE 'Si' END AS verificado, 
                         CASE WHEN Administrador = 0 THEN 'No' ELSE 'Si' END AS Administrador, centros.Nombre AS centro
FROM            Usuarios LEFT OUTER JOIN
                         centros ON Usuarios.Centro = centros.Id LEFT OUTER JOIN
                         empresas ON Usuarios.Empresa = empresas.Id"
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
        Dim campo As String
        valor = EliminarCaracteresEspeciales(txt_valorescampo.Text)

        operador = combo_operador.SelectedValue
        If valor <> VACIO Then
            campo = combo_campo.SelectedValue
            If campo = "Empresa" Then campo = " Empresas.nombre "
            If campo = "Centro" Then campo = " centros.nombre "
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
                .Tabla = "usuarios"
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
    Protected Function Guardar(ByVal clave As String, ByVal nombre As String, estado As String) As Boolean
        Dim resul As Boolean

        Dim sen As SentenciaSQL
        Dim con As ConexionSQL
        Dim id As String
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        ocultar_mensaje()
        If txtid.Text = "" Then
            Try
                sen = New SentenciaSQL
                With sen
                    .sql_from = "usuarios"
                    .add_campo_select("id")
                    .add_condicion("email", txtemail.Text)
                End With
                id = con.ejecuta1v_string(sen.texto_sql)

                If id = VACIO Then
                    Dim ins As insert_sql
                    ins = New insert_sql
                    With ins
                        .Tabla = "usuarios"
                        .add_valor(CAMPO_NOMBRE, txtnombre.Text)
                        .add_valor(CAMPO_EMAIL, txtemail.Text)
                        .add_valor(CAMPO_PASSWORD, txtPassword.Text)
                        .add_valor("Empresa", cmbempresa.Text)
                        .add_valor("centro", cmbcentro.Text)
                        Dim adm As String
                        If cmbadmini.Text = "Si" Then adm = "1" Else adm = "0"
                        .add_valor("administrador", adm)
                        Dim ver As String
                        If cmbveri.Text = "Si" Then ver = "1" Else ver = "0"
                        .add_valor("verificado", ver)
                        If txtnotifica.Checked = True Then
                            .add_valor("notificar", 1)
                        Else
                            .add_valor("notificar", 0)
                        End If
                    End With
                    con.Ejecuta(ins.texto_sql)

                    If lunes1.Text <> "" And lunes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "lunes" & "," & v(lunes1.Text) & "," & v(lunes2.Text) & ")")
                    End If
                    If martes1.Text <> "" And martes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "martes" & "," & v(martes1.Text) & "," & v(martes2.Text) & ")")
                    End If
                    If miercoles1.Text <> "" And miercoles2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "miércoles" & "," & v(miercoles1.Text) & "," & v(miercoles2.Text) & ")")
                    End If
                    If jueves1.Text <> "" And jueves2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "jueves" & "," & v(jueves1.Text) & "," & v(jueves2.Text) & ")")
                    End If
                    If viernes1.Text <> "" And viernes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "viernes" & "," & v(viernes1.Text) & "," & v(viernes2.Text) & ")")
                    End If
                    If sabado1.Text <> "" And sabado2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("sábado") & "," & v(sabado1.Text) & "," & v(sabado2.Text) & ")")
                    End If
                    If domingo1.Text <> "" And domingo2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("domingo") & "," & v(domingo1.Text) & "," & v(domingo2.Text) & ")")
                    End If

                    If tlunes1.Text <> "" And tlunes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "lunes tarde" & "," & v(tlunes1.Text) & "," & v(tlunes2.Text) & ")")
                    End If
                    If tmartes1.Text <> "" And tmartes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "martes tarde" & "," & v(tmartes1.Text) & "," & v(tmartes2.Text) & ")")
                    End If
                    If tmiercoles1.Text <> "" And tmiercoles2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "miércoles tarde" & "," & v(tmiercoles1.Text) & "," & v(tmiercoles2.Text) & ")")
                    End If
                    If tjueves1.Text <> "" And tjueves2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "jueves tarde" & "," & v(tjueves1.Text) & "," & v(tjueves2.Text) & ")")
                    End If
                    If tviernes1.Text <> "" And tviernes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "viernes tarde" & "," & v(tviernes1.Text) & "," & v(tviernes2.Text) & ")")
                    End If
                    If tsabado1.Text <> "" And tsabado2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("sábado tarde") & "," & v(tsabado1.Text) & "," & v(tsabado2.Text) & ")")
                    End If
                    If tdomingo1.Text <> "" And tdomingo2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("domingo tarde") & "," & v(tdomingo1.Text) & "," & v(tdomingo2.Text) & ")")
                    End If



                    msj_alerta = "El Usuario se ha guardado correctamente."
                    mostraralerta()

                    activar_panel()
                    activar_panel_lista()
                    resul = True
                Else
                    msj_alerta = "El email de Usuario ya existe"
                    mostraralerta()

                    activar_panel_lista()
                    activar_panel()

                    resul = False
                End If
            Catch ex As Exception

                msj_alerta = "Error insertando usuario:  " & ex.Message
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
                    .Tabla = "usuarios"
                    .add_valor(CAMPO_NOMBRE, txtnombre.Text)
                    .add_valor(CAMPO_EMAIL, txtemail.Text)
                    .add_valor(CAMPO_PASSWORD, txtPassword.Text)
                    .add_valor("Empresa", cmbempresa.Text)
                    .add_valor("centro", cmbcentro.Text)
                    Dim adm As String
                    If cmbadmini.Text = "Si" Then adm = "1" Else adm = "0"
                    .add_valor("administrador", adm)
                    Dim ver As String
                    If cmbveri.Text = "Si" Then ver = "1" Else ver = "0"
                    .add_valor("verificado", ver)
                    If txtnotifica.Checked = True Then
                        .add_valor("notificar", 1)
                    Else
                        .add_valor("notificar", 0)
                    End If
                    .add_condicion("id", txtid.Text)
                End With
                con.Ejecuta("delete from  [UsuariosHorarios] where email =" & v(txtemail.Text))

                If lunes1.Text <> "" And lunes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "lunes" & "," & v(lunes1.Text) & "," & v(lunes2.Text) & ")")
                End If
                If martes1.Text <> "" And martes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "martes" & "," & v(martes1.Text) & "," & v(martes2.Text) & ")")
                End If
                If miercoles1.Text <> "" And miercoles2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "miércoles" & "," & v(miercoles1.Text) & "," & v(miercoles2.Text) & ")")
                End If
                If jueves1.Text <> "" And jueves2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "jueves" & "," & v(jueves1.Text) & "," & v(jueves2.Text) & ")")
                End If
                If viernes1.Text <> "" And viernes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "viernes" & "," & v(viernes1.Text) & "," & v(viernes2.Text) & ")")
                End If
                If sabado1.Text <> "" And sabado2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("sábado") & "," & v(sabado1.Text) & "," & v(sabado2.Text) & ")")
                End If
                If domingo1.Text <> "" And domingo2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("domingo") & "," & v(domingo1.Text) & "," & v(domingo2.Text) & ")")
                End If

                If tlunes1.Text <> "" And tlunes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "lunes tarde" & "," & v(tlunes1.Text) & "," & v(tlunes2.Text) & ")")
                End If
                If tmartes1.Text <> "" And tmartes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "martes tarde" & "," & v(tmartes1.Text) & "," & v(tmartes2.Text) & ")")
                End If
                If tmiercoles1.Text <> "" And tmiercoles2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "miércoles tarde" & "," & v(tmiercoles1.Text) & "," & v(tmiercoles2.Text) & ")")
                End If
                If tjueves1.Text <> "" And tjueves2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "jueves tarde" & "," & v(tjueves1.Text) & "," & v(tjueves2.Text) & ")")
                End If
                If tviernes1.Text <> "" And tviernes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "viernes tarde" & "," & v(tviernes1.Text) & "," & v(tviernes2.Text) & ")")
                End If
                If tsabado1.Text <> "" And tsabado2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("sábado tarde") & "," & v(tsabado1.Text) & "," & v(tsabado2.Text) & ")")
                End If
                If tdomingo1.Text <> "" And tdomingo2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("domingo tarde") & "," & v(tdomingo1.Text) & "," & v(tdomingo2.Text) & ")")
                End If
                con.Ejecuta(upd.texto_sql)
                resul = True
                msj_alerta = "El Usuario se ha guardado correctamente."
                mostraralerta()
                activar_panel()
                activar_panel_lista()
            Catch ex As Exception

                msj_alerta = "Error insertando usuario: " & ex.Message
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
            .sql_from = "usuarios"
            .add_campo_select("*")
            .add_condicion("id", txtid.Text)
        End With
        dt = con.SelectSQL(sen.texto_sql)
        For Each dr In dt.Rows
            txtnombre.Text = CStr(dr(CAMPO_NOMBRE))
            If dr(CAMPO_EMAIL).ToString <> "" Then txtemail.Text = CStr(dr(CAMPO_EMAIL))
            If dr("nombre").ToString <> "" Then txtnombre.Text = CStr(dr("nombre"))
            If dr("password").ToString <> "" Then txtPassword.Text = CStr(dr("password"))
            If dr("Empresa").ToString <> "" Then cmbempresa.Text = CStr(dr("Empresa"))
            If dr("centro").ToString <> "" Then cmbcentro.Text = CStr(dr("centro"))
            If dr("verificado").ToString = "1" Then cmbveri.SelectedValue = "Si" Else cmbveri.SelectedValue = "No"
            If dr("notificar").ToString = "1" Then txtnotifica.Checked = True Else txtnotifica.Checked = False

            If dr("administrador").ToString = "1" Then cmbadmini.SelectedValue = "Si" Else cmbadmini.SelectedValue = "No"

        Next

        sen = New SentenciaSQL
        With sen
            .sql_from = "usuarioshorarios"
            .add_campo_select("*")
            .add_condicion("email", txtemail.Text)
        End With
        dt = con.SelectSQL(sen.texto_sql)
        For Each dr In dt.Rows
            If dr("dia").ToString = "lunes" Then
                lunes1.Text = CStr(dr("horainicio"))
                lunes2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "martes" Then
                martes1.Text = CStr(dr("horainicio"))
                martes2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "miércoles" Then
                miercoles1.Text = CStr(dr("horainicio"))
                miercoles2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "jueves" Then
                jueves1.Text = CStr(dr("horainicio"))
                jueves2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "viernes" Then
                viernes1.Text = CStr(dr("horainicio"))
                viernes2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "sábado" Then
                sabado1.Text = CStr(dr("horainicio"))
                sabado2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "domingo" Then
                domingo1.Text = CStr(dr("horainicio"))
                domingo2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "lunes tarde" Then
                tlunes1.Text = CStr(dr("horainicio"))
                tlunes2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "martes tarde" Then
                tmartes1.Text = CStr(dr("horainicio"))
                tmartes2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "miércoles tarde" Then
                tmiercoles1.Text = CStr(dr("horainicio"))
                tmiercoles2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "jueves tarde" Then
                tjueves1.Text = CStr(dr("horainicio"))
                tjueves2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "viernes tarde" Then
                tviernes1.Text = CStr(dr("horainicio"))
                tviernes2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "sábado tarde" Then
                tsabado1.Text = CStr(dr("horainicio"))
                tsabado2.Text = CStr(dr("horafin"))
            End If
            If dr("dia").ToString = "domingo tarde" Then
                tdomingo1.Text = CStr(dr("horainicio"))
                tdomingo2.Text = CStr(dr("horafin"))
            End If


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
        txtemail.Text = String.Empty
        txtnombre.Text = String.Empty
        txtPassword.Text = String.Empty
        cmbadmini.Text = "No"
        cmbveri.Text = "No"
        lunes1.Text = String.Empty
        lunes2.Text = String.Empty
        martes1.Text = String.Empty
        martes2.Text = String.Empty
        miercoles1.Text = String.Empty
        miercoles2.Text = String.Empty
        jueves1.Text = String.Empty
        jueves2.Text = String.Empty
        viernes1.Text = String.Empty
        viernes2.Text = String.Empty
        sabado1.Text = String.Empty
        sabado2.Text = String.Empty
        domingo1.Text = String.Empty
        domingo2.Text = String.Empty
        tlunes1.Text = String.Empty
        tlunes2.Text = String.Empty
        tmartes1.Text = String.Empty
        tmartes2.Text = String.Empty
        tmiercoles1.Text = String.Empty
        tmiercoles2.Text = String.Empty
        tjueves1.Text = String.Empty
        tjueves2.Text = String.Empty
        tviernes1.Text = String.Empty
        tviernes2.Text = String.Empty
        tsabado1.Text = String.Empty
        tsabado2.Text = String.Empty
        tdomingo1.Text = String.Empty
        tdomingo2.Text = String.Empty
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
        txtid.Enabled = False
    End Sub


    Protected Sub btnsave_Click(sender As Object, e As EventArgs)
        If txtemail.Text <> VACIO Then
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
    Protected Function Guardar(ByVal clave As String) As Boolean
        Dim resul As Boolean

        Dim sen As SentenciaSQL
        Dim con As ConexionSQL
        Dim id As String
        con = New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        ocultar_mensaje()
        If txtid.Text = "" Then
            Try
                sen = New SentenciaSQL
                With sen
                    .sql_from = "usuarios"
                    .add_campo_select("id")
                    .add_condicion("email", txtemail.Text)
                End With
                id = con.ejecuta1v_string(sen.texto_sql)

                If id = VACIO Then
                    Dim ins As insert_sql
                    ins = New insert_sql
                    With ins
                        .Tabla = "usuarios"
                        .add_valor(CAMPO_NOMBRE, txtnombre.Text)
                        .add_valor(CAMPO_EMAIL, txtemail.Text)
                        .add_valor(CAMPO_PASSWORD, txtPassword.Text)
                        .add_valor("Empresa", cmbempresa.Text)
                        .add_valor("centro", cmbcentro.Text)
                        Dim adm As String
                        If cmbadmini.Text = "Si" Then adm = "1" Else adm = "0"
                        .add_valor("administrador", adm)
                        Dim ver As String
                        If cmbveri.Text = "Si" Then ver = "1" Else ver = "0"
                        .add_valor("verificado", ver)
                        If txtnotifica.Checked = True Then
                            .add_valor("notificar", 1)
                        Else
                            .add_valor("notificar", 0)
                        End If
                    End With
                    con.Ejecuta(ins.texto_sql)

                    If lunes1.Text <> "" And lunes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "lunes" & "," & v(lunes1.Text) & "," & v(lunes2.Text) & ")")
                    End If
                    If martes1.Text <> "" And martes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "martes" & "," & v(martes1.Text) & "," & v(martes2.Text) & ")")
                    End If
                    If miercoles1.Text <> "" And miercoles2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "miércoles" & "," & v(miercoles1.Text) & "," & v(miercoles2.Text) & ")")
                    End If
                    If jueves1.Text <> "" And jueves2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "jueves" & "," & v(jueves1.Text) & "," & v(jueves2.Text) & ")")
                    End If
                    If viernes1.Text <> "" And viernes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & "viernes" & "," & v(viernes1.Text) & "," & v(viernes2.Text) & ")")
                    End If
                    If sabado1.Text <> "" And sabado2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("sábado") & "," & v(sabado1.Text) & "," & v(sabado2.Text) & ")")
                    End If
                    If domingo1.Text <> "" And domingo2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("domingo") & "," & v(domingo1.Text) & "," & v(domingo2.Text) & ")")
                    End If
                    If tlunes1.Text <> "" And tlunes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("lunes tarde") & "," & v(tlunes1.Text) & "," & v(tlunes2.Text) & ")")
                    End If
                    If tmartes1.Text <> "" And tmartes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("martes tarde") & "," & v(tmartes1.Text) & "," & v(tmartes2.Text) & ")")
                    End If
                    If tmiercoles1.Text <> "" And tmiercoles2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("miércoles tarde") & "," & v(tmiercoles1.Text) & "," & v(tmiercoles2.Text) & ")")
                    End If
                    If tjueves1.Text <> "" And tjueves2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("jueves tarde") & "," & v(tjueves1.Text) & "," & v(tjueves2.Text) & ")")
                    End If
                    If tviernes1.Text <> "" And tviernes2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("viernes tarde") & "," & v(tviernes1.Text) & "," & v(tviernes2.Text) & ")")
                    End If
                    If tsabado1.Text <> "" And tsabado2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("sábado tarde") & "," & v(tsabado1.Text) & "," & v(tsabado2.Text) & ")")
                    End If
                    If tdomingo1.Text <> "" And tdomingo2.Text <> "" Then
                        con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("domingo tarde") & "," & v(tdomingo1.Text) & "," & v(tdomingo2.Text) & ")")
                    End If

                    msj_alerta = "El Usuario se ha guardado correctamente."
                    mostraralerta()

                    activar_panel()
                    activar_panel_lista()
                    resul = True
                Else
                    msj_alerta = "El email de Usuario ya existe"
                    mostraralerta()

                    activar_panel_lista()
                    activar_panel()

                    resul = False
                End If
            Catch ex As Exception

                msj_alerta = "Error insertando usuario: " & ex.Message
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
                    .Tabla = "usuarios"
                    .add_valor(CAMPO_NOMBRE, txtnombre.Text)
                    .add_valor(CAMPO_EMAIL, txtemail.Text)
                    .add_valor(CAMPO_PASSWORD, txtPassword.Text)
                    .add_valor("Empresa", cmbempresa.Text)
                    .add_valor("centro", cmbcentro.Text)
                    Dim adm As String
                    If cmbadmini.Text = "Si" Then adm = "1" Else adm = "0"
                    .add_valor("administrador", adm)
                    Dim ver As String
                    If cmbveri.Text = "Si" Then ver = "1" Else ver = "0"
                    .add_valor("verificado", ver)
                    If txtnotifica.Checked = True Then
                        .add_valor("notificar", 1)
                    Else
                        .add_valor("notificar", 0)
                    End If

                    .add_condicion("id", txtid.Text)

                End With

                con.Ejecuta(upd.texto_sql)

                con.Ejecuta("delete from  [UsuariosHorarios] where email =" & v(txtemail.Text))

                If lunes1.Text <> "" And lunes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("lunes") & "," & v(lunes1.Text) & "," & v(lunes2.Text) & ")")
                End If
                If martes1.Text <> "" And martes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("martes") & "," & v(martes1.Text) & "," & v(martes2.Text) & ")")
                End If
                If miercoles1.Text <> "" And miercoles2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("miércoles") & "," & v(miercoles1.Text) & "," & v(miercoles2.Text) & ")")
                End If
                If jueves1.Text <> "" And jueves2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("jueves") & "," & v(jueves1.Text) & "," & v(jueves2.Text) & ")")
                End If
                If viernes1.Text <> "" And viernes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("viernes") & "," & v(viernes1.Text) & "," & v(viernes2.Text) & ")")
                End If
                If sabado1.Text <> "" And sabado2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("sábado") & "," & v(sabado1.Text) & "," & v(sabado2.Text) & ")")
                End If
                If domingo1.Text <> "" And domingo2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("domingo") & "," & v(domingo1.Text) & "," & v(domingo2.Text) & ")")
                End If
                If tlunes1.Text <> "" And tlunes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("lunes tarde") & "," & v(tlunes1.Text) & "," & v(tlunes2.Text) & ")")
                End If
                If tmartes1.Text <> "" And tmartes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("martes tarde") & "," & v(tmartes1.Text) & "," & v(tmartes2.Text) & ")")
                End If
                If tmiercoles1.Text <> "" And tmiercoles2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("miércoles tarde") & "," & v(tmiercoles1.Text) & "," & v(tmiercoles2.Text) & ")")
                End If
                If tjueves1.Text <> "" And tjueves2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("jueves tarde") & "," & v(tjueves1.Text) & "," & v(tjueves2.Text) & ")")
                End If
                If tviernes1.Text <> "" And tviernes2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("viernes tarde") & "," & v(tviernes1.Text) & "," & v(tviernes2.Text) & ")")
                End If
                If tsabado1.Text <> "" And tsabado2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("sábado tarde") & "," & v(tsabado1.Text) & "," & v(tsabado2.Text) & ")")
                End If
                If tdomingo1.Text <> "" And tdomingo2.Text <> "" Then
                    con.Ejecuta("INSERT INTO [UsuariosHorarios]([email] ,[dia],[horainicio],[horafin]) VALUES (" &
                            v(txtemail.Text) & "," & v("domingo tarde") & "," & v(tdomingo1.Text) & "," & v(tdomingo2.Text) & ")")
                End If


                resul = True
                msj_alerta = "El Usuario se ha guardado correctamente."
                mostraralerta()
                activar_panel()
                activar_panel_lista()
            Catch ex As Exception

                msj_alerta = "Error insertando usuario: " & ex.Message
                mostraralerta()
                activar_panel_lista()
                activar_panel()
                resul = False
            End Try
        End If
        If resul = True Then borra_campos()

        Return resul
    End Function
    Public Function compruebausuario() As Boolean

        Return True
    End Function

    Protected Sub btncancel_Click(sender As Object, e As EventArgs) Handles btncancel.Click
        activar_panel_lista()
        Ocultar_Panel()
    End Sub

    Private Sub cmbempresa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbempresa.SelectedIndexChanged
        activar_panel_lista()
        activar_panel()
    End Sub

    Private Sub cmbadmini_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbadmini.SelectedIndexChanged
        activar_panel_lista()
        activar_panel()
    End Sub

    Private Sub cmbveri_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbveri.SelectedIndexChanged
        activar_panel_lista()
        activar_panel()
    End Sub

    Private Sub cmbcentro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbcentro.SelectedIndexChanged
        activar_panel_lista()
        activar_panel()
    End Sub

    Private Sub btnemail_Click(sender As Object, e As EventArgs) Handles btnemail.Click
        If txtid.Text <> "" Then
            If enviaremail(txtPassword.Text) Then
                msj_alerta = "Le hemos enviado un correo al usuario con su contraseña"
                Grabar_Accion(0, ACCION_EMAIL_RECUPERAR_CONTRASEÑA, Now())
                mostraralerta()
                activar_panel_lista()
            End If
        Else
            msj_alerta = "Debe crear primero el usuario"
            mostraralerta()
        End If
    End Sub

    Private Function enviaremail(contraseña) As Boolean
        Dim correo As New EnviodeEmail
        Dim valor As Boolean
        Dim asunto As String

        Dim miemail As String = VariablesGlobales.Email_WebCampus(Application)
        asunto = "Usuario y contraseña de AppRegistro"
        Dim urlTemplate As String = Server.MapPath("~/Plantillas/RecuperarContraseña.html")
        Dim urlimagen As String = System.Configuration.ConfigurationManager.AppSettings("url").ToString

        Dim Template As New StringBuilder
        Template.Append(GetHTMLFromAddress(urlTemplate))
        Template.Replace("$USER$", txtemail.Text)
        Template.Replace("$CONTRASE$", contraseña)
        Template.Replace("$URLIMAGEN$", urlimagen)


        If correo.EnviarEmail_html(miemail, txtemail.Text, asunto, (Template.ToString), "", Application) Then
            valor = True
        Else
            msj_alerta = "No se pudo enviar correo, intentelo denuevo"
            mostraralerta()
            valor = False
        End If
        Return valor
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

End Class
