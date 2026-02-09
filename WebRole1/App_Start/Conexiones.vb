Option Strict On
Option Explicit On 
Imports System.Text.RegularExpressions
Imports system.IO.Path
Imports System.IO.File
Imports System.Data
#If Not compact_framework Then
Imports System.Data.OleDb




Public Class ConexionTexto
    Inherits DefinicionFuncionTexto

    Private Const ETIQUETA_ENTERO As String = "ENTERO" & Util.PARENTESISI
    Private Const ETIQUETA_FLOTANTE As String = "FLOTANTE" & Util.PARENTESISI
    Private Const ETIQUETA_CADENA As String = "CADENA" & Util.PARENTESISI
    Private Const ETIQUETA_FECHA As String = "FECHA" & Util.PARENTESISI

    Private Const ETIQUETA_RUTA As String = "RUTA" & Util.PARENTESISI
    Private Const ETIQUETA_FIN_REGISTRO As String = "FINREG" & Util.PARENTESISI
    Private Const ETIQUETA_CREARVACIO As String = "CREARVACIO" & Util.PARENTESISI
    Private Const ETIQUETA_FICHEROTESTIGO As String = "TESTIGO" & Util.PARENTESISI
    Private Const ETIQUETA_DIRHISTORICO As String = "DIRHISTORICO" & Util.PARENTESISI

    Private tokens() As String = {ETIQUETA_ENTERO, ETIQUETA_CADENA, ETIQUETA_FLOTANTE, ETIQUETA_FECHA, ETIQUETA_RUTA, ETIQUETA_FIN_REGISTRO, ETIQUETA_CREARVACIO, ETIQUETA_FICHEROTESTIGO, ETIQUETA_DIRHISTORICO}



    Private blk_caracteresfinalregistro() As Char
    Private buffer_cfr() As Char
    Private lon_caracteresfinalregistro As Integer

    Private CrearVacio As Boolean
    Private ruta As String = VACIO

    Protected ficherotestigo As Boolean
    Protected exttestigo As String
    
    Protected PasarAHistorico As Boolean
    Protected rutHist As String

    'debemos tener una coleccion de campos


	Const ptr_NUMERO As String = "[0-9]"
    Const ptr_NUMERO2 As String = ptr_NUMERO & ptr_NUMERO
    Const PATRON_IDUNICO As String = "\$IDUNICO" & ptr_NUMERO2 & "\$"   

    Public Property extensionFicheroTestigo() As String
        Get
            Return exttestigo
        End Get
        Set(ByVal value As String)
            If value.Contains("\") Or value.Contains(".") Or value.Contains(":") Then
                Throw New Exception("Extensión del fichero testigo incorrecta, (p.e se esperaba 'OK', 'BAk') se ha encontrado: " & value)
            End If
            exttestigo = value
        End Set
    End Property
   
    Public Sub EstablecerFinaldeRegistro(ByRef cadHex As String)
        Dim i As Integer

        Me.lon_caracteresfinalregistro = cadHex.Length \ 2
        ReDim blk_caracteresfinalregistro(lon_caracteresfinalregistro)
        ReDim buffer_cfr(lon_caracteresfinalregistro)

        For i = 1 To Len(cadHex) - 1 Step 2
            blk_caracteresfinalregistro((i - 1) \ 2) = Chr(Hex2Dec(Mid(cadHex, i, 1), Mid(cadHex, i + 1, 1)))
        Next i

    End Sub
    Public Sub EscribirFinaldeRegistro(ByVal sw As System.IO.StreamWriter)
        If Me.lon_caracteresfinalregistro > 0 Then
            Try
                sw.Write(Me.blk_caracteresfinalregistro, 0, lon_caracteresfinalregistro)
            Catch ex As Exception
                Throw New Exception("Escribiendo Caracter(es) final de registro " & ex.Message)
            End Try
        End If
    End Sub

    Public Sub LeerFinalderegistro(ByVal sr As System.IO.StreamReader)

        If Me.lon_caracteresfinalregistro > 0 Then
            Try
                If sr.ReadBlock(Me.buffer_cfr, 0, Me.lon_caracteresfinalregistro) < Me.lon_caracteresfinalregistro Then
                    Throw New Exception("Encontrado final de fichero (se esperaban caracteres final de registro)")
                End If

                For i As Integer = 0 To Me.lon_caracteresfinalregistro - 1
                    If Me.buffer_cfr(i) <> Me.blk_caracteresfinalregistro(i) Then
                        Throw New Exception("Se esperaba caracter " & Asc(Me.blk_caracteresfinalregistro).ToString & " y se encuentra " & Asc(Me.buffer_cfr(i)).ToString)
                    End If
                Next
            Catch ex As Exception
                Throw New Exception("Leyendo Final de registro " & ex.Message)
            End Try
        End If

    End Sub

    Dim campos As Coleccion_Campostxt

    Public Sub New(ByVal s As String, ByVal constantes As Coleccion_Constantes)
        'en s viene la estructura del fichero
        Dim col As ColecciondeTokens
        Dim tn As Token_Nodo
        Dim i, il As Integer
        Dim contenido As String

        campos = New Coleccion_Campostxt
        Me.ficherotestigo = False

        col = Parser_TPR.SplitPrefijoMultiple(tokens, s)

        il = col.Count - 1

        For i = 0 To il
            Try
                tn = col.Item(i)
                contenido = Me.QuitaParentesisFinal(tn.contenido)
                Select Case tn.TOKEN
                    Case ETIQUETA_FIN_REGISTRO
                        Me.EstablecerFinaldeRegistro(Parser_TPR.ExtraeParametroHexadecimal(contenido))
                    Case ETIQUETA_RUTA
                        ruta = Parser_TPR.ExtraeParametroCadena(contenido, constantes)
                    Case ETIQUETA_ENTERO
                        campos.Add(New campotxtEntero(contenido, constantes))
                    Case ETIQUETA_FLOTANTE
                        campos.Add(New campotxtFlotante(contenido, constantes))
                    Case ETIQUETA_CADENA
                        campos.Add(New campotxtCadena(contenido, constantes))
                    Case ETIQUETA_FECHA
                        campos.Add(New campotxtFecha(contenido, constantes))
                    Case ETIQUETA_CREARVACIO
                        CrearVacio = Parser_TPR.ExtraeParametrobooleano(contenido)
                    Case ETIQUETA_FICHEROTESTIGO
                        Me.ficherotestigo = True
                        Me.extensionFicheroTestigo = Parser_TPR.ExtraeParametroCadena(contenido, constantes)
                    Case ETIQUETA_DIRHISTORICO
                        Me.PasarAHistorico = True
                        Me.rutHist = Parser_TPR.ExtraeParametroCadena(contenido, constantes)
                End Select
            Catch ex As Exception
                Throw New Exception("Definición Conexion Texto: " & ex.Message)
            End Try
        Next
    End Sub

    

    Public Function ReemplazaVariables(ByRef cadena As String, ByVal e As ColecciondeString) As Boolean
        Dim cambiosrealizados As Boolean = False
        Dim i, il As Integer

        il = e.Count - 1

        For i = 0 To il
            If cadena.Contains(e.Indice(i)) Then
                cadena = cadena.Replace(e.Indice(i), e.Item(i))
                cambiosrealizados = True
            End If
        Next


        Return cambiosrealizados

    End Function

    Public Sub Escribe(ByVal dt As DataTable, ByVal cm As coleccion_mapeocampos)
        Dim dr As DataRow
        Dim i, il As Integer
        Dim fichero As System.IO.StreamWriter
        Dim msj as string = string.empty

        Dim nombretemporal As String

        Try
            il = campos.Count - 1

            Util.CreaDirectorioTemporalAplicacion("TEMP_ESP_TPR")
            nombretemporal = Util.CreaFicherotemporal("tempTPR")

            If cm.ComprobarMapeo(campos, dt, msj) Then
                fichero = New System.IO.StreamWriter(nombretemporal, False, System.Text.Encoding.ASCII)
                Try
                    For Each dr In dt.Rows
                        For i = 0 To il
                            With campos.Item(i)
                                .Escribir(fichero, dr.Item(cm.Mapea(.Nombrecampo)))
                            End With
                        Next
                        Me.EscribirFinaldeRegistro(fichero)
                    Next
                Finally
                    fichero.Close()
                End Try


                'ahora tenemos que copiarlo a la ruta especificada
                If dt.Rows.Count > 0 Or Me.CrearVacio Then
                    MueveaDirectorioDestino(nombretemporal, ruta, Me.ficherotestigo, Me.extensionFicheroTestigo)
                Else
                    Util.BorraficheroSiExiste(nombretemporal)
                End If
            Else
                Throw New Exception("Mapeo de Campos: " & msj)
            End If
        Catch ex As Exception
            Throw New Exception("Error al Escribir Fichero " & ex.Message)
        End Try
    End Sub
    Public Sub Lee(ByVal con As ConexionSQL, ByVal nombretabla As String, ByVal mapeos As Coleccion_mapeocampos)
        Dim c As ColecciondeString
        Dim f As String
        Dim i, il As Integer
        Dim existeunatransaccionglobal As Boolean
        Dim ficherosmovidos As New ColecciondeString
        Dim origen, destino As String

        existeunatransaccionglobal = con.Entransaccion

        c = BuscaFicheroParaLectura(Me.ruta, Me.ficherotestigo, Me.extensionFicheroTestigo)

        il = c.Count - 1
        
        Try
            If Not existeunatransaccionglobal Then
                con.ComienzaTransaccion()
            End If
            For i = 0 To il
                f = c.Item(i)
                Me.Lee(con, nombretabla, mapeos, f)
                If Me.PasarAHistorico Then
                    If ficherotestigo Then
                        origen = GetDirectoryName(f) & System.IO.Path.DirectorySeparatorChar & NombreArchivoTestigo(f, Me.extensionFicheroTestigo)
                        destino = PasaAHistorico(origen)
                        ficherosmovidos.Add(origen, destino)
                    End If
                    destino = PasaAHistorico(f)
                    ficherosmovidos.Add(f, destino)
                End If
            Next
            If Not existeunatransaccionglobal Then
                con.TerminaTransaccion(True)
            End If
        Catch ex As Exception
            If Not existeunatransaccionglobal Then
                con.TerminaTransaccion(False)
            End If
            'Hay que mover los archivos que estaban en el historico
            If Me.PasarAHistorico Then
                il = ficherosmovidos.Count - 1
                For i = 0 To il
                    System.IO.File.Move(ficherosmovidos.Item(i), ficherosmovidos.Indice(i))
                Next
            End If
            Throw (New Exception("Leyendo Fichero Texto: " & ex.Message))
        End Try

    End Sub

    Private Function PasaAHistorico(ByVal f As String) As String
        'devuelve la ruta y el nuevo nombre en el historico
        Dim s as string = string.empty
        Try
            s = ConcatenaDir(Me.rutHist, GetFileNameWithoutExtension(f) & "-" & Now.ToString("yyyyMMddHHmmssfff") & System.IO.Path.GetExtension(f))

            System.IO.File.Move(f, s)
            Return s
        Catch ex As Exception
            Throw New Exception("Pasando a Historico " & f & ".Destino:" & s)
        End Try
    End Function
    Private Sub Lee(ByVal con As ConexionSQL, ByVal nombretabla As String, ByVal mapeos As Coleccion_mapeocampos, ByVal nombreunitario As String)
        'A partir de la definición del fichero
        'extraemos los datos y los guardamos en una tabla 
        '¿como vinculamos los campos de texto a los campos de la tabla?
        'La tabla puede tener mas campos de los que vienen en el fichero
        'El fichero puede tener mas campos de los que necesitamos
        'Necesitamos un MAPEO.  Nombre Campo Texto <-> Nombre Campo tabla

        'MAPEOCONEXION::NombreMapeo:: M(nombrecampofichero,nombrecampotabla)
        '                             M(nombrecampofichero, nombrecampotabla)

        'LEER::conexionficherotexto::conexion2.tabla4
        '¿Cómo le especifico que utilice un mapeo?

        'LEER::conexionficherotexto.Mapeo::conexion2.tabla4

        'Leemos linea a linea y añadimos datos en la tabla SQL

        Dim i, il As Integer
        Dim sr As System.IO.StreamReader = Nothing
        Dim ins As New insert_sql
        Dim c As campotxt
        Dim c_registros, c_intraregistro, c_total As Long
        Dim encontrado As Boolean
        Dim contenido As String
        Dim nombrecampotabla as string = string.empty

        ' Dim dt As DataTable




        il = campos.Count - 1

        Try
            'dt = con.SelectSQL(Util.SQL_SELECT & Util.SQL_TOP & "0" & Util.SQL_FROM & nombretabla)
            'Try
            '    mapeos.ComprobarMapeo(campos, dt)
            'Finally
            '    con.Libera(dt)
            'End Try

            sr = New System.IO.StreamReader(nombreunitario, System.Text.Encoding.ASCII)
            c_registros = 0
            'Tenemos que hacer un bucle para leer todos los registros 


            While Not sr.EndOfStream
                c_intraregistro = 0
                With ins
                    .Limpia()
                    .Tabla = nombretabla
                End With
                Try
                    For i = 0 To il
                        c = campos.Item(i)
                        contenido = c.Leer(sr)
                        With mapeos

                            If .EsAutoMapeo Then
                                encontrado = True
                                nombrecampotabla = c.Nombrecampo
                            Else
                                If .Existe(c.Nombrecampo) Then
                                    encontrado = True
                                    nombrecampotabla = .Item(c.Nombrecampo)
                                Else
                                    encontrado = False
                                End If
                            End If
                            If encontrado Then
                                ins.add_valor_expr(nombrecampotabla, contenido)
                                c_intraregistro += c.longitudcampo
                                c_total += c.longitudcampo
                            End If
                        End With

                    Next
                    'Leemos el final de registro
                    LeerFinalderegistro(sr)

                Catch ex As Exception
                    Throw New Exception("Error leyendo Nº Registro:" & c_registros.ToString & " posición: " & c_intraregistro.ToString & " Posición global: " & c_total & " " & ex.Message)
                End Try
                Try
                    con.Ejecuta(ins.texto_sql)
                Catch ex As Exception
                    Throw New Exception("Error guardando registro leido " & " Nº Registro:" & c_registros.ToString & " " & ex.Message)
                End Try
                c_registros += 1
                c_total += Me.lon_caracteresfinalregistro
            End While
        Catch ex As Exception
            Throw New Exception("Error al leer fichero: " & nombreunitario & "." & ex.Message)
        Finally
            sr.Close()
        End Try

    End Sub
End Class



Public MustInherit Class Conexion_x_SQL
    'interfaz comun para SQLServer, MySQL, ...
    Protected entrasaccion As Boolean
    Protected cadenaconexion As String ' guardamos la cadena de conexino por si hace falta clonar esta conexion
    Protected idtabla As Decimal
    Friend dbDataset As DataSet
    Public Const NOMBRE_TABLA_EXPORTACION_ESQUEMA_SQL As String = "Table"

    Public Sub New(ByVal cadconex As String, ByVal dbDataSetExistente As Data.DataSet)
        cadenaconexion = cadconex
        dbDataset = dbDataSetExistente
        entrasaccion = False
        Me.idtabla = 0
    End Sub
    Public Sub New(ByVal cadconex As String)
        cadenaconexion = cadconex
        dbDataset = New Data.DataSet
        entrasaccion = False
        Me.idtabla = 0
    End Sub

    Public Sub New(ByVal usuario As String, ByVal servidor As String, ByVal bd As String, ByVal passw As String)
        Me.new(MontaCadenaConexion(usuario, servidor, bd, passw))
    End Sub

    Protected Shared Function MontaCadenaConexion(ByVal usuario As String, ByVal servidor As String, ByVal bd As String, ByVal passw As String) As String
        Return "Server=" & servidor & ";Database=" & bd & ";User ID=" & usuario & ";Password=" & passw & ""
    End Function

    Public Sub ComienzaTransaccion()
        entrasaccion = True
        Me.ComienzaTransaccion_interno()
    End Sub
    Protected MustOverride Sub ComienzaTransaccion_interno()
    Protected MustOverride Sub TerminaTransaccion_interno(ByVal exito As Boolean)
    Public MustOverride Sub BulkCopy(ByVal dt As System.Data.DataTable, ByVal nombretabladestino As String)


    Public Overridable Sub TerminaTransaccion(ByVal exito As Boolean)
        If entrasaccion Then
            Me.TerminaTransaccion_interno(exito)
            entrasaccion = False
        End If
    End Sub

    Public ReadOnly Property Entransaccion() As Boolean
        Get
            Return entrasaccion
        End Get
    End Property


    Public MustOverride Sub CerrarConexion()
    Public MustOverride Sub Ejecuta(ByVal expresion As String)
    Public MustOverride Function SelectSQL(ByVal expresion As String, Optional ByVal nombre As String = "") As DataTable
    Public MustOverride Function ExisteTablaSQL(ByVal nombretabla As String) As Boolean
    Protected MustOverride Sub RellenaDataset(ByVal expresion As String, ByVal datasetexistente As DataSet, ByVal nombretabla As String, ByVal maximoregistros As Integer)

    Public Function ejecuta1v_string(ByVal s As String) As String
        Dim t As DataTable
        Dim obj As Object
        Dim r As String

        t = Me.SelectSQL(s)
        If t.Rows.Count = 0 Then
            r = ""
        Else
            obj = t.Rows(0)(0)
            If obj Is System.DBNull.Value Then
                r = ""
            Else
                r = CStr(obj)
            End If
        End If
        Return r
        Me.Libera(t)
    End Function

    Public Function ejecuta1v_Double(ByVal s As String) As Double
        Dim t As DataTable
        Dim r As Double
        Dim obj As Object

        t = Me.SelectSQL(s)
        If t.Rows.Count = 0 Then
            r = 0.0
        Else
            obj = t.Rows(0)(0)
            If obj Is System.DBNull.Value Then
                r = 0.0
            Else
                r = CDbl(obj)
            End If
        End If
        Me.Libera(t)
        Return r
    End Function
    Public Function ejecuta1v_byte(ByVal s As String) As String
        Dim t As DataTable
        Dim b As Byte()
        Dim r As String
        Dim i As Integer
        Dim temp As String = String.Empty


        t = Me.SelectSQL(s)
        If t.Rows.Count = 0 Then
            r = ""
        Else
            b = CType(((t.Rows(0)(0))), Byte())

            r = BytesaCadena(b)
            For i = 0 To b.Length - 1
                temp &= Hex(b(i))
            Next i
            r = temp

        End If

        Return r
        Me.Libera(t)
    End Function
    Public Function ejecuta1v_long(ByVal s As String) As Long
        Dim t As DataTable
        Dim r As Long
        Dim obj As Object

        t = Me.SelectSQL(s)
        If t.Rows.Count = 0 Then
            r = 0
        Else
            obj = t.Rows(0)(0)
            If obj Is System.DBNull.Value Then
                r = 0
            Else
                r = CLng(obj)
            End If
        End If
        Me.Libera(t)
        Return r
    End Function

    Public Function ejecuta1v_bool(ByVal s As String) As Boolean
        Dim t As DataTable
        Dim r As Boolean
        Dim obj As Object

        t = Me.SelectSQL(s)
        If t.Rows.Count = 0 Then
            r = False
        Else
            obj = t.Rows(0)(0)
            If obj Is System.DBNull.Value Then
                r = False
            Else
                r = CBool(obj)
            End If
        End If
        Me.Libera(t)
        Return r
    End Function

    Public Sub Libera(ByVal dt As DataTable)
        Dim idtablaaliberar As String
        If dt Is Nothing Then Exit Sub


        idtablaaliberar = dt.TableName
        With dbDataset.Tables
            If .Contains(idtablaaliberar) Then
                .Remove(dt)
            End If
        End With
        If dbDataset.Tables.Count = 0 Then
            idtabla = 0 ' asi reseteamos el contador de tablas cada vez que se 
            ' se liberen todas.
        Else
            If IsNumeric(idtablaaliberar) Then
                If CDec(idtablaaliberar) = idtabla - 1 Then
                    idtabla = idtabla - 1
                End If
            End If
        End If
    End Sub

    Public Sub LiberaTodo()
        dbDataset.Tables.Clear()
    End Sub

    Public Sub EscribeEsquema_XML_de_Dataset(ByVal fichero As String)
        Me.dbDataset.WriteXml(fichero, XmlWriteMode.WriteSchema)
    End Sub
    Public Sub EscribeEsquema_XML_de_consulta(ByVal expresion As String, ByVal fichero As String, Optional ByVal maximoregistros As Integer = 0)
        Dim conjunto As New DataSet

        Me.RellenaDataset(expresion, conjunto, NOMBRE_TABLA_EXPORTACION_ESQUEMA_SQL, maximoregistros)

        If maximoregistros = 0 Then
            conjunto.WriteXmlSchema(fichero)
        Else
            conjunto.WriteXml(fichero, XmlWriteMode.WriteSchema)
        End If


    End Sub
End Class

Public Class ConexionSQL
    Inherits Conexion_x_SQL

    Friend dbConnection As SqlClient.SqlConnection

    Friend dbDataAdapter As SqlClient.SqlDataAdapter

    Protected tran As System.Data.SqlClient.SqlTransaction

    ReadOnly Property Servidor() As String
        Get
            Return dbConnection.DataSource
        End Get
    End Property
    ReadOnly Property Basededatos() As String
        Get
            Return dbConnection.Database
        End Get
    End Property



    Public Sub New(ByVal cadconex As String, ByVal dbDataSetExistente As Data.DataSet)
        MyBase.New(cadconex, dbDataSetExistente)

        dbConnection = New SqlClient.SqlConnection(cadenaconexion)
        dbConnection.Open()
        Me.InicializaProcedimientos()

    End Sub
    Public Sub New(ByVal cadconex As String)
        MyBase.New(cadconex)
        
        dbConnection = New SqlClient.SqlConnection(cadenaconexion)
        dbConnection.Open()
        Me.InicializaProcedimientos()
    End Sub

    Public Sub New(ByVal usuario As String, ByVal servidor As String, ByVal bd As String, ByVal passw As String)
        Me.new(MontaCadenaConexion(usuario, servidor, bd, passw))
    End Sub



    Protected Overridable Sub InicializaProcedimientos()

    End Sub



    Public Function Clonar() As ConexionSQL
        Return New ConexionSQL(cadenaconexion)
    End Function

    Protected Overrides Sub ComienzaTransaccion_interno()
        tran = dbConnection.BeginTransaction()
    End Sub
    Protected Overrides Sub TerminaTransaccion_interno(ByVal exito As Boolean)
        If entrasaccion Then
            If exito Then
                tran.Commit()
            Else
                tran.Rollback()
            End If

            tran = Nothing
            '¿una coleccion indizada por el nombre?

            Me.InicializaProcedimientos()
            ' asi forzamos a que se vuelvan a crear con la siguiente transaccion
        End If
    End Sub
    Public Function ObtenSqlCommand(ByVal expresion As String, ByVal tipo As System.Data.CommandType) As System.Data.SqlClient.SqlCommand
        Dim comando As SqlClient.SqlCommand
        If entrasaccion Then
            comando = New SqlClient.SqlCommand(expresion, dbConnection, tran)
        Else
            comando = New SqlClient.SqlCommand(expresion, dbConnection)
        End If
        comando.CommandType = tipo
        comando.CommandTimeout = 300
        Return comando
    End Function

    Protected Overrides Sub RellenaDataset(ByVal expresion As String, ByVal datasetexistente As DataSet, ByVal nombretabla As String, ByVal maximoregistros As Integer)
        RellenaDataset(New SqlClient.SqlDataAdapter(expresion, dbConnection), datasetexistente, nombretabla, maximoregistros)
    End Sub

    'ESTA FUNCION RELLENA EL DATASET CON EL DATAADAPTER, y devuelve la tabla que acaba de crear
    Protected Overloads Function RellenaDataset(ByVal dbDA As System.Data.SqlClient.SqlDataAdapter, ByVal datasetexistente As DataSet, ByVal nombretabla As String, ByVal maximoregistros As Integer) As DataTable
        Dim esvacio As Boolean = (nombretabla = "")

        If esvacio Then
			While datasetexistente.Tables.Contains(idtabla.ToString)
                idtabla += 100
            End While
            nombretabla = idtabla.ToString
        End If

        Try
            With dbDA
                If entrasaccion Then
                    .SelectCommand.Transaction = tran
                    .SelectCommand.CommandTimeout = 500
                End If
                .MissingSchemaAction = MissingSchemaAction.Add
                If maximoregistros < 0 Then
                    .Fill(datasetexistente, nombretabla)
                Else
                    .Fill(datasetexistente, 0, maximoregistros, nombretabla)
                End If
            End With
            If esvacio Then
                idtabla += 1
            End If
        Catch ex As System.Data.SqlClient.SqlException
            Throw New Exception(ProcesaErrorSQL(ex))
        End Try

        Return datasetexistente.Tables(nombretabla)
    End Function


    Public Sub SelectSql_enDataSet(ByVal expresion As String, ByVal datasetexistente As DataSet)
        RellenaDataset(expresion, datasetexistente, NOMBRE_TABLA_EXPORTACION_ESQUEMA_SQL, -1)
    End Sub
    Public Function SelectSql_endataset(ByVal expresion As String) As DataSet
        Dim nuevodataset As DataSet = New DataSet
        SelectSql_endataset(expresion, nuevodataset)
        Return nuevodataset
    End Function

    Public Function SelectSQL_cmd(ByVal cmd As SqlClient.SqlCommand, Optional ByVal nombre As String = "") As DataTable
        Return RellenaDataset(New SqlClient.SqlDataAdapter(cmd), dbDataset, nombre, -1)
    End Function

    Public Overrides Function SelectSQL(ByVal expresion As String, Optional ByVal nombre As String = "") As DataTable
        Return Me.SelectSQL_cmd(Me.ObtenSqlCommand(expresion, CommandType.Text), nombre)
    End Function
   

    Public Overrides Sub BulkCopy(ByVal dt As System.Data.DataTable, ByVal nombretabladestino As String)
        Dim bcp As System.Data.SqlClient.SqlBulkCopy

        Try
            If Me.entrasaccion Then
                bcp = New System.Data.SqlClient.SqlBulkCopy(Me.dbConnection, SqlClient.SqlBulkCopyOptions.TableLock, tran)
            Else
                bcp = New System.Data.SqlClient.SqlBulkCopy(Me.dbConnection)
            End If

            bcp.DestinationTableName = nombretabladestino
            bcp.WriteToServer(dt)
        Catch e As System.Data.SqlClient.SqlException
            Throw New Exception(Me.ProcesaErrorSQL(e))
        End Try
    End Sub

    Public Overrides Function ExisteTablaSQL(ByVal nombretabla As String) As Boolean
        Dim cadsql As String
        If nombretabla.StartsWith("#") Then
            nombretabla = "tempdb.." & nombretabla
        End If
        cadsql = "select isnull(OBJECT_ID('" & nombretabla & "'),0)"
        Return Me.ejecuta1v_long(cadsql) > 0
    End Function


    Public Overrides Sub Ejecuta(ByVal expresion As String)
        Ejecuta(ObtenSqlCommand(expresion, CommandType.Text))
    End Sub

    Public Overloads Sub Ejecuta(ByVal c As SqlClient.SqlCommand)

        Try
            c.ExecuteNonQuery()
        Catch e As System.Data.SqlClient.SqlException
            Throw New Exception(ProcesaErrorSQL(e))
        Finally
            c.Dispose()
        End Try

    End Sub




    Public Sub AddParametros(ByVal c As SqlClient.SqlCommand, ByVal parametros As Collection)
        Dim par As SqlClient.SqlParameter
        For Each o As Object In parametros
            par = DirectCast(o, SqlClient.SqlParameter)
            c.Parameters.Add(par)
        Next
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()

        'CerrarConexion
    End Sub

    Public Overrides Sub CerrarConexion()
        If dbConnection.State = ConnectionState.Open Then
            dbConnection.Close()
        End If
    End Sub

    Public Function ProcesaErrorSQL(ByVal e As System.Data.SqlClient.SqlException) As String
        Dim errorMessages as string = string.empty
        Dim i As Integer

        For i = 0 To e.Errors.Count - 1
            errorMessages += "Index #" & i.ToString() & ControlChars.NewLine _
                          & "Message: " & e.Errors(i).Message & ControlChars.NewLine _
                          & "LineNumber: " & e.Errors(i).LineNumber & ControlChars.NewLine _
                          & "Source: " & e.Errors(i).Source & ControlChars.NewLine _
                          & "Procedure: " & e.Errors(i).Procedure & ControlChars.NewLine
        Next i
        ' MsgBox(errorMessages, MsgBoxStyle.Critical, "Atención")
        Return errorMessages
    End Function



End Class




Public Class ConexionExcel
    Inherits DefinicionFuncionTexto

    Private Const ETIQUETA_RUTA As String = "RUTA" & Util.PARENTESISI
    ' Private Const ETIQUETA_MOSTRAR As String = "MOSTRAR" & Util.PARENTESISI
    Private Const ETIQUETA_TITULO As String = "TITULO" & Util.PARENTESISI
    Private Const ETIQUETA_CADENA As String = "CADENA" & Util.PARENTESISI
    Private Const ETIQUETA_FLOTANTE As String = "FLOTANTE" & Util.PARENTESISI
    Private Const ETIQUETA_ENTERO As String = "ENTERO" & Util.PARENTESISI
    Private Const ETIQUETA_FECHA As String = "FECHA" & Util.PARENTESISI
    Private Const ETIQUETA_HORA As String = "HORA" & Util.PARENTESISI
    Private Const ETIQUETA_HOJA As String = "HOJA" & Util.PARENTESISI
    Private Const ETIQUETA_OCULTARDUPLICADOS As String = "OCULTARDUPLICADOS" & Util.PARENTESISI

    'Para la lectura
    Private Const ETIQUETA_TUPLA As String = "TUPLA" & Util.PARENTESISI
    Private Const ETIQUETA_TUPLA_CTE_ENTERO As String = "ELEMENTO_TUPLA_ENTERO_CTE" & Util.PARENTESISI
    Private Const ETIQUETA_TUPLA_CTE_CADENA As String = "ELEMENTO_TUPLA_CADENA_CTE" & Util.PARENTESISI
    Private Const ETIQUETA_TUPLA_CTE_FLOTANTE As String = "ELEMENTO_TUPLA_FLOTANTE_CTE" & Util.PARENTESISI
    Private Const ETIQUETA_TUPLA_CTE_FECHA As String = "ELEMENTO_TUPLA_FECHA_CTE" & Util.PARENTESISI
    Private Const ETIQUETA_TUPLA_TITULO As String = "ELEMENTO_TUPLA_CAMPOFIJO" & Util.PARENTESISI
    Private Const ETIQUETA_TUPLA_CAMPO As String = "ELEMENTO_TUPLA_CAMPO" & Util.PARENTESISI

    Private Const ETIQUETA_DIRHISTORICO As String = "DIRHISTORICO" & Util.PARENTESISI

    Private tokens() As String = {ETIQUETA_RUTA, ETIQUETA_TITULO, ETIQUETA_ENTERO, ETIQUETA_CADENA, ETIQUETA_FLOTANTE, ETIQUETA_HOJA, ETIQUETA_FECHA, ETIQUETA_HORA, ETIQUETA_OCULTARDUPLICADOS, ETIQUETA_TUPLA, ETIQUETA_TUPLA_CTE_ENTERO, ETIQUETA_TUPLA_CTE_CADENA, ETIQUETA_TUPLA_CTE_FLOTANTE, ETIQUETA_TUPLA_CTE_FECHA, ETIQUETA_TUPLA_TITULO, ETIQUETA_TUPLA_CAMPO, ETIQUETA_DIRHISTORICO}


    Private ruta As String
    Private previsualizar As Boolean
    Private hojas As Coleccion_HojasExcel

    'Agregar Referencia Microsoft Excel 11.0 Object Library
    Shared appexcel As New Microsoft.Office.Interop.Excel.Application
    Dim libroexcel As Microsoft.Office.Interop.Excel.Workbook


    Private Const MARCAPARAOCULTAR As String = "$$)(34@"
    Private numerohojasparaocultar As Integer

    Private numerodeescrituras As Integer

    Private padre As Ejecucion_TPR

    Private rutahistorico As String
    Private pasarahistorico As Boolean


    Private tuplaslectura As New Coleccion_tupla_lecturaExcel
    Private archivosexistentes As ColecciondeString

    Public Sub New(ByVal objpadre As Ejecucion_TPR, ByVal s As String, ByVal constantes As Coleccion_Constantes)
        'en s viene la estructura del fichero
        Dim col As ColecciondeTokens
        Dim tn As Token_Nodo = Nothing
        Dim i, il As Integer
        Dim contenido As String
        Dim ct As campo_titulo_excel
        Dim c_e As campo_excel
        Dim tuplacte As Tupla_Elemento_CTE


        padre = objpadre
        hojas = New Coleccion_HojasExcel

        col = Parser_TPR.SplitPrefijoMultiple(tokens, s)

        il = col.Count - 1

        For i = 0 To il
            Try
                tn = col.Item(i)
                contenido = Me.QuitaParentesisFinal(tn.contenido)
                Select Case tn.TOKEN

                    Case ETIQUETA_RUTA
                        ruta = Parser_TPR.ExtraeParametroCadena(contenido, constantes)
                        '   Case ETIQUETA_MOSTRAR
                        '      Me.previsualizar = Parser_TPR.ExtraeParametrobooleano(contenido)
                    Case ETIQUETA_HOJA
                        hojas.Add(New HojaExcel(Parser_TPR.ExtraeParametroCadena(contenido, constantes).ToUpper))
                    Case ETIQUETA_TITULO
                        Try
                            ct = New campo_titulo_excel(contenido, constantes)
                            hojas.Item(ct.nombrehoja).titulos.Add(ct)
                        Catch ex As Exception
                            Throw New Exception("Titulo : " & ex.Message)
                        End Try
                    Case ETIQUETA_OCULTARDUPLICADOS
                        Try
                            Dim od As Excel_ocultar_duplicados
                            od = New Excel_ocultar_duplicados(contenido, constantes)
                            With hojas.Item(od.nombrehoja)
                                .campos.Item(od.campoaocultar).ocultarduplicados = od.ocultarduplicados
                                .campos.Item(od.campoaocultar).camporeferencia_ocultar = od.camporeferencia

                            End With
                        Catch ex As Exception
                            Throw New Exception("Ocultar Duplicados (nombrehoja, Booleano, camporeferencia) " & ex.Message)
                        End Try

                    Case ETIQUETA_ENTERO, ETIQUETA_CADENA, ETIQUETA_FLOTANTE, ETIQUETA_FECHA, ETIQUETA_HORA
                        Select Case tn.TOKEN
                            Case ETIQUETA_ENTERO

                                c_e = New campo_excel_entero(contenido, constantes)

                            Case ETIQUETA_CADENA

                                c_e = New campo_excel_cadena(contenido, constantes)

                            Case ETIQUETA_FLOTANTE
                                c_e = New campo_excel_flotante(contenido, constantes)

                            Case ETIQUETA_FECHA

                                c_e = New campo_excel_fecha(contenido, constantes)

                            Case ETIQUETA_HORA
                                c_e = New campo_excel_hora(contenido, constantes)
                            Case Else
                                c_e = Nothing
                        End Select

                        hojas.Item(c_e.nombrehoja).campos.Add(c_e)

                    Case ETIQUETA_TUPLA
                        'CREAMOS un nuevo objeto tupla
                        'lo insertamos en la hoja
                        Dim tupla As New Tupla_lecturaExcel(contenido, constantes)

                        If Me.hojas.Existe(tupla.nombrehoja) Then

                            'La tupla hace referencia a una hoja, pero realmente
                            'pertenece a la conexion 
                            Me.tuplaslectura.add(tupla)
                        Else
                            Throw New Exception("No existe la hoja Excel que hace referencia la Tupla " & tupla.nombretupla)
                        End If



                    Case ETIQUETA_TUPLA_CTE_ENTERO, ETIQUETA_TUPLA_CTE_CADENA, ETIQUETA_TUPLA_CTE_FLOTANTE, ETIQUETA_TUPLA_CTE_FECHA
                        Select Case tn.TOKEN
                            Case ETIQUETA_TUPLA_CTE_ENTERO
                                tuplacte = New Tupla_ELEMENTO_CTE_Entero(contenido, constantes)
                            Case ETIQUETA_TUPLA_CTE_CADENA
                                tuplacte = New Tupla_ELEMENTO_CTE_Cadena(contenido, constantes)
                            Case ETIQUETA_TUPLA_CTE_FLOTANTE
                                tuplacte = New Tupla_ELEMENTO_CTE_flotante(contenido, constantes)
                            Case ETIQUETA_TUPLA_CTE_FECHA
                                tuplacte = New Tupla_ELEMENTO_CTE_Fecha(contenido, constantes)
                            Case Else
                                Throw New Exception("Token encontrado " & tn.TOKEN & " no tiene definicion")
                        End Select
                        'para cada una de los elementos constantes de tupla, comprobamos que la tupla exista 
                        If Me.tuplaslectura.Existe(tuplacte.nombretupla) Then
                            Me.tuplaslectura.Item(tuplacte.nombretupla).valores_constantes.Add(tuplacte)
                        End If

                    Case ETIQUETA_TUPLA_TITULO
                        Dim tuplatit As Tupla_Elemento_Titulo
                        tuplatit = New Tupla_Elemento_Titulo(contenido, constantes)
                        With tuplatit
                            If Me.tuplaslectura.Existe(.nombretupla) Then
                                Me.tuplaslectura.Item(.nombretupla).valores_titulos.Add(tuplatit)
                            End If
                        End With

                    Case ETIQUETA_TUPLA_CAMPO
                        Dim tuplacamp As Tupla_Elemento_Campo
                        tuplacamp = New Tupla_Elemento_Campo(contenido, constantes)
                        With tuplacamp
                            If Me.tuplaslectura.Existe(.nombretupla) Then
                                Me.tuplaslectura.Item(.nombretupla).valores_campos.Add(tuplacamp)
                            End If
                        End With

                    Case ETIQUETA_DIRHISTORICO
                        rutahistorico = Util.NormalizaRuta(Parser_TPR.ExtraeParametroCadena(contenido, constantes))
                        pasarahistorico = rutahistorico <> VACIO


                    Case Else
                        Throw New Exception("Token sin definicion")
                End Select
            Catch ex As Exception
                Throw New Exception("Definición Conexion EXCEL: " & tn.TOKEN & ".Posicion:" & tn.posicion & ".Contenido:" & tn.contenido & ex.Message)
            End Try
        Next

        previsualizar = (Me.ruta = VACIO)

        Abre()

    End Sub

    Public Sub Abre()
        Dim h As Microsoft.Office.Interop.Excel.Worksheet

        numerodeescrituras = 0
        ' appexcel = New Microsoft.Office.Interop.Excel.Application
        'Si existe el libro, lo abrimos, sino creamos una nuevo

        archivosexistentes = BuscaFicheroParaLectura(Me.ruta, False, "")

        If archivosexistentes.Count = 0 Then ' NO existe el fichero, lo creamos
            libroexcel = appexcel.Workbooks.Add()
            'creamos un libro vacio, como por defecto trae tres hojas, las borramos

            numerohojasparaocultar = 0
            For Each h In libroexcel.Worksheets
                h.Name = MARCAPARAOCULTAR & h.Name
                numerohojasparaocultar += 1
            Next

            'else
            'si existe el fichero (o los ficheros) en la lectura se consultara la 
            'variable archivosexistentes
        End If
    End Sub

    Public Sub CerrarAbortando()
        Try
            Dim wb As Microsoft.Office.Interop.Excel.Workbook

            appexcel.DisplayAlerts = False
            For Each wb In appexcel.Workbooks
                wb.Close(False)
            Next
            appexcel.Quit()
        Catch ex As Exception
            'Ya estamos en una excepcion, no genereamos nada
        End Try
    End Sub
    Public Sub Cerrar()
        Dim h As Microsoft.Office.Interop.Excel.Worksheet

        If numerodeescrituras > 0 Then
            If libroexcel.Worksheets.Count > numerohojasparaocultar Then
                For Each h In libroexcel.Worksheets
                    If h.Name.StartsWith(MARCAPARAOCULTAR) Then
                        h.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden
                    End If
                Next
            End If
            If previsualizar Then
                If appexcel.Visible = False Then
                    appexcel.Visible = True
                End If
            Else
                libroexcel.SaveAs(ruta)
                appexcel.Quit()
            End If
        Else
            appexcel.DisplayAlerts = False
            appexcel.Quit()
        End If
    End Sub


    Public Sub Escribe(ByVal dt As DataTable, ByVal cm As Coleccion_mapeocampos, ByVal nombrehoja As String)
        'Abrimos el Excel
        'Escribimos el Excel 

        Dim hoja As Microsoft.Office.Interop.Excel.Worksheet
        Dim i, il, filamaxima As Integer
        Dim dr As DataRow
        Dim msj as string = string.empty
        Dim encontradahoja As Boolean
        Dim h As HojaExcel
        Dim ultimahojaencontrada As Microsoft.Office.Interop.Excel.Worksheet = Nothing


        Try
            h = Me.hojas.Item(nombrehoja)


            padre.InformaEventoHijo_Inicio(Ejecucion_TPR.NIVEL_INFORMACION_EVENTO.NIVEL2, dt.Rows.Count + h.titulos.Count)

            numerodeescrituras += 1
            encontradahoja = False
            hoja = Nothing
            For i = 1 To libroexcel.Worksheets.Count
                hoja = CType(libroexcel.Worksheets.Item(i), Microsoft.Office.Interop.Excel.Worksheet)
                ultimahojaencontrada = hoja
                If hoja.Name = nombrehoja Then
                    encontradahoja = True
                    Exit For
                End If
            Next

            If Not encontradahoja Then
                hoja = CType(libroexcel.Worksheets.Add(, ultimahojaencontrada), Microsoft.Office.Interop.Excel.Worksheet)
                hoja.Name = nombrehoja
            End If


            filamaxima = 1


            h.hoja = hoja

            il = h.titulos.Count - 1

            For i = 0 To il
                With h.titulos.Item(i)
                    .Escribe(h)
                    If .Fila > filamaxima Then
                        filamaxima = .Fila
                    End If
                End With
                padre.InformaEventoHijo_Avance(Ejecucion_TPR.NIVEL_INFORMACION_EVENTO.NIVEL2, 1)
            Next
            filamaxima += 1


            il = h.campos.Count - 1


            If cm.ComprobarMapeo(h.campos, dt, msj) Then
                For i = 0 To il
                    h.campos.Item(i).Fila = filamaxima
                Next
                For Each dr In dt.Rows
                    If padre.cancelar Then
                        Throw New Exception_Proceso_ejecucion_TPR_cancelado_por_usuario("Cancelado por el usuario")
                    End If
                    For i = 0 To il
                        With h.campos.Item(i)
                            .Escribe(dr, cm.Mapea(.Nombrecampo), h)
                            .AddFila(1) 'Sumamos uno a la fila por la vamos escribiendo
                        End With
                    Next
                    padre.InformaEventoHijo_Avance(Ejecucion_TPR.NIVEL_INFORMACION_EVENTO.NIVEL2, 1)
                Next

                h.AjustarAnchoColumnas()
            Else
                Throw New Exception("Mapeo de Campos: " & msj)
            End If


        Catch ex_cancelado As Exception_Proceso_ejecucion_TPR_cancelado_por_usuario
            Throw New Exception_Proceso_ejecucion_TPR_cancelado_por_usuario(ex_cancelado.Message)
        Catch ex As Exception
            Throw New Exception("Escribir Excel: " & ex.Message)
        Finally
            padre.InformaEventoHijo_Fin(Ejecucion_TPR.NIVEL_INFORMACION_EVENTO.NIVEL2)
        End Try
    End Sub


    Public Sub Lee(ByVal tabladestino As DataTable, ByVal mapeo As Coleccion_mapeocampos)
        'Tenemos que montar la sentencia Ms-SQL para grabar 
        'DE TODOS LOS FICHEROS
        'TODAS las tuplas
        'todos los registros de cada tupla
        Dim ifichero, itupla, j As Integer
        Dim msj As String = VACIO
        Dim hoja As Microsoft.Office.Interop.Excel.Worksheet
        Dim encontrada As Boolean
        Dim nombreyrutafichero As String
        Dim ilocal, iglobal, nlineas As Integer




        Dim dr As DataRow

        'Primero comprobamos los mapeos de las tuplas

        padre.InformaEventoHijo_Inicio(Ejecucion_TPR.NIVEL_INFORMACION_EVENTO.NIVEL2, 0)

        For itupla = 0 To Me.tuplaslectura.Count - 1
            With Me.tuplaslectura.Item(itupla)
                If Not mapeo.ComprobarMapeo(.valores_constantes, tabladestino, msj) Then
                    Throw New Exception(msj)
                End If

                If Not mapeo.ComprobarMapeo(.valores_titulos, tabladestino, msj) Then
                    Throw New Exception(msj)
                End If

                If Not mapeo.ComprobarMapeo(.valores_campos, tabladestino, msj) Then
                    Throw New Exception(msj)
                End If
            End With
        Next

        tabladestino.BeginLoadData()
        iglobal = 0
        For ifichero = 0 To archivosexistentes.Count - 1
            nombreyrutafichero = archivosexistentes.Item(ifichero)

            If System.IO.File.Exists(nombreyrutafichero) Then
                Me.libroexcel = appexcel.Workbooks.Open(nombreyrutafichero, , False)

                For itupla = 0 To Me.tuplaslectura.Count - 1
                    'Asignamos la hoja del fichero excel

                    With Me.tuplaslectura.Item(itupla)
                        encontrada = False
                        hoja = Nothing
                        For Each h As Microsoft.Office.Interop.Excel.Worksheet In libroexcel.Worksheets
                            If h.Name.ToUpper = .nombrehoja.ToUpper Then
                                hoja = h
                                encontrada = True
                            End If
                        Next
                        If encontrada = False Then
                            Throw New Exception("No se encuentra la hoja " & .nombrehoja)
                        End If
                        'hoja = CType(libroexcel.Worksheets.Item(.nombrehoja), Microsoft.Office.Interop.Excel.Worksheet)

                        For j = 0 To .valores_titulos.Count - 1
                            .valores_titulos.Item(j).hoja = hoja
                        Next

                        For j = 0 To .valores_campos.Count - 1
                            .valores_campos.Item(j).hoja = hoja
                        Next
                    End With


                    'If .valores_campos.Count > 0 Then
                    'While Not .valores_campos.Item(0).Estavacio()





                    With Me.tuplaslectura.Item(itupla)
                        'Aqui incluimos el bucle.

                        'Si vamos leyendo campos hasta que no haya mas
                        'valores, no sabemos cuantas lineas son
                        'nlineas se va actualizando en cada iteracion

                        nlineas = 1
                        ilocal = 0
                        While ilocal < nlineas
                            dr = tabladestino.NewRow
                            For j = 0 To .valores_constantes.Count - 1
                                With .valores_constantes.Item(j)
                                    .AsignaValor(dr, mapeo.Mapea(.nombreelementotupla), ilocal, iglobal)
                                End With

                            Next


                            For j = 0 To .valores_titulos.Count - 1
                                With .valores_titulos.Item(j)
                                    .AsignaValor(dr, mapeo.Mapea(.nombreelementotupla), ilocal, iglobal)
                                End With
                            Next

                            'Secundario
                            'Al añadir los campos habrá que iterar a partir de aqui
                            'hay que definir una condicion de terminación (un campo "clave" que si esta vacio hemos acabado?)


                            For j = 0 To .valores_campos.Count - 1
                                With .valores_campos.Item(j)
                                    .AsignaValor(dr, mapeo.Mapea(.nombreelementotupla), ilocal, iglobal)
                                End With
                            Next
                            tabladestino.Rows.Add(dr)
                            iglobal += 1
                            ilocal += 1
                            If .valores_campos.Count > 0 AndAlso _
                                Not .valores_campos.Item(0).Estavacio() Then
                                nlineas += 1
                            End If

                            ' If iglobal Mod 200 = 0 Then
                            padre.InformaEventoHijo_Avance(Ejecucion_TPR.NIVEL_INFORMACION_EVENTO.NIVEL2, 1)
                            'End If
                        End While
                    End With
                Next

                libroexcel.Close()
                If pasarahistorico Then
                    Parser_TPR.MueveaDirectorioDestino(nombreyrutafichero, Me.rutahistorico & System.IO.Path.GetFileNameWithoutExtension(nombreyrutafichero) & "." & Now.ToString("yyyyMMddHHmmssfff"), False, "")
                End If


            End If

        Next
        padre.InformaEventoHijo_Fin(Ejecucion_TPR.NIVEL_INFORMACION_EVENTO.NIVEL2)
        tabladestino.EndLoadData()




    End Sub


End Class
#End If


Public MustInherit Class base_where_sql
    Protected p_where As condicion_sql
    Protected nombretabla As String

    'Public Enum tipooperacionsql
    '    igual
    '    menor
    '    mayor
    '    menorigual
    '    mayorigual
    '    distinto
    '    como
    'End Enum

    Public Sub New()
        p_where = New condicion_sql(Util.SQL_WHERE)
        Limpia()
    End Sub

    Public Overridable Sub Limpia()
        p_where.Limpia()
        nombretabla = VACIO
    End Sub

    Public Overridable WriteOnly Property Tabla() As String
        Set(ByVal Value As String)
            nombretabla = Value
        End Set
    End Property

    Protected Function colocaw(ByVal s As String) As String
        If s = "" Then
            Return Util.SQL_WHERE
        Else
            Return s & Util._Y_
        End If
    End Function

    'Protected Function l_op(ByVal t As condicion_sql.tipooperacionsql) As String
    '    Select Case t
    '        Case condicion_sql.tipooperacionsql.igual
    '            Return " = "
    '        Case condicion_sql.tipooperacionsql.menor
    '            Return " < "
    '        Case condicion_sql.tipooperacionsql.mayor
    '            Return " > "
    '        Case condicion_sql.tipooperacionsql.menorigual
    '            Return " <= "
    '        Case condicion_sql.tipooperacionsql.mayorigual
    '            Return " >= "
    '        Case condicion_sql.tipooperacionsql.distinto
    '            Return " <> "
    '        Case condicion_sql.tipooperacionsql.como
    '            Return " LIKE "

    '    End Select
    'End Function


    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As String, Optional ByVal op As condicion_sql.tipooperacionsql = condicion_sql.tipooperacionsql.igual)
        p_where.add_condicion(nombrecampo, valor, op)
    End Sub
    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As Integer, Optional ByVal op As condicion_sql.tipooperacionsql = condicion_sql.tipooperacionsql.igual)
        p_where.add_condicion(nombrecampo, valor, op)
    End Sub
    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As Long, Optional ByVal op As condicion_sql.tipooperacionsql = condicion_sql.tipooperacionsql.igual)
        p_where.add_condicion(nombrecampo, valor, op)
    End Sub

    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As Double, Optional ByVal op As condicion_sql.tipooperacionsql = condicion_sql.tipooperacionsql.igual)
        p_where.add_condicion(nombrecampo, valor, op)
    End Sub

    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As Date, Optional ByVal op As condicion_sql.tipooperacionsql = condicion_sql.tipooperacionsql.igual)
        p_where.add_condicion(nombrecampo, valor, op)
    End Sub

    Public Overridable Sub add_condicion_expr(ByVal expr1 As String)
        p_where.add_condicion_expr(expr1)
    End Sub
    Public Overridable Sub add_condicion_expr(ByVal expr1 As String, ByVal expr2 As String, ByVal op As condicion_sql.tipooperacionsql)
        p_where.add_condicion_expr(expr1, expr2, op)

    End Sub

    Public Overridable Sub add_condicion_func(ByVal f As String, ByVal ParamArray param() As String)
        p_where.add_condicion_func(f, param)
    End Sub

    Protected Function s_coma(ByVal s As String) As String
        If s <> VACIO Then
            Return s & COMA
        Else
            Return s
        End If
    End Function

    Public MustOverride ReadOnly Property texto_sql() As String

End Class


Public Class SentenciaSQL
    Inherits base_where_sql

    Private p_select As String
    Private p_from As String
    Private p_groupby As String
    Private p_having As String
    Private p_orderby As String

    Private dist As Boolean
    Private top As Integer

    Private Shadows WriteOnly Property Tabla() As String
        Set(ByVal Value As String)
            'Hago que no sea visible, no me interesa en esta clase (private + shadows)
        End Set
    End Property

    Public WriteOnly Property distinct() As Boolean
        Set(ByVal Value As Boolean)
            dist = Value
        End Set
    End Property

    Public WriteOnly Property top_n() As Integer
        Set(ByVal Value As Integer)
            top = Value
        End Set
    End Property

    Public WriteOnly Property sql_select() As String
        Set(ByVal Value As String)

            p_select = coloca_sel() & Value

        End Set
    End Property
    Public WriteOnly Property sql_from() As String
        Set(ByVal Value As String)
            p_from = Util.SQL_FROM & Value
        End Set
    End Property
    Public WriteOnly Property sql_where() As String
        Set(ByVal Value As String)
            p_where.texto_sql() = Value
        End Set
    End Property
    Public WriteOnly Property sql_having() As String
        Set(ByVal Value As String)
            p_having = Util.SQL_HAVING & Value
        End Set
    End Property
    Public WriteOnly Property sql_orderby() As String
        Set(ByVal Value As String)
            p_orderby = Util.sql_ORDERBY & Value
        End Set
    End Property
    Public WriteOnly Property sql_groupby() As String
        Set(ByVal Value As String)
            p_groupby = Util.sql_GROUPBY & Value
        End Set
    End Property

    Public Sub add_campo_select(ByVal ParamArray campo() As String)
        For Each s As String In campo
            add_campo_select(s)
        Next
    End Sub

    Public Enum dir_orderby
        ASC
        DESC
        none
    End Enum
    Public Sub add_orderby(ByVal campo As String, Optional ByVal sentido As dir_orderby = dir_orderby.none)
        Dim o As String
        Select Case sentido
            Case dir_orderby.ASC
                o = SQL_ASC
            Case dir_orderby.DESC
                o = SQL_DESC
            Case Else
                o = ""
        End Select

        If p_orderby = "" Then
            p_orderby = Util.sql_ORDERBY
        Else
            p_orderby &= COMA
        End If
        p_orderby &= campo & o
    End Sub
    Private Function coloca_sel() As String
        If dist Then
            Return Util.SQL_SELECT & Util.SQL_DISTINCT
        Else
            If top >= 0 Then
                Return Util.SQL_SELECT & Util.SQL_TOP & top.ToString
            Else
                Return Util.SQL_SELECT
            End If
        End If
    End Function
    Public Sub add_campo_select(ByVal campo As String)


        If p_select = "" Then
            p_select = coloca_sel()
        Else
            p_select &= Util.COMA
        End If
        p_select &= campo


    End Sub
    Public Sub New()
        MyBase.New()
        Limpia()
    End Sub

    Public Overrides Sub Limpia()
        MyBase.Limpia()
        Me.p_from = VACIO
        Me.p_groupby = VACIO
        Me.p_having = VACIO
        Me.p_orderby = VACIO
        Me.p_select = VACIO
        dist = False
        top = -1
    End Sub
    Public Sub New(ByVal s As String, ByVal f As String, ByVal w As String, ByVal g As String, ByVal h As String, ByVal o As String)
        Me.p_select = s
        Me.p_from = f
        Me.p_where.texto_sql() = w
        Me.p_groupby = g
        Me.p_having = h
        Me.p_orderby = o
    End Sub
    Public Overrides ReadOnly Property texto_sql() As String
        Get
            Return Me.p_select & Me.p_from & p_where.texto_sql() & Me.p_groupby & Me.p_having & Me.p_orderby
        End Get
    End Property

    Public Sub Importa(ByVal cadenasql As String)
        Importa(Util.SQL_SustCadenas(cadenasql), cadenasql)
    End Sub

    Public Sub Importa(ByVal base_mayuscula As String, ByVal cadenasql As String)
        'CUIDADO CON EL TOP, EN PRUEBAS
        Dim partes As ColecciondeTokens
        Dim i, il As Integer
        Dim t As Token_Nodo
        Dim tokens() As String = {Util.SQL_SELECT, Util.SQL_FROM, Util.SQL_WHERE, Util.sql_ORDERBY, Util.sql_GROUPBY, Util.SQL_HAVING}
        Dim s, s_top As String
        Const expr_TOP As String = "\s*TOP\s+[0-9]+\s+"
        Dim pos As Integer

        partes = Parser_TPR.SplitPrefijoMultiple(tokens, base_mayuscula, cadenasql)

        il = partes.Count - 1

        For i = 0 To il
            t = partes.Item(i)
            Select Case t.TOKEN
                Case Util.SQL_SELECT
                    s = t.contenido
                    If Regex.IsMatch(s, expr_TOP) Then
                        pos = Regex.Matches(s, expr_TOP).Item(0).Length
                        s_top = s.Substring(0, pos) 'sabemos que empieza en 0 por la expresion
                        s = s.Substring(pos)

                        'Ahora a analizar s_top en busca del numero
                        Me.top = CInt(Regex.Matches(s_top, "\s[0-9]+\s").Item(0).Value)
                    End If
                    Me.sql_select = s
                Case Util.SQL_FROM
                    Me.sql_from = t.contenido
                Case Util.SQL_WHERE
                    Me.sql_where = t.contenido
                Case Util.sql_ORDERBY
                    Me.sql_orderby = t.contenido
                Case Util.sql_GROUPBY
                    Me.sql_groupby = t.contenido
                Case Util.SQL_HAVING
                    Me.sql_having = t.contenido
            End Select
        Next
    End Sub





End Class



Public Class Update_sql
    Inherits base_where_sql

    'UPDATE (NOMBRE) SET campo=valor, campo2=valor2 where condicion

    Private cad_set As String

    Public Sub New()
        MyBase.new()
        Limpia()
    End Sub

    Public Overrides Sub Limpia()
        MyBase.Limpia()
        cad_set = VACIO
    End Sub

    Public Shadows WriteOnly Property Tabla() As String
        Set(ByVal Value As String)
            nombretabla = Value
        End Set
    End Property

    Public Overrides ReadOnly Property texto_sql() As String
        Get
            Return sql_UPDATE & nombretabla & Util._set_ & cad_set & p_where.texto_sql()
        End Get
    End Property


    Public Sub add_valor(ByVal campo As String, ByVal valor As String)
        cad_set = s_coma(cad_set) & campo & IGUAL & v(valor)
    End Sub

    Public Sub add_valor(ByVal campo As String, ByVal valor As Integer)
        cad_set = s_coma(cad_set) & campo & IGUAL & valor.ToString
    End Sub
    Public Sub add_valor(ByVal campo As String, ByVal valor As Double)
        cad_set = s_coma(cad_set) & campo & IGUAL & Util.doublesql(valor)
    End Sub
    Public Sub add_valor(ByVal campo As String, ByVal valor As Date)
        cad_set = s_coma(cad_set) & campo & IGUAL & fechasql(valor)
    End Sub
    Public Sub add_valor(ByVal campo As String, ByVal valor As Date, ByVal formatofecha As Util.tipofecha)
        cad_set = s_coma(cad_set) & campo & IGUAL & fechasql(valor, formatofecha)
    End Sub

    Public Sub add_valor_expr(ByVal campo As String, ByVal valor As String)
        cad_set = s_coma(cad_set) & campo & IGUAL & valor


    End Sub

End Class

Public Class insert_sql
    ' 
    Inherits base_where_sql

    Private listacampos As String
    Private listavalores As String

    Public Overrides ReadOnly Property texto_sql() As String
        Get
            Return SQL_INSERT_INTO & nombretabla & entreparen(listacampos) & _VALUES_ & entreparen(listavalores)
        End Get
    End Property
    Public Sub New()
        MyBase.new()
        Limpia()
    End Sub

    Public Overrides Sub Limpia()
        MyBase.Limpia()
        listacampos = VACIO
        listavalores = VACIO

    End Sub
    Private Sub add_campo(ByVal campo As String)
        listacampos = s_coma(listacampos) & campo
    End Sub
    Public Sub add_valor(ByVal campo As String, ByVal valor As String)
        add_campo(campo)
        listavalores = s_coma(listavalores) & v(valor)
    End Sub
    Public Sub add_valor(ByVal campo As String, ByVal valor As Byte)
        add_campo(campo)
        listavalores = s_coma(listavalores) & valor.ToString
    End Sub
    Public Sub add_valor(ByVal campo As String, ByVal valor As Integer)
        add_campo(campo)
        listavalores = s_coma(listavalores) & valor.ToString
    End Sub
    Public Sub add_valor(ByVal campo As String, ByVal valor As Double)
        add_campo(campo)
        listavalores = s_coma(listavalores) & Util.doublesql(valor)
    End Sub
    Public Sub add_valor(ByVal campo As String, ByVal valor As Date, Optional ByVal t As tipofecha = tipofecha.ddMMyyyy)
        add_campo(campo)
        listavalores = s_coma(listavalores) & fechasql(valor, t)
    End Sub

    Public Sub add_valor_expr(ByVal campo As String, ByVal valor As String)
        add_campo(campo)
        listavalores = s_coma(listavalores) & valor
    End Sub

End Class


Public Class delete_sql
    Inherits base_where_sql


    Public Overrides ReadOnly Property texto_sql() As String
        Get
            Return SQL_DELETE & Util.SQL_FROM & Me.nombretabla & p_where.texto_sql()

        End Get
    End Property

End Class


Public Class condicion_sql
    Private _t As String
    Private _prefijo As String

    Public Enum tipooperacionsql
        igual
        menor
        mayor
        menorigual
        mayorigual
        distinto
        como
        comono
    End Enum

    Public Sub New(ByVal prefijo As String)
        Limpia()
        _prefijo = prefijo
    End Sub

    Public Overridable Sub Limpia()
        _t = VACIO
    End Sub


    Protected Function colocaw(ByVal s As String) As String
        If s = VACIO Then
            Return VACIO
        Else
            Return s & Util._Y_
        End If
    End Function

    Protected Function l_op(ByVal t As tipooperacionsql) As String
        Select Case t
            Case tipooperacionsql.igual
                Return " = "
            Case tipooperacionsql.menor
                Return " < "
            Case tipooperacionsql.mayor
                Return " > "
            Case tipooperacionsql.menorigual
                Return " <= "
            Case tipooperacionsql.mayorigual
                Return " >= "
            Case tipooperacionsql.distinto
                Return " <> "
            Case tipooperacionsql.como
                Return " LIKE "
            Case tipooperacionsql.comono
                Return " NOT LIKE "
            Case Else
                Return " = "

        End Select
    End Function


    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As String, Optional ByVal op As tipooperacionsql = tipooperacionsql.igual)
        _t = colocaw(_t) & nombrecampo & l_op(op) & v(valor)
    End Sub
    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As Integer, Optional ByVal op As tipooperacionsql = tipooperacionsql.igual)
        _t = colocaw(_t) & nombrecampo & l_op(op) & valor
    End Sub
    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As Long, Optional ByVal op As tipooperacionsql = tipooperacionsql.igual)
        _t = colocaw(_t) & nombrecampo & l_op(op) & valor
    End Sub

    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As Double, Optional ByVal op As tipooperacionsql = tipooperacionsql.igual)
        _t = colocaw(_t) & nombrecampo & l_op(op) & Util.doublesql(valor)
    End Sub

    Public Overridable Sub add_condicion(ByVal nombrecampo As String, ByVal valor As Date, Optional ByVal op As tipooperacionsql = tipooperacionsql.igual)
        _t = colocaw(_t) & nombrecampo & l_op(op) & Util.fechasql(valor)
    End Sub

    Public Overridable Sub add_condicion_expr(ByVal expr1 As String)
        If expr1 <> "" Then
            _t = colocaw(_t) & expr1
        End If
    End Sub
    Public Overridable Sub add_condicion_expr(ByVal expr1 As String, ByVal expr2 As String, ByVal op As tipooperacionsql)
        _t = colocaw(_t) & expr1 & l_op(op) & expr2
    End Sub

    Public Overridable Sub add_condicion_func(ByVal f As String, ByVal ParamArray param() As String)
        Dim s As String
        Dim i, ini, fin As Integer
        s = VACIO
        ini = param.GetLowerBound(0)
        fin = param.GetUpperBound(0)
        For i = ini To fin
            s = s_coma(s) & param(i)
        Next
        _t = colocaw(_t) & f & entreparen(s)
    End Sub

    Protected Function s_coma(ByVal s As String) As String
        If s <> VACIO Then
            Return s & COMA
        Else
            Return s
        End If
    End Function

    Public Property texto_sql() As String
        Set(ByVal value As String)
            _t = value
        End Set
        Get
            If _t <> VACIO Then
                Return _prefijo & _t
            Else
                Return VACIO
            End If
        End Get
    End Property
End Class





