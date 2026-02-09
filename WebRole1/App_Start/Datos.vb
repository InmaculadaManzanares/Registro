Imports System.Data
Public Module Util_Basededatos
    'Hay que iniciarlo con los valores corectos al inicio de la aplicacion
    Public Tipo_DateTime As System.Type
    Public Tipo_String As System.Type
    Public Tipo_Int32 As System.Type
    Public Tipo_Double As System.Type
    Public Tipo_Single As System.Type
    'Es decir, llama a Util_BAsedeDatos.Inicializa antes de la primera declaracion de un objeto
    'de una clase heredada de Basededatos

    Public Const PrefijoPARAMETRO As String = "@PaRaMeTrO"

    Public Const COMILLA_DOBLE As String = Chr(34)
    Public Const DISTINTO As String = "<>"
    Public Const CADENAVACIA As String = ""
    Public Const ZERO As Long = 0
    Public Const ZERO_SNG As Single = 0.0

    Public Const ACT_NODEFINITIVO As Boolean = False
    Public Const ACT_DEFINITIVO As Boolean = True


    Public Const IGUAL As String = "="
    Public Const MAYOR As String = " > "
    Public Const MENOR As String = " < "
    Public Const ENTRE As String = " BETWEEN "

    Public Const MAS As String = " + "
    Public Const MENOS As String = " - "

    Public Const SUM As String = "SUM"
    Public Const COUNT As String = "COUNT"




    Public Sub Inicializa()
        Tipo_DateTime = System.Type.GetType("System.DateTime")
        Tipo_String = System.Type.GetType("System.String")
        Tipo_Int32 = System.Type.GetType("System.Int32")
        Tipo_Double = System.Type.GetType("System.Double")
        Tipo_Single = System.Type.GetType("System.Single")


    End Sub




End Module
Public Class BasedeDatos
    'Tiene el conjunto de ficheros (datos: .xml y esquema:.xsd) 
    'que componen la BD.

    'Normalmente debe ser solo uno (un par .xml y .xsd)
    'POR HACER:
    'Pero puede haber tablas que se deban poner solas ya que 
    'tienen un gran tamaño y no sea conveniente cargarlas
    'siempre en memoria

    Protected _ds As data.DataSet
    Protected _nombre As String
    Protected Friend _COLECCION_CONTENDORES As Coleccion_Contenedores

    Protected _tienecambiosdesdeultimaescritura As Boolean
    Public Enum TipoOrigendedatos
        XML
        Conexion
    End Enum



    Public RutaFicheroXML As String
    Public Const EXTENSION_XML As String = ".xml"
    Public Const EXTENSION_ESQUEMA As String = ".xsd"

    ReadOnly Property FicheroDatos() As String
        Get
            Return RutaFicheroXML + _nombre + EXTENSION_XML
        End Get
    End Property

    ReadOnly Property FicheroEsquema() As String
        Get
            Return RutaFicheroXML + _nombre + EXTENSION_ESQUEMA
        End Get
    End Property

    Property TieneCambiosDesdeUltimaEscritura() As Boolean
        Get
            Return _tienecambiosdesdeultimaescritura
        End Get
        Set(ByVal Value As Boolean)
            _tienecambiosdesdeultimaescritura = Value
        End Set
    End Property

    Public Sub New()
        Me.New("")
    End Sub
    Public Sub New(ByVal nombre As String)
        Me.New(nombre, Directorioaplicacion)
    End Sub

    Public Sub New(ByVal nombre As String, ByVal rutaficheros As String)
        _nombre = nombre
        RutaFicheroXML = rutaficheros
        _ds = New data.DataSet
        _tienecambiosdesdeultimaescritura = True 'Para forzar 
    End Sub
    Public Sub CreaEstructuraBasedeDatos()
    End Sub


    Public Sub VuelcaDatosAdisco()
        _ds.WriteXml(Me.FicheroDatos)
    End Sub

    Public Sub VuelcaEstructura()
        VuelcaDatosAdisco()
        _ds.WriteXmlSchema(Me.FicheroEsquema)
        _tienecambiosdesdeultimaescritura = False
    End Sub

    Public Function ExisteBD() As Boolean
        Return System.IO.File.Exists(RutaFicheroXML + _nombre + EXTENSION_ESQUEMA) And _
               System.IO.File.Exists(RutaFicheroXML + _nombre + EXTENSION_XML)
    End Function
    Public Sub Lee()
        If System.IO.File.Exists(RutaFicheroXML + _nombre + EXTENSION_ESQUEMA) Then
            _ds.ReadXmlSchema(RutaFicheroXML + _nombre + EXTENSION_ESQUEMA)

          
        Else
            Throw New Exception("No se encuentra el esquema:" & RutaFicheroXML + _nombre + EXTENSION_ESQUEMA)
        End If

        If System.IO.File.Exists(RutaFicheroXML + _nombre + EXTENSION_XML) Then
            Dim t As data.DataTable
            Dim strm As New System.IO.FileStream(RutaFicheroXML + _nombre + EXTENSION_XML, IO.FileMode.Open)
            Dim lector As System.Xml.XmlTextReader = New System.Xml.XmlTextReader(strm)

            For Each t In _ds.Tables
                t.BeginLoadData()
            Next

            _ds.ReadXml(lector, XmlReadMode.IgnoreSchema)

            For Each t In _ds.Tables
                t.EndLoadData()
            Next

            strm.Close()
            lector.Close()
        Else
            Throw New Exception("No se encuentra el fichero:" & RutaFicheroXML & _nombre & EXTENSION_XML)
        End If
        _tienecambiosdesdeultimaescritura = False
    End Sub
End Class


Public MustInherit Class Contenedor
    Enum origendelosregistros
        tabla = 1
        filtrado = 2
    End Enum

    Enum abrirpara
        lectura = 1
        escritura = 2
    End Enum
    Protected _nombrecontenedor As String
    Protected _id As Long
    Protected _proximoid As Long
    Public _tabla As data.DataTable
    Protected Friend _registroactual As data.DataRow
    Protected _pos As Integer
    Public Const CAMPO_timestamp As String = "t_s"
    Public Const CAMPO_ID As String = "id"
    Public Const N_CAMPO_timestamp As Integer = 0
    Public Const N_CAMPO_ID As Integer = 1
    Protected Const CADENAMAXIMOID As String = "MAX(" & CAMPO_ID & ")"
    Protected _t_s As DateTime
    Protected _modificado As Boolean
    Protected _add As Boolean
    Protected Friend _cacheado As Boolean
    Protected _relaciones As System.Data.DataRelationCollection

    Protected Const PRIMERA_POSICION As Integer = 0
    Protected Const POSICION_INDETERMINADA As Integer = -1

    Protected Friend _COLECCION_CAMPOS As Coleccion_Campos

    Protected Friend _bdpadre As BasedeDatos




    Public ReadOnly Property dirty() As Boolean
        Get
            Return _modificado Or _add
        End Get
    End Property
    Public ReadOnly Property SehaAñadidoRegistro() As Boolean
        Get
            Return _add
        End Get
    End Property

    Public Property timestamp() As DateTime
        Get
            If Me._cacheado Then
                Return _t_s
            Else
                Return CDate(_registroactual(N_CAMPO_timestamp))
            End If
        End Get
        Set(ByVal value As DateTime)
            Me.Pre_Set()
            _t_s = value
            _modificado = True
        End Set
    End Property
    Public ReadOnly Property id() As Long
        Get
            If Me._cacheado Then
                Return _id
            Else
                Return CLng(_registroactual(N_CAMPO_ID))
            End If
        End Get
    End Property

    Public Property Posicion() As Integer
        Get
            Return _pos
        End Get
        Set(ByVal Value As Integer)
            If (Value < Recordcount) And (Value > POSICION_INDETERMINADA) Then
                _pos = Value
                _registroactual = _tabla.Rows(_pos)
                _modificado = False
                _add = False
                _cacheado = False
            End If
        End Set
    End Property

    Public ReadOnly Property EOF() As Boolean
        Get
            Return (_pos = (Recordcount - 1))
        End Get
    End Property

    Public ReadOnly Property Vacio() As Boolean
        Get
            Return (Recordcount = 0)
        End Get
    End Property

    Public ReadOnly Property Recordcount() As Integer
        Get
            Return (_tabla.Rows.Count)
        End Get
    End Property

    Public ReadOnly Property NombreContenedor() As String
        Get
            Return _nombrecontenedor
        End Get
    End Property

    Public MustOverride ReadOnly Property AliasSql() As String




    Private Sub Preparanuevoregistro()
        _registroactual = _tabla.NewRow()
        _add = True
        _modificado = False
        _cacheado = True
        _id = POSICION_INDETERMINADA
    End Sub
    Public Sub NuevoRegistro(ByVal origen As Contenedor, Optional ByVal definitivo As Boolean = True)
        Me.Preparanuevoregistro()
        Me.Copia(origen)
        Me.Actualiza(definitivo)
    End Sub
    Public Sub Nuevoregistro(ByVal r As data.DataRow, ByVal definitivo As Boolean)
        Me.Preparanuevoregistro()
        Me.LeeCamposdeTabla(r)
        Me.Actualiza(definitivo)
    End Sub
    Public Sub NuevoRegistro()
        Me.Preparanuevoregistro()
        Me.ValoresPorDefecto()
    End Sub
    Public Sub BorraRegistro(ByVal definitivo As Boolean)
        'borramos el registro actual
        If _pos <> POSICION_INDETERMINADA And Not _add Then
            'Si lo esta añadiendo es que todavia no esta dentro,
            'por lo que no se puede borrar
            If Me.id = _proximoid - 1 Then
                _proximoid = _proximoid - 1
            End If
            _tabla.Rows.Remove(_registroactual)
            If definitivo Then
                _tabla.AcceptChanges()
            End If
            _pos = POSICION_INDETERMINADA
        End If
    End Sub
    Public Overridable Sub Copia(ByVal origen As Contenedor)
        'NO copiamos ni id ni el t_s, la copia tendrá uno nuevo
    End Sub

    Protected Function RegCamp(ByVal c As CampoEntero) As CampoEntero
        RegistraCampo(c)
        Return c
    End Function
    Protected Function RegCamp(ByVal c As CampoCadena) As CampoCadena
        RegistraCampo(c)
        Return c
    End Function

    Protected Function RegCamp(ByVal c As CampoCadena_Calculado) As CampoCadena_Calculado
        RegistraCampo(c)
        Return c
    End Function
    Protected Function RegCamp(ByVal c As CampoDouble) As CampoDouble
        RegistraCampo(c)
        Return c
    End Function

    Protected Function RegCamp(ByVal c As CampoSingle) As CampoSingle
        RegistraCampo(c)
        Return c
    End Function
    Protected Function RegCamp(ByVal c As CampoSingle_Calculado) As CampoSingle_Calculado
        RegistraCampo(c)
        Return c
    End Function
    Protected Function RegCamp(ByVal c As CampoFecha) As CampoFecha
        RegistraCampo(c)
        Return c
    End Function

    Protected Function RegCamp(ByVal c As CampoHora) As CampoHora
        RegistraCampo(c)
        Return c
    End Function

    Protected Function RegCamp(ByVal c As CampoBinario) As CampoBinario
        RegistraCampo(c)
        Return c
    End Function



    Protected Sub RegistraCampo(ByVal c As Campo)
        _COLECCION_CAMPOS.Add(c)
    End Sub

    Protected Friend Sub Pre_Set()
        If Not Me._cacheado Then
            Me._cacheado = True
            Me.Get_Campos()
        End If
    End Sub

    Public Sub Importa(ByVal objeto_externo As Contenedor)
        Dim i As Integer
        objeto_externo.MoveFirst()
        For i = 1 To objeto_externo.Recordcount
            Me.NuevoRegistro(objeto_externo, Util_Basededatos.ACT_NODEFINITIVO)
            objeto_externo.MoveNext()
        Next i
    End Sub

    Public Sub LeedeTabla(ByVal t As data.DataTable)
        Dim r As data.DataRow
        For Each r In t.Rows
            NuevoRegistro(r, Util_Basededatos.ACT_NODEFINITIVO)
        Next r
        Definitivo()
    End Sub
    Public Overridable Sub LeeCamposdeTabla(ByVal r As data.DataRow)
        'no leo  t_s
        If Not (r(CAMPO_ID) Is System.DBNull.Value) Then
            _id = CLng(r(CAMPO_ID))
            _proximoid = POSICION_INDETERMINADA
        End If
        For Each c As Campo In _COLECCION_CAMPOS
            c.LeeCampoDeTabla(r)
        Next
    End Sub
    Public Sub BorraTodosLosRegistros()
        'Borramos todos los registros
        While Me.Recordcount > 0
            Me.MoveFirst()
            Me.BorraRegistro(False)
        End While
        _registroactual = Nothing
    End Sub

    Public Function Funcionagregado(ByVal exp As String, ByVal filtro As String) As Long 'Por ahora solo sería un long
        Dim obj As Object
        obj = _tabla.Compute(exp, filtro)

        If obj Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(obj)
        End If
    End Function

    Public Function Funcionagregado_sng(ByVal exp As String, ByVal filtro As String) As Single
        Dim obj As Object
        obj = _tabla.Compute(exp, filtro)

        If obj Is System.DBNull.Value Then
            Return 0
        Else
            Return CSng(obj)
        End If
    End Function

    Public Sub MarcaModificado()
        _modificado = True
        'timestamp = Now()
    End Sub

    Public Function Existe_ID(ByVal idabuscar As Long) As Boolean
        Return Existe(CAMPO_ID & IGUAL & idabuscar.ToString)
    End Function

    Public Function Existe(ByRef expresion As String) As Boolean
        'Es como Posiciona, pero no cambia el estado del objeto actual
        Dim r As data.DataRow()
        r = busca(expresion)
        Return Not (r.Length = 0)
    End Function
    Public Function Posiciona_ID(ByVal r As data.DataRow) As Boolean
        Return Posiciona(Contenedor.CAMPO_ID & IGUAL & CStr(r(Contenedor.CAMPO_ID)))
    End Function

    Public Function Posiciona_ID(ByVal identificador As Long) As Boolean
        Return Posiciona(Contenedor.CAMPO_ID & IGUAL & identificador.ToString)
    End Function
    Public Function Posiciona(ByRef expresion As String) As Boolean
        'NO ES RECOMENDABLE, hacer obj.posiciona y despues movenext
        'Se movería al segundo de la tabla, que no tiene nada que ver
        'con la busqueda que se ha hecho 
        Dim r As data.DataRow()
        r = busca(expresion)
        If r.Length = 0 Then
            _pos = POSICION_INDETERMINADA
            Return False
        Else
            _registroactual = r(PRIMERA_POSICION) 'La primera de la selección actual
            _pos = PRIMERA_POSICION
            '¿que pasa con _pos? ¿en que posicion estamos?
            _modificado = False
            _add = False
            _cacheado = False
            'Me.Get_Campos()
            Return True
        End If
    End Function
    Public Function busca(ByRef expresion As String) As data.DataRow()
        Dim dr() As data.DataRow

        dr = _tabla.Select(expresion)
        If dr.Length = 0 Then
            Return New data.DataRow() {}
        Else
            Return dr
        End If

    End Function
    Public Sub New(ByVal tabla As data.DataTable, ByVal bdpadre As BasedeDatos)
        _nombrecontenedor = tabla.TableName
        _tabla = tabla
        _registroactual = Nothing
        _proximoid = POSICION_INDETERMINADA ' Obligamos a llamar a CalculaProximoID() en la proxima insercion
        _pos = POSICION_INDETERMINADA
        _relaciones = _tabla.DataSet.Relations()

        _COLECCION_CAMPOS = New Coleccion_Campos
        _bdpadre = bdpadre
    End Sub

    Public Sub Actualiza(Optional ByVal definitivo As Boolean = True)
        Dim m As Boolean
        Dim a As Boolean
        m = _modificado
        a = _add
        'Cambian al hacer Set_campos        
        If m Or a Then
            Me.Set_Campos()
            If a Then
                _tabla.Rows.Add(_registroactual)
            End If
            If definitivo Then
                _tabla.AcceptChanges()
            End If
            _bdpadre.TieneCambiosDesdeUltimaEscritura = True
        End If
    End Sub

    Public Sub Definitivo()
        _tabla.AcceptChanges()
        _bdpadre.TieneCambiosDesdeUltimaEscritura = True
    End Sub

    Protected Friend Overridable Sub Get_Campos()
        _t_s = CDate(_registroactual(N_CAMPO_timestamp))
        _id = CLng(_registroactual(N_CAMPO_ID))

        For Each c As Campo In _COLECCION_CAMPOS
            c.Get_Campo()
        Next

        Me._cacheado = True
        Me._modificado = False
        Me._add = False
    End Sub

    Protected Overridable Sub Set_Campos()
        If _add Then
            If _id = POSICION_INDETERMINADA Then
                If _proximoid = POSICION_INDETERMINADA Then
                    _id = CalculaProximoID()
                Else
                    _id = _proximoid
                End If
                _proximoid = _id + 1
            End If
        End If
        _t_s = Now()
        _registroactual(CAMPO_timestamp) = _t_s
        _registroactual(CAMPO_ID) = _id
        _modificado = False
        _add = False
        _cacheado = True
        For Each c As Campo In _COLECCION_CAMPOS
            c.Set_Campo()
        Next
    End Sub
    Protected Overridable Sub ValoresPorDefecto()
        _id = POSICION_INDETERMINADA
        'No le ponemos valores a _t_s ya que a la hora de Actualizar es cuando se establece la propiedad
        For Each c As Campo In _COLECCION_CAMPOS
            c.ValorPorDefecto()
        Next
    End Sub

    Public Overridable Function RegistrosRelacionados(ByVal nombre As String) As data.DataRow()
        If Not _registroactual Is Nothing And _pos <> POSICION_INDETERMINADA Then
            Dim dr() As data.DataRow
            dr = _registroactual.GetChildRows(nombre)
            If dr Is Nothing Then
                Return New data.DataRow() {}
            Else
                Return dr
            End If
        Else
            Return New data.DataRow() {}
        End If
    End Function


    Public Sub CreaEstructura()
        Dim dc(0) As DataColumn
        Dim col As System.Data.DataColumnCollection

        col = _tabla.Columns

        col.Add(CAMPO_timestamp, Util_Basededatos.Tipo_DateTime)
        col.Add(CAMPO_ID, Util_Basededatos.Tipo_Int32)
        dc(0) = col(N_CAMPO_ID)
        _tabla.PrimaryKey = dc

        For Each c As Campo In _COLECCION_CAMPOS
            c.CreaEstructura(col)
        Next

    End Sub

    Public Sub MoveFirst()
        Posicion = PRIMERA_POSICION
    End Sub
    Public Sub MoveLast()
        Posicion = Recordcount - 1
    End Sub

    Public Sub MoveNext()
        Posicion += 1
    End Sub

    Public Sub MovePrevious()
        Posicion -= 1
    End Sub

    Private Function CalculaProximoID() As Long

        Return Funcionagregado(CADENAMAXIMOID, "") + 1
    End Function

#If Not COMPACT_FRAMEWORk Then
    Protected Sub Importacion(ByVal conex As ConexionSQL, ByVal borrarantes As Boolean)

        Dim i As Integer
        Me.MoveFirst()
        If borrarantes AndAlso Me.Recordcount > 0 Then
            conex.Ejecuta(Util.SQL_DELETE_FROM & Me.AliasSql)
        End If

        For i = 1 To Me.Recordcount
            EscribeRegistroEnSql(conex, False)
            Me.MoveNext()
        Next i


    End Sub


    Public Sub EscribeRegistroEnSql(ByVal conex As ConexionSQL, ByVal borraranterior As Boolean)
        Dim borrar As New delete_sql
        Dim insertar As New insert_sql

        Dim parametros As New Collection
        Dim par As data.SqlClient.SqlParameter
        Dim comando As data.SqlClient.SqlCommand

        insertar.Limpia()
        insertar.Tabla = Me.AliasSql

        If borraranterior Then
            With borrar
                .Tabla = Me.AliasSql
                .add_condicion(Contenedor.CAMPO_ID, Me.id)
                conex.Ejecuta(.texto_sql)
            End With
        End If

        For Each c As Campo In _COLECCION_CAMPOS
            insertar.add_valor_expr(c._CAMPO, c.ValorparaSQL)
            If c.ValorparaSQL Like CampoBinario.PrefijoPARAMETROBINARIO & ASTERISCO Then
                Dim datos() As Byte
                datos = DirectCast(c, CampoBinario).Valor
                par = New Data.SqlClient.SqlParameter(c.ValorparaSQL, Data.SqlDbType.Binary, datos.Length, Data.ParameterDirection.Input, False, 0, 0, Nothing, Data.DataRowVersion.Current, datos)
                parametros.Add(par)
            End If
        Next
        comando = conex.ObtenSqlCommand(insertar.texto_sql, Data.CommandType.Text)
        conex.AddParametros(comando, parametros)
        conex.Ejecuta(comando)

        Util.VaciaColleccion(parametros)

    End Sub

    Public Overridable Sub ImportaDesdeHost(ByVal conex As ConexionSQL)
        Me.Importacion(conex, False)
    End Sub
#End If


End Class



Public MustInherit Class Campo
    Protected Friend _contenedor As Contenedor
    Protected Friend _N_CAMPO As Integer
    Protected Friend _CAMPO As String

    Protected Friend _expr As String

    Public ReadOnly Property Expresion() As String
        Get
            Return _expr
        End Get
    End Property

    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String, ByVal expresion As String)
        _contenedor = padre
        _N_CAMPO = numerocampo
        _CAMPO = nombrecampo
        _expr = expresion
    End Sub


    Protected Friend MustOverride Sub Set_Campo()
    Protected Friend MustOverride Sub Get_Campo()
    Protected Friend MustOverride Sub ValorPorDefecto()
    Protected Friend MustOverride Sub LeeCampoDeTabla(ByRef r As Data.DataRow)
    Protected Friend MustOverride Sub CreaEstructura(ByRef col As Data.DataColumnCollection)

    Public MustOverride Function ValorparaSQL() As String

    Public MustOverride Sub Copia(ByVal c As Campo)

    Public MustOverride Function ValorParaInterfaz() As String


End Class




Public Class CampoEntero
    Inherits Campo

    Protected _valor As Long
    Const CTE_NULA As Integer = ZERO
    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String)
        MyBase.new(padre, numerocampo, nombrecampo, "")
    End Sub

    Public Property Valor() As Long
        Get
            If _contenedor._cacheado Then
                Return _valor
            Else
                If _contenedor._registroactual Is Nothing Then
                    Return CTE_NULA
                Else
                    If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                        Return CTE_NULA
                    Else
                        Return CLng(_contenedor._registroactual(_N_CAMPO))
                    End If
                End If
            End If
        End Get

        Set(ByVal Value As Long)
            _contenedor.Pre_Set()
            _valor = Value
            _contenedor.MarcaModificado()
        End Set
    End Property

    Public Overrides Function ValorParaInterfaz() As String
        Return Me.Valor.ToString
    End Function

    Protected Friend Overrides Sub Set_Campo()
        If _contenedor._cacheado Then
            _contenedor._registroactual(_N_CAMPO) = _valor
        Else
            If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                _contenedor._registroactual(_N_CAMPO) = CTE_NULA
            End If
        End If
    End Sub

    Protected Friend Overrides Sub Get_Campo()
        _valor = CLng(_contenedor._registroactual(_N_CAMPO))
    End Sub

    Protected Friend Overrides Sub CreaEstructura(ByRef col As DataColumnCollection)
        col.Add(_CAMPO, Util_Basededatos.Tipo_Int32)
    End Sub

    Protected Friend Overrides Sub ValorPorDefecto()
        _valor = CTE_NULA
    End Sub

    Protected Friend Overrides Sub LeeCampoDeTabla(ByRef r As data.DataRow)
        If Not (r(_CAMPO) Is System.DBNull.Value) Then Me.Valor = CInt(r(_CAMPO))
    End Sub

    Public Overrides Function ValorparaSQL() As String
        Return Me.Valor.ToString
    End Function

    Public Overrides Sub Copia(ByVal c As Campo)
        Me.Valor = DirectCast(c, CampoEntero).Valor
    End Sub
End Class

Public Class CampoCadena
    Inherits Campo

    Protected _valor As String
    Const CTE_NULA As String = CADENAVACIA
    Public _tammaximo As Integer

    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String, ByVal tammaximo As Integer)
        MyBase.new(padre, numerocampo, nombrecampo, "")
        _tammaximo = tammaximo
    End Sub

    Public Property Valor() As String
        Get
            If _contenedor._cacheado Then
                Return _valor
            Else
                If _contenedor._registroactual Is Nothing Then
                    Return CTE_NULA
                Else
                    If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                        Return CTE_NULA
                    Else
                        Return CStr(_contenedor._registroactual(_N_CAMPO))
                    End If
                End If
            End If
        End Get

        Set(ByVal Value As String)
            _contenedor.Pre_Set()
            _valor = Value
            _contenedor.MarcaModificado()
        End Set
    End Property

    Public Overrides Function ValorParaInterfaz() As String
        Return Me.Valor
    End Function
    Protected Friend Overrides Sub Set_Campo()
        If _contenedor._cacheado Then
            _contenedor._registroactual(_N_CAMPO) = _valor
        Else
            If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                _contenedor._registroactual(_N_CAMPO) = CTE_NULA
            End If
        End If
    End Sub

    Protected Friend Overrides Sub Get_Campo()
        _valor = CStr(_contenedor._registroactual(_N_CAMPO))
    End Sub

    Protected Friend Overrides Sub CreaEstructura(ByRef col As DataColumnCollection)
        col.Add(_CAMPO, Util_Basededatos.Tipo_String)
    End Sub

    Protected Friend Overrides Sub ValorPorDefecto()
        _valor = CTE_NULA
    End Sub

    Protected Friend Overrides Sub LeeCampoDeTabla(ByRef r As data.DataRow)
        If Not (r(_CAMPO) Is System.DBNull.Value) Then Me.Valor = CStr(r(_CAMPO))
    End Sub

    Public Overrides Function ValorparaSQL() As String
        Return Util.v(Left(Me.Valor, _tammaximo))
    End Function

    Public Overrides Sub Copia(ByVal c As Campo)
        Me.Valor = DirectCast(c, CampoCadena).Valor
    End Sub



End Class

Public Class CampoCadena_Calculado
    Inherits Campo

    Const CTE_NULA As String = CADENAVACIA

    Public _tammaximo As Integer
    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String, ByVal expresion As String, ByVal tammaximo As Integer)
        MyBase.new(padre, numerocampo, nombrecampo, expresion)

        _tammaximo = tammaximo
    End Sub

    Public Shadows ReadOnly Property Valor() As String
        Get
            If _contenedor._registroactual Is Nothing Then
                Return CTE_NULA
            Else
                If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                    Return CTE_NULA
                Else
                    Return CStr(_contenedor._registroactual(_N_CAMPO))
                End If
            End If
        End Get
    End Property

    Public Overrides Function ValorParaInterfaz() As String
        Return Me.Valor
    End Function
    Protected Friend Overrides Sub CreaEstructura(ByRef col As DataColumnCollection)
        Dim dc As System.Data.DataColumn

        dc = col.Add(_CAMPO, Util_Basededatos.Tipo_String)
        dc.Expression = _expr
    End Sub
    Protected Friend Overrides Sub Set_Campo()
    End Sub

    Protected Friend Overrides Sub Get_Campo()
    End Sub
    Protected Friend Overrides Sub ValorPorDefecto()
    End Sub
    Protected Friend Overrides Sub LeeCampoDeTabla(ByRef r As data.DataRow)
    End Sub

    Public Overrides Function ValorparaSQL() As String
        Return v(Left(Me.Valor, _tammaximo))
    End Function

    Public Overrides Sub Copia(ByVal c As Campo)
        'No hacemos nada, es calculado
    End Sub
End Class

Public Class CampoSingle
    Inherits Campo

    Protected _valor As Single
    Const CTE_NULA As Single = ZERO_SNG
    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String)
        MyBase.new(padre, numerocampo, nombrecampo, "")
    End Sub

    Public Property Valor() As Single
        Get
            If _contenedor._cacheado Then
                Return _valor
            Else
                If _contenedor._registroactual Is Nothing Then
                    Return CTE_NULA
                Else
                    If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                        Return CTE_NULA
                    Else
                        Return CSng(_contenedor._registroactual(_N_CAMPO))
                    End If
                End If
            End If
        End Get

        Set(ByVal Value As Single)
            _contenedor.Pre_Set()
            _valor = Value
            _contenedor.MarcaModificado()
        End Set
    End Property

    Public Overrides Function ValorParaInterfaz() As String
        Return Format(Me.Valor, "##,00")
    End Function
    Protected Friend Overrides Sub Set_Campo()
        If _contenedor._cacheado Then
            _contenedor._registroactual(_N_CAMPO) = _valor
        Else
            If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                _contenedor._registroactual(_N_CAMPO) = CTE_NULA
            End If
        End If
    End Sub

    Protected Friend Overrides Sub Get_Campo()
        _valor = CSng(_contenedor._registroactual(_N_CAMPO))
    End Sub

    Protected Friend Overrides Sub CreaEstructura(ByRef col As DataColumnCollection)
        col.Add(_CAMPO, Util_Basededatos.Tipo_Single)
    End Sub

    Protected Friend Overrides Sub ValorPorDefecto()
        _valor = CTE_NULA
    End Sub

    Protected Friend Overrides Sub LeeCampoDeTabla(ByRef r As data.DataRow)
        If Not (r(_CAMPO) Is System.DBNull.Value) Then Me.Valor = CSng(r(_CAMPO))
    End Sub


    Public Overrides Function ValorparaSQL() As String
        Dim separadordecimal As String
        Dim separadormiles As String
        Dim s As String

        separadordecimal = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator()
        separadormiles = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator()

        s = Me.Valor.ToString.Trim
        If s = "" Then
            Return "0.0"
        Else
            If InStr(s, COMA) > 0 And InStr(s, PUNTO) > 0 Then
                'tenemos los dos. Lo dejamos igual
                Return s
            Else
                Return Replace(Replace(s, separadormiles, ""), separadordecimal, PUNTO)
            End If
        End If
    End Function

    Public Overrides Sub Copia(ByVal c As Campo)
        Me.Valor = DirectCast(c, CampoSingle).Valor
    End Sub


End Class


Public Class CampoSingle_Calculado
    Inherits Campo

    Const CTE_NULA As Single = ZERO_SNG
    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String, ByVal expresion As String)
        MyBase.new(padre, numerocampo, nombrecampo, expresion)
    End Sub

    Public Shadows ReadOnly Property Valor() As Single
        Get
            If _contenedor._registroactual Is Nothing Then
                Return CTE_NULA
            Else
                If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                    Return CTE_NULA
                Else
                    Return CSng(_contenedor._registroactual(_N_CAMPO))
                End If
            End If
        End Get
    End Property

    Public Overrides Function ValorParaInterfaz() As String
        Return Format(Me.Valor, "##,0")
    End Function
    Protected Friend Overrides Sub CreaEstructura(ByRef col As DataColumnCollection)
        Dim dc As System.Data.DataColumn

        dc = col.Add(_CAMPO, Util_Basededatos.Tipo_Single)
        dc.Expression = _expr
    End Sub

    Protected Friend Overrides Sub Set_Campo()
    End Sub

    Protected Friend Overrides Sub Get_Campo()

    End Sub
    Protected Friend Overrides Sub ValorPorDefecto()

    End Sub
    Protected Friend Overrides Sub LeeCampoDeTabla(ByRef r As data.DataRow)

    End Sub

    Public Overrides Function ValorparaSQL() As String
        Dim separadordecimal As String
        Dim separadormiles As String
        Dim s As String

        separadordecimal = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator()
        separadormiles = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator()

        s = Me.Valor.ToString.Trim
        If s = "" Then
            Return "0.0"
        Else
            If InStr(s, COMA) > 0 And InStr(s, PUNTO) > 0 Then
                'tenemos los dos. Lo dejamos igual
                Return s
            Else
                Return Replace(Replace(s, separadormiles, ""), separadordecimal, PUNTO)
            End If
        End If
    End Function

    Public Overrides Sub Copia(ByVal c As Campo)
        'No hacemos nada es calculado
    End Sub
End Class

Public Class CampoDouble
    Inherits Campo

    Protected _valor As Double
    Const CTE_NULA As Double = ZERO_SNG
    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String)
        MyBase.new(padre, numerocampo, nombrecampo, "")
    End Sub

    Public Property Valor() As Double
        Get
            If _contenedor._cacheado Then
                Return _valor
            Else
                If _contenedor._registroactual Is Nothing Then
                    Return CTE_NULA
                Else
                    If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                        Return CTE_NULA
                    Else
                        Return CDbl(_contenedor._registroactual(_N_CAMPO))
                    End If
                End If
            End If
        End Get

        Set(ByVal Value As Double)
            _contenedor.Pre_Set()
            _valor = Value
            _contenedor.MarcaModificado()
        End Set
    End Property
    Public Overrides Function ValorParaInterfaz() As String
        Return Format(Me.Valor, "##,00")
    End Function
    Protected Friend Overrides Sub Set_Campo()
        If _contenedor._cacheado Then
            _contenedor._registroactual(_N_CAMPO) = _valor
        Else
            If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                _contenedor._registroactual(_N_CAMPO) = CTE_NULA
            End If
        End If
    End Sub

    Protected Friend Overrides Sub Get_Campo()
        _valor = CSng(_contenedor._registroactual(_N_CAMPO))
    End Sub

    Protected Friend Overrides Sub CreaEstructura(ByRef col As DataColumnCollection)
        col.Add(_CAMPO, Util_Basededatos.Tipo_Double)
    End Sub

    Protected Friend Overrides Sub ValorPorDefecto()
        _valor = CTE_NULA
    End Sub

    Protected Friend Overrides Sub LeeCampoDeTabla(ByRef r As data.DataRow)
        If Not (r(_CAMPO) Is System.DBNull.Value) Then Me.Valor = CDbl(r(_CAMPO))
    End Sub

    Public Overrides Function ValorparaSQL() As String
        Dim separadordecimal As String
        Dim separadormiles As String
        Dim s As String

        separadordecimal = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator()
        separadormiles = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator()

        s = Me.Valor.ToString.Trim
        If s = "" Then
            Return "0.0"
        Else
            If InStr(s, COMA) > 0 And InStr(s, PUNTO) > 0 Then
                'tenemos los dos. Lo dejamos igual
                Return s
            Else
                Return Replace(Replace(s, separadormiles, ""), separadordecimal, PUNTO)
            End If
        End If
    End Function

    Public Overrides Sub Copia(ByVal c As Campo)
        Me.Valor = DirectCast(c, CampoDouble).Valor
    End Sub
End Class


Public MustInherit Class CampoFechaBase
    Inherits Campo

    Protected _valor As DateTime




    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String)
        MyBase.new(padre, numerocampo, nombrecampo, "")
    End Sub

    Public Property Valor() As DateTime
        Get
            If _contenedor._cacheado Then
                Return _valor
            Else
                If _contenedor._registroactual Is Nothing Then
                    Return Now()
                Else
                    If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                        Return Now()
                    Else
                        Return CDate(_contenedor._registroactual(_N_CAMPO))
                    End If
                End If
            End If
        End Get

        Set(ByVal Value As DateTime)
            _contenedor.Pre_Set()
            _valor = Value
            _contenedor.MarcaModificado()
        End Set
    End Property

    Protected Friend Overrides Sub Set_Campo()
        If _contenedor._cacheado Then
            _contenedor._registroactual(_N_CAMPO) = _valor
        Else
            If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                _contenedor._registroactual(_N_CAMPO) = Now()
            End If
        End If
    End Sub

    Protected Friend Overrides Sub Get_Campo()
        _valor = CDate(_contenedor._registroactual(_N_CAMPO))
    End Sub

    Protected Friend Overrides Sub CreaEstructura(ByRef col As DataColumnCollection)
        col.Add(_CAMPO, Util_Basededatos.Tipo_DateTime)
    End Sub

    Protected Friend Overrides Sub ValorPorDefecto()
        _valor = New DateTime(2001, 1, 1, 0, 0, 0, 0)
    End Sub
    Protected Friend Overrides Sub LeeCampoDeTabla(ByRef r As data.DataRow)
        If Not (r(_CAMPO) Is System.DBNull.Value) Then Me.Valor = CDate(r(_CAMPO))
    End Sub

    Public Overrides Sub Copia(ByVal c As Campo)
        Me.Valor = DirectCast(c, CampoFechaBase).Valor
    End Sub

End Class

Public Class CampoFecha
    Inherits CampoFechaBase

    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String)
        MyBase.new(padre, numerocampo, nombrecampo)
    End Sub

    Public Overrides Function ValorparaSQL() As String
        Return Util.fechasql(Me.Valor)
    End Function

    Public Overrides Function ValorParaInterfaz() As String
        Return Format(Me.Valor, "dd/MM/yyyy")
    End Function
End Class

Public Class CampoHora
    Inherits CampoFechaBase

    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String)
        MyBase.new(padre, numerocampo, nombrecampo)
    End Sub

    Public Overrides Function ValorparaSQL() As String
        Return Util.HoraSQL(Me.Valor)
    End Function

    Public Overrides Function ValorParaInterfaz() As String
        Return Format(Me.Valor, "HH:mm:ss")
    End Function
End Class

Public Class Coleccion_Contenedores
    Inherits CollectionBase

    Public Sub Add(ByVal c As Campo)
        List.Add(c)
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index < Count And index >= 0 Then
            List.RemoveAt(index)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Contenedor
        Get
            Return CType(List.Item(index), Contenedor)
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal nombre As String) As Contenedor
        Get
            'buscamos el nombre            
            For Each c As Contenedor In Me
                If c.NombreContenedor = nombre Then
                    Return c
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class

Public Class Coleccion_Campos
    Inherits CollectionBase

    Public Sub Add(ByVal c As Campo)
        List.Add(c)
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index < Count And index >= 0 Then
            List.RemoveAt(index)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Campo
        Get
            Return CType(List.Item(index), Campo)
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal nombre As String) As Campo
        Get
            'buscamos el nombre            
            For Each c As Campo In Me
                If c._CAMPO = nombre Then
                    Return c
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class

Public Class CampoBinario
    Inherits Campo
    Protected _valor() As Byte
    Private CTE_NULA As Byte() = {}



    Public Const PrefijoPARAMETROBINARIO As String = PrefijoPARAMETRO & "_Binario"

    Public Sub New(ByVal padre As Contenedor, ByVal numerocampo As Integer, ByVal nombrecampo As String)
        MyBase.new(padre, numerocampo, nombrecampo, "")
    End Sub

    Public Property Valor() As Byte()
        Get

            If _contenedor._cacheado Then
                Return _valor
            Else
                If _contenedor._registroactual Is Nothing Then
                    Return CTE_NULA
                Else
                    If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                        Return CTE_NULA
                    Else
                        Return DirectCast(_contenedor._registroactual(_N_CAMPO), Byte())
                    End If
                End If
            End If
        End Get

        Set(ByVal Value As Byte())
            _contenedor.Pre_Set()
            _valor = Value
            _contenedor.MarcaModificado()
        End Set
    End Property

    Public Overrides Sub Copia(ByVal c As Campo)
        Me.Valor = DirectCast(c, CampoBinario).Valor
    End Sub

    Protected Friend Overrides Sub CreaEstructura(ByRef col As System.Data.DataColumnCollection)
        col.Add(_CAMPO, CTE_NULA.GetType)
    End Sub

    Protected Friend Overrides Sub Get_Campo()
        _valor = DirectCast(_contenedor._registroactual(_N_CAMPO), Byte())
    End Sub

    Protected Friend Overrides Sub LeeCampoDeTabla(ByRef r As System.Data.DataRow)
        If Not (r(_CAMPO) Is System.DBNull.Value) Then Me.Valor = DirectCast(r(_CAMPO), Byte())
    End Sub

    Protected Friend Overrides Sub Set_Campo()
        If _contenedor._cacheado Then
            _contenedor._registroactual(_N_CAMPO) = _valor
        Else
            If (_contenedor._registroactual(_N_CAMPO) Is System.DBNull.Value) Then
                _contenedor._registroactual(_N_CAMPO) = CTE_NULA
            End If
        End If
    End Sub

    Public Overrides Function ValorparaSQL() As String
        'Dim fin As Integer
        'Dim s As String
        'Dim result As String

        'fin = UBound(_valor)
        'For i As Integer = 0 To fin
        '    s = Hex(_valor(i))
        '    If Len(s) = 1 Then
        '        s = s & "0"
        '    End If
        '    result &= s
        'Next
        'Return "0x" & result

        'Return "0x" 'Binario vacio. Despues actualizamos en Ejecuta_actulizabinary

        Return PrefijoPARAMETROBINARIO & Me._CAMPO

    End Function

    Protected Friend Overrides Sub ValorPorDefecto()
        _valor = CTE_NULA
    End Sub

    Public Overrides Function ValorParaInterfaz() As String
        Return "#DATOS BINARIOS#"
    End Function
End Class








