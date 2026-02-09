Option Explicit On
Option Strict On
Imports System.Text.RegularExpressions
Imports System.Data
Public Module Parser_TPR
    Public Function LeeParametros(ByVal param As String) As ColecciondeTokens
        ' Return SplitCadbase(t_coma, Util.SQL_SustCadenas(param), param)
        Return SplitCadbase(Util.COMA, Util.SQL_SustCadenas(param), param)
    End Function

    Public Function ExtraeParametroEntero(ByVal s As String) As Long
        Try
            Return CLng(s)
        Catch ex As Exception
            Throw New Exception("Se esperaba un entero: " & s)
        End Try
    End Function


    Public Function ExtraeParametroCadena(ByVal s As String, ByVal constantes As Coleccion_Constantes) As String
        s = s.Trim
        If s.StartsWith("'"c) And s.EndsWith("'"c) Then
            Return constantes.Reemplaza(s.Substring(1, s.Length - 2))
        Else
            Throw New Exception("Se esperaba una cadena: " & s)
        End If
    End Function

    Public Function ExtraeParametroHexadecimal(ByVal s As String) As String
        If Not Regex.IsMatch(s, "^([0-9A-F][0-9A-F])*$") Then
            Throw New Exception("Se esperaba una cadena Hexadecimal (sin comillas y siempre con dos digitos ej: 0A0D)  " & s)
        End If

        Return s
    End Function
 
    Public Function ExtraeParametroCaracter(ByVal s As String) As String
        s = s.Trim
        If s.StartsWith("'"c) And s.EndsWith("'"c) Then
            s = s.Substring(1, s.Length - 2)
            If s.Length <= 1 Then
                Return s
            Else
                Throw New Exception("Se esperaba un caracter: " & s)
            End If
        Else
            Throw New Exception("Se esperaba un caracter: " & s)
        End If
    End Function

    Public Function ExtraeParametrobooleano(ByVal s As String) As Boolean
        Select Case s.Trim.ToUpper
            Case "S"
                Return True
            Case "N"
                Return False
            Case Else
                Throw New Exception("Se esperaba un valor booleno (S/N): " & s)
        End Select

    End Function

    Public Function ExtraeParametroFlotante(ByVal s As String) As Single
        'Viene con punto decimal, sin separadores de miles
        Dim p, p1 As Integer
        If Regex.IsMatch(s, "^[0-9]+.[0-9]+$") Then
            p = s.IndexOf("."c)
            p1 = p + 1
            Return CSng(s.Substring(0, p)) + CSng(s.Substring(p1)) / CSng(Math.Pow(10, s.Length - p1))
        Else
            Throw New Exception("Se espera un flotante (con separador decimal el punto y sin separador de miles)")
        End If
    End Function

    Public Function ExtraeParametroFecha(ByVal s As String) As Date
        If Regex.IsMatch(s, "^[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]$") Then
            Return New Date(CInt(s.Substring(6, 4)), CInt(s.Substring(3, 2)), CInt(s.Substring(0, 2)))
        ElseIf Regex.IsMatch(s, "^[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]$") Then
            Return New Date(CInt(s.Substring(6, 4)), CInt(s.Substring(3, 2)), CInt(s.Substring(0, 2)), CInt(s.Substring(11, 2)), CInt(s.Substring(14, 2)), CInt(s.Substring(17, 2)))
        Else
            Throw New Exception("Se espera una fecha (dd/MM/yyyy ó dd/MM/yyyy hh:mm:ss)")

        End If
    End Function

    Public Function SplitCadbase(ByVal separador As String, ByVal cadbase As String, ByVal cad As String) As ColecciondeTokens
        'funciona igual que split, pero usamos una cadena base que hemos modificado
        Dim partes(), s As String
        Dim col As New ColecciondeTokens
        Dim tn As Token_Nodo
        Dim pos As Integer



        partes = Split(cadbase, separador)
        pos = 0

        For Each s In partes
            tn = New Token_Nodo()
            tn.TOKEN = separador

            tn.contenido = cad.Substring(pos, s.Length)
            tn.posicion = pos
            col.Add(pos.ToString, tn)
            pos += s.Length + separador.Length
        Next

        Return col




    End Function

    Public Function SplitPrefijoMultiple(ByVal tokens() As String, ByVal cad As String) As ColecciondeTokens
        Return SPlitPrefijoMultiple(tokens, Util.SQL_SustCadenas(cad), cad)
    End Function
    Public Function SplitPrefijoMultiple(ByVal tokens() As String, ByVal cadbase As String, ByVal cad As String) As ColecciondeTokens

        Const NOEXISTE As Integer = -1
        Dim pos As Integer
        Dim resultado, cs, listaelegida As New ColecciondeTokens
        Dim conjunto_parcial As New ColdeColdeTokens

        Dim ds As Token_Nodo
        Dim t As String
        Dim i, il As Integer
        Dim quedaalguno As Boolean

        Dim numerotokens As Integer
        Dim minimo As Integer
        Dim posbase As Integer
        Dim s As String

        numerotokens = tokens.Length

        For Each t In tokens
            pos = cadbase.IndexOf(t)
            cs = New ColecciondeTokens
            While pos <> NOEXISTE
                'marcamos la posicion
                ds = New Token_Nodo
                ds.TOKEN = t
                ds.posicion = pos
                cs.Add(pos.ToString, ds)
                pos = cadbase.IndexOf(t, pos + 1)
            End While
            conjunto_parcial.Add(cs)
        Next

        'Ahora reordenamos 
        il = conjunto_parcial.Count - 1

        quedaalguno = True
        While quedaalguno
            quedaalguno = False
            minimo = -1
            ds = Nothing
            listaelegida = Nothing
            For i = 0 To il 'recorremos las listas de todos los tokens
                cs = conjunto_parcial.Item(i)
                If cs.Count > 0 Then
                    ds = cs.Item(0)
                    If minimo = -1 Or minimo > ds.posicion Then
                        listaelegida = cs
                        minimo = ds.posicion
                    End If
                    quedaalguno = True
                End If
            Next
            If minimo > -1 Then
                resultado.Add(minimo.ToString, listaelegida.Item(0))
                listaelegida.Remove(0)
            End If
        End While
        conjunto_parcial.Clear()

        il = resultado.Count - 1
        For i = 0 To il
            ds = resultado.Item(i)
            posbase = ds.posicion + ds.TOKEN.Length
            If i = il Then   'No hay siguiente
                s = cad.Substring(posbase)
            Else 'existe un siguiente
                s = cad.Substring(posbase, resultado.Item(i + 1).posicion - posbase)
            End If
            ds.contenido = s
        Next

        Return resultado
    End Function
    Public Function Minimo_Multiple(ByVal ParamArray valores() As Integer) As Integer
        'SOLO NUMEROS POSITIVOS, -1 si la matriz no tiene valores
        Dim i As Integer
        Dim f_i As Integer = valores.Length - 1
        Dim minimo As Integer = -1
        Dim encontrado_alguno As Boolean = False
        Dim z As Integer


        For i = 0 To f_i
            z = valores(i)
            If z > -1 Then
                If encontrado_alguno Then
                    If z < minimo Then
                        minimo = z
                    End If
                Else
                    minimo = z
                    encontrado_alguno = True
                End If
            End If
        Next


        Return minimo
    End Function

    Const ptr_NUMERO As String = "[0-9]"
    Const ptr_NUMERO2 As String = ptr_NUMERO & ptr_NUMERO
    Const PATRON_IDUNICO As String = "\$IDUNICO" & ptr_NUMERO2 & "\$"

    Public Function NombreArchivoTestigo(ByVal nombrearchivo As String, ByVal exttestigo As String) As String
        Return System.IO.Path.GetFileNameWithoutExtension(nombrearchivo) & "." & exttestigo
    End Function
    Public Function BuscaFicheroParaLectura(ByVal ruta As String, ByVal ficherotestigo As Boolean, ByVal exttestigo As String) As ColecciondeString
        'devuelve una lista de los ficheros encontrados
        Dim resultado As New ColecciondeString
        Dim directorio, nombrefichero, etiquetaidunico As String
        Dim testigoencontrado As Boolean
        Dim patronbusqueda As String
        Dim longitudid, i As Integer
        Dim str_numeroexistente As String
        Dim ptr_numeroN As String = String.empty
        Dim nombretestigo As String

        directorio = System.IO.Path.GetDirectoryName(ruta)
        nombrefichero = System.IO.Path.GetFileName(ruta)

        etiquetaidunico = Regex.Match(nombrefichero, PATRON_IDUNICO).Value

        If etiquetaidunico <> "" Then
            longitudid = CInt(Regex.Match(etiquetaidunico, ptr_NUMERO2).Value)
            For i = 1 To longitudid
                ptr_numeroN &= ptr_NUMERO
            Next
            If ficherotestigo Then
                nombretestigo = NombreArchivoTestigo(nombrefichero, exttestigo)

                patronbusqueda = nombretestigo.Replace(etiquetaidunico, New String("?"c, longitudid))
                For Each ficheroencontrado As String In My.Computer.FileSystem.GetFiles(directorio, FileIO.SearchOption.SearchTopLevelOnly, patronbusqueda)
                    testigoencontrado = False
                    If System.IO.Path.GetFileName(ficheroencontrado).Length = patronbusqueda.Length Then
                        If Regex.Match(ficheroencontrado, ptr_numeroN).Value <> "" Then
                            testigoencontrado = True
                        End If
                    End If

                    If testigoencontrado Then
                        ficheroencontrado = directorio & System.IO.Path.DirectorySeparatorChar & System.IO.Path.GetFileNameWithoutExtension(ficheroencontrado) & System.IO.Path.GetExtension(nombrefichero)
                        If System.IO.File.Exists(ficheroencontrado) Then
                            resultado.Add(ficheroencontrado)
                        End If
                    End If
                Next
            Else
                patronbusqueda = nombrefichero.Replace(etiquetaidunico, New String("?"c, longitudid))
                For Each ficheroencontrado As String In My.Computer.FileSystem.GetFiles(directorio, FileIO.SearchOption.SearchTopLevelOnly, patronbusqueda)
                    If System.IO.Path.GetFileName(ficheroencontrado).Length = patronbusqueda.Length Then
                        str_numeroexistente = Regex.Match(ficheroencontrado, ptr_numeroN).Value
                        If str_numeroexistente <> "" Then
                            resultado.Add(directorio & System.IO.Path.DirectorySeparatorChar & nombrefichero.Replace(etiquetaidunico, str_numeroexistente))
                        End If
                    End If
                Next
            End If
        Else
            If exttestigo = VACIO Then
                testigoencontrado = True
            Else
                'buscamos el testigo
                testigoencontrado = System.IO.File.Exists(directorio & System.IO.Path.DirectorySeparatorChar & System.IO.Path.GetFileNameWithoutExtension(nombrefichero) & "." & exttestigo)
            End If
            If testigoencontrado Then
                If System.IO.File.Exists(ruta) Then
                    resultado.Add(ruta)
                End If
            End If
        End If


        Return resultado
    End Function

    Public Sub MueveaDirectorioDestino(ByVal ficheroorigen As String, ByVal ficherodestino As String, ByVal creartestigo As Boolean, ByVal extensionFicheroTestigo As String)

        Dim ptr_numeroN As String = String.empty
        Dim patronbusqueda As String

        Dim etiquetaidunico As String
        Dim longitudid As Integer
        Dim continuar As Boolean
        Dim contador As Long
        Dim str_numeroexistente As String
        Dim numeroexistente As Long
        Dim i As Integer
        Dim directoriodestino, nombreficherodestino, rutacandidata As String
        Dim fichero As System.IO.StreamWriter

        directoriodestino = System.IO.Path.GetDirectoryName(ficherodestino)
        nombreficherodestino = System.IO.Path.GetFileName(ficherodestino)

        etiquetaidunico = Regex.Match(nombreficherodestino, PATRON_IDUNICO).Value
        If etiquetaidunico <> "" Then
            longitudid = CInt(Regex.Match(etiquetaidunico, ptr_NUMERO2).Value)
            For i = 1 To longitudid
                ptr_numeroN &= ptr_NUMERO
            Next

            contador = 1
            patronbusqueda = nombreficherodestino.Replace(etiquetaidunico, New String("?"c, longitudid))
            For Each ficheroencontrado As String In My.Computer.FileSystem.GetFiles(directoriodestino, FileIO.SearchOption.SearchTopLevelOnly, patronbusqueda)
                If System.IO.Path.GetFileName(ficheroencontrado).Length = patronbusqueda.Length Then
                    str_numeroexistente = Regex.Match(ficheroencontrado, ptr_numeroN).Value
                    If str_numeroexistente <> "" Then
                        numeroexistente = CLng(str_numeroexistente)
                        If numeroexistente >= contador Then
                            contador = numeroexistente + 1
                        End If
                    End If
                End If
            Next

            continuar = True
            While continuar
                rutacandidata = directoriodestino & System.IO.Path.DirectorySeparatorChar & nombreficherodestino.Replace(etiquetaidunico, Format(contador, New String("0"c, longitudid)))
                If System.IO.File.Exists(rutacandidata) Then
                    contador += 1
                Else
                    Try
                        System.IO.File.Move(ficheroorigen, rutacandidata)
                        ficherodestino = rutacandidata
                        continuar = False
                    Catch ex As Exception
                        'alguien se nos ha adelantado ¿? 
                        contador += 1
                    End Try
                End If
            End While
        Else
            System.IO.File.Move(ficheroorigen, ficherodestino)
        End If

        If creartestigo Then
            Try
                fichero = New System.IO.StreamWriter(directoriodestino & System.IO.Path.DirectorySeparatorChar & NombreArchivoTestigo(ficherodestino, extensionFicheroTestigo), False, System.Text.Encoding.ASCII)
                fichero.Close()
            Catch ex As Exception
                Throw New Exception("No ha podido grabarse testigo: " & directoriodestino & System.IO.Path.DirectorySeparatorChar & NombreArchivoTestigo(ficherodestino, extensionFicheroTestigo))
            End Try
        End If


    End Sub


    Public Class Token_Nodo
        Public TOKEN As String
        Public contenido As String 'SIN EL TOKEN
        'Esta es la posicion en la que se encuentra el token
        Public posicion As Integer
    End Class

    Public Class ColecciondeTokens
        Inherits coleeccion_object


        <System.Diagnostics.DebuggerStepThrough()> Public Overloads Sub Add(ByVal nombre As String, ByVal valor As Token_Nodo)
            MyBase.Add(nombre, valor)
        End Sub

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function Item(ByVal index As Integer) As Token_Nodo
            Return DirectCast(MyBase.Item(index).valor, Token_Nodo)
        End Function

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function ITem(ByVal nombre As String) As Token_Nodo
            Return DirectCast(MyBase.Item(nombre).valor, Token_Nodo)
        End Function
    End Class
    Public Class ColdeColdeTokens
        Inherits coleeccion_object


        <System.Diagnostics.DebuggerStepThrough()> Public Overloads Sub Add(ByVal valor As ColecciondeTokens)
            MyBase.Add((MyBase.Count + 1).ToString, valor)
        End Sub

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function Item(ByVal index As Integer) As ColecciondeTokens
            Return DirectCast(MyBase.Item(index).valor, ColecciondeTokens)
        End Function

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function ITem(ByVal nombre As String) As ColecciondeTokens
            Return DirectCast(MyBase.Item(nombre).valor, ColecciondeTokens)
        End Function
    End Class
End Module




'Se encarga de reconocer los parametros de de las funciones 
Public MustInherit Class DefinicionFuncionTexto


    Protected Enum TipoParametroDefinicionFuncion
        ParametroEntero
        ParametroFlotante
        ParametroCadena
        ParametroCaracter
        ParametroFecha
        ParametroBooleano
    End Enum
    Protected TipoParametro() As TipoParametroDefinicionFuncion
    Protected DescripcionParametro() As String


    Protected Function ExtraeParametros(ByVal param As String, ByVal constantes As Coleccion_Constantes) As parametros_iniciales
        Dim s As String
        Dim col As ColecciondeTokens
        Dim i, il As Integer
        Dim pi As New parametros_iniciales

        Dim tipo As TipoParametroDefinicionFuncion
        Dim desc As String

        'Valores por defecto
        il = Me.TipoParametro.Length - 1
        For i = 0 To il
            Select Case Me.TipoParametro(i)
                Case TipoParametroDefinicionFuncion.ParametroBooleano
                    pi.Add_Bool(Me.DescripcionParametro(i), False)
                Case TipoParametroDefinicionFuncion.ParametroCadena
                    pi.Add_String(Me.DescripcionParametro(i), "")
                Case TipoParametroDefinicionFuncion.ParametroCaracter
                    pi.Add_Char(Me.DescripcionParametro(i), " "c)
                Case TipoParametroDefinicionFuncion.ParametroEntero
                    pi.Add_Integer(Me.DescripcionParametro(i), 0)
                Case TipoParametroDefinicionFuncion.ParametroFlotante
                    pi.Add_Single(Me.DescripcionParametro(i), 0.0)

            End Select
        Next

        col = LeeParametros(param)
        If col.Count > Me.TipoParametro.Length Then
            Throw New Exception("Número de parámetros incorrecto")
        End If

        il = col.Count - 1
        For i = 0 To il
            s = col.Item(i).contenido
            tipo = Me.TipoParametro(i)
            desc = Me.DescripcionParametro(i)

            Try
                Select Case tipo
                    Case TipoParametroDefinicionFuncion.ParametroCadena
                        pi.Add_String(desc, Parser_TPR.ExtraeParametroCadena(s, constantes))
                    Case TipoParametroDefinicionFuncion.ParametroCaracter
                        pi.Add_String(desc, Parser_TPR.ExtraeParametroCaracter(s))
                    Case TipoParametroDefinicionFuncion.ParametroEntero
                        pi.Add_Long(desc, ExtraeParametroEntero(s))
                    Case TipoParametroDefinicionFuncion.ParametroBooleano
                        pi.Add_Bool(desc, ExtraeParametrobooleano(s))
                    Case TipoParametroDefinicionFuncion.ParametroFecha
                        pi.Add_Date(desc, ExtraeParametroFecha(s))
                    Case TipoParametroDefinicionFuncion.ParametroFlotante
                        pi.Add_Single(desc, ExtraeParametroFlotante(s))
                End Select
            Catch ex As Exception
                Throw New Exception(desc & ":" & ex.Message)
            End Try
        Next

        Return pi

    End Function
    Protected Function ValoresInicialesTipoParametro(ByVal ParamArray cad() As TipoParametroDefinicionFuncion) As TipoParametroDefinicionFuncion()
        Return cad
    End Function

    Protected Function ValoresInicialesDescripcionParametro(ByVal ParamArray cad() As String) As String()
        Return cad
    End Function

    Protected Function QuitaParentesisFinal(ByVal contenido As String) As String
        Dim tmps As String
        tmps = contenido.Trim
        With tmps
            If .EndsWith(Util.PARENTESISD) Then
                Return .Substring(0, .Length - 1) 'le quitamos el parentesis final
            Else
                Throw New Exception("No se encuentra final de parentesis." & contenido)
            End If
        End With
    End Function

End Class

Public Module TPR
    Public Function CalculaExpresionesInternas(ByVal d As Date) As ColecciondeString
        Const VAR_FECHA As String = "$FECHA$"
        Const VAR_FECHALEGIBLE As String = "$FECHALEGIBLE$"
        Const var_HORALEGIBLE As String = "$HORALEGIBLE$"
        Const VAR_DIA As String = "$DIA$"
        Const VAR_MES As String = "$MES$"
        Const VAR_AÑO As String = "$AÑO$"
        Const VAR_HORA As String = "$HORA$"
        Const VAR_MINUTO As String = "$MINUTO$"
        Const VAR_SEGUNDO As String = "$SEGUNDO$"
        Const VAR_MILESIMA As String = "$MILESIMA$"
        'Const VAR_IDUNICOGLOBAL As String = "$IDUNICOGLOBAL$"

        Dim expr_internas As New ColecciondeString

        With expr_internas
            .Add(VAR_FECHA, d.ToString("yyyyMMdd"))
            .Add(VAR_HORA, d.ToString("HH"))
            .Add(VAR_MINUTO, d.ToString("mm"))
            .Add(VAR_SEGUNDO, d.ToString("ss"))
            .Add(VAR_MILESIMA, d.ToString("fff"))
            .Add(VAR_DIA, d.ToString("dd"))
            .Add(VAR_MES, d.ToString("MM"))
            .Add(VAR_AÑO, d.ToString("yyyy"))
            .Add(VAR_FECHALEGIBLE, d.ToString("dd/MM/yyyy"))
            .Add(var_HORALEGIBLE, d.ToString("HH:mm:ss"))

        End With


        Return expr_internas

    End Function

End Module

Public MustInherit Class campoMapeable
    Inherits DefinicionFuncionTexto
    Protected Nombre As String

    Public Property Nombrecampo() As String
        Get
            Return Nombre
        End Get
        Set(ByVal value As String)
            If value = "" Then
                Throw New Exception("El nombre del campo debe ser no vacio")
            Else
                Nombre = value
            End If
        End Set
    End Property

End Class

Public MustInherit Class campotxt
    Inherits campoMapeable

    Protected Const NOMBRE_CAMPO As String = "Nombre Campo"
    Protected Const NOMBRE_LONGITUD As String = "Longitud"

    Private longitud As Integer
    Protected bufferlectura() As Char
    Protected Enum tiporellenocampotxt
        ninguno
        izquierda
        derecha
    End Enum

    Protected relleno As tiporellenocampotxt
    Protected caracterrelleno As Char



    Public Property longitudcampo() As Integer
        Get
            Return longitud
        End Get
        Set(ByVal value As Integer)
            If value > 0 Then
                ReDim bufferlectura(value)
                longitud = value
            Else
                Throw New Exception("Longitud debe ser mayor que cero")
            End If
        End Set
    End Property



    Public MustOverride Sub Escribir(ByVal f As System.IO.StreamWriter, ByVal o As Object)
    'Devuelve el la cadena de lo que ha leido, formateado según SQL
    'P.e: Si leo una cadena de caracteres devuelvo "'loquesea'"
    'Si es un número "3.12"
    'Si es una fecha "Convert(date,...")
    Public MustOverride Function Leer(ByVal f As System.IO.StreamReader) As String


    Protected Function LeerCaracteres(ByVal f As System.IO.StreamReader) As String
        Dim s As String = String.empty
        Dim i As Integer
        If f.ReadBlock(bufferlectura, 0, longitud) <> longitud Then
            Throw New Exception("Se ha encontrado el final del fichero inesperadamente")
        End If

        'Return CStr(bufferlectura)
        For i = 0 To longitud - 1
            s &= CStr(bufferlectura(i))
        Next
        Return s
    End Function

    Protected Function Rellenadato(ByVal s As String) As String
        Dim l As Integer
        l = longitud - s.Length

        If l < 0 Then
            Throw New Exception("El valor a escribir: " & s & " es mayor que la longitud del campo" & longitud.ToString)
        Else
            If l > 0 Then
                Select Case Me.relleno
                    Case tiporellenocampotxt.derecha
                        s = s & New String(Me.caracterrelleno, l)
                    Case tiporellenocampotxt.izquierda
                        s = New String(Me.caracterrelleno, l) & s
                    Case tiporellenocampotxt.ninguno
                End Select
            End If
        End If
        Return s
    End Function

End Class


Public Class Coleccion_camposmapeables
    Inherits coleeccion_object
    Public Overloads Sub Add(ByVal valor As campoMapeable)
        MyBase.Add(valor.Nombrecampo, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As campoMapeable
        Return DirectCast(MyBase.Item(index).valor, campoMapeable)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As campoMapeable
        Return DirectCast(MyBase.Item(nombre).valor, campoMapeable)
    End Function
End Class
Public Class Coleccion_Campostxt
    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal valor As campotxt)
        MyBase.Add(valor.Nombrecampo, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As campotxt
        Return DirectCast(MyBase.Item(index).valor, campotxt)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As campotxt
        Return DirectCast(MyBase.Item(nombre).valor, campotxt)
    End Function
End Class



Public Class campotxtEntero
    Inherits campotxt


    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Me.TipoParametro = Me.ValoresInicialesTipoParametro(TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroEntero)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(NOMBRE_CAMPO, NOMBRE_LONGITUD)

        relleno = tiporellenocampotxt.izquierda
        caracterrelleno = "0"c

        Try
            With ExtraeParametros(param, Constantes)
                Me.longitudcampo = .ITem_Integer(NOMBRE_LONGITUD)
                Me.Nombrecampo = .ITem_String(NOMBRE_CAMPO)
            End With

        Catch ex As Exception
            Throw New Exception("En la definición de Campo Entero: " & ex.Message)
        End Try


    End Sub
    Public Overrides Sub Escribir(ByVal f As System.IO.StreamWriter, ByVal o As Object)
        Dim dato As Long
        Dim s As String

        Try

            If o Is System.DBNull.Value Then
                s = Me.Rellenadato("")
            Else
                Try
                    dato = CLng(o)
                Catch ex As Exception
                    Throw New Exception("Valor no numérico")
                End Try
                s = Rellenadato(dato.ToString)
            End If
            f.Write(s)
        Catch ex As Exception
            Throw New Exception("Escribir Campo Entero." & ex.Message)
        End Try
    End Sub
    Public Overrides Function Leer(ByVal f As System.IO.StreamReader) As String
        Try
            Return CLng(Me.LeerCaracteres(f)).ToString
        Catch ex As Exception
            Throw New Exception("Leer Campo Entero." & ex.Message)
        End Try

    End Function
End Class

Public Class campotxtFlotante
    Inherits campotxt

    Protected ParteEntera As Integer
    Protected ParteDecimal As Integer
    Protected Sepdecimal As String 'Puede tener 0 ó 1 caracter
    Protected longsepdecimal As Integer
    Protected Property Separadordecimal() As String
        Get
            Return Sepdecimal
        End Get
        Set(ByVal value As String)
            If value.Length > 1 Then
                Throw New Exception("El separador decimal debe tener longitud 1 ó 0")
            End If
            Sepdecimal = value
            longsepdecimal = Sepdecimal.Length
        End Set
    End Property

    Protected ReadOnly Property LongitudSeparadordecimal() As Integer
        Get
            Return longsepdecimal
        End Get
    End Property
    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Dim tmps As String
        Const PARAMETRO_PARTEENTERA As String = "Parte Entera"
        Const PARAMETRO_PARTEDECIMAL As String = "Parte Decimal"
        Const PARAMETRO_SEPARADORDECIMAL As String = "Separador Decimal"

        Me.TipoParametro = ValoresInicialesTipoParametro(TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroEntero, TipoParametroDefinicionFuncion.ParametroEntero, TipoParametroDefinicionFuncion.ParametroCaracter)
        Me.DescripcionParametro = ValoresInicialesDescripcionParametro(NOMBRE_CAMPO, PARAMETRO_PARTEENTERA, PARAMETRO_PARTEDECIMAL, PARAMETRO_SEPARADORDECIMAL)

        Me.relleno = tiporellenocampotxt.izquierda
        Me.caracterrelleno = "0"c

        Try
            With Me.ExtraeParametros(param, constantes)

                Me.ParteEntera = .ITem_Integer(PARAMETRO_PARTEENTERA)
                Me.ParteDecimal = .ITem_Integer(PARAMETRO_PARTEDECIMAL)
                Me.Nombrecampo = .ITem_String(NOMBRE_CAMPO)

                tmps = .ITem_String(PARAMETRO_SEPARADORDECIMAL)

            End With
            Select Case tmps.Length
                Case 0
                    Me.Separadordecimal = ""
                Case 1
                    Me.Separadordecimal = tmps.Chars(0)
                Case Else
                    Throw New Exception("El separador decimal debe tener exactamente un caracter")
            End Select

            Me.longitudcampo = Me.ParteEntera + Me.ParteDecimal + Separadordecimal.Length

        Catch ex As Exception
            Throw New Exception("En la definicion de Campo Flotante:" & ex.Message)
        End Try
    End Sub
    Public Overrides Sub Escribir(ByVal f As System.IO.StreamWriter, ByVal o As Object)
        Dim dato As Single
        Dim s As String

        Try
            If o Is System.DBNull.Value Then
                'TODO ¿los nulos son espacios en blanco o todo 0 (con o sin formato)?
                s = Rellenadato("")
            Else
                Try
                    dato = CSng(o)
                Catch ex As Exception
                    Throw New Exception("Valor no numérico")
                End Try
                s = Rellenadato(dato.ToString("#." & New String("0"c, Me.ParteDecimal)).Replace(",", Me.Separadordecimal))
            End If
            f.Write(s)
        Catch ex As Exception
            Throw New Exception("Escribir Campo Flotante." & ex.Message)
        End Try
    End Sub

    Public Overrides Function Leer(ByVal f As System.IO.StreamReader) As String
        Dim s As String
        Dim sent, sdec, ssep As String
        Try
            s = Me.LeerCaracteres(f)

            Try
                sent = CLng(s.Substring(0, Me.ParteEntera)).ToString
            Catch ex As Exception
                Throw New Exception("Parte entera no numérica: " & s)
            End Try


            If Me.LongitudSeparadordecimal > 0 Then
                ssep = s.Substring(Me.ParteEntera, LongitudSeparadordecimal)
                If ssep <> Me.Separadordecimal Then
                    Throw New Exception("Se esperaba separador decimal '" & Me.Separadordecimal & "' se encuentra: " & ssep & " en: " & s)
                End If
            End If

            Try
                sdec = CLng(s.Substring(Me.ParteEntera + LongitudSeparadordecimal, Me.ParteDecimal)).ToString
            Catch ex As Exception
                Throw New Exception("Parte decimal no numérica: " & s)
            End Try

            Return sent & "." & sdec
        Catch ex As Exception
            Throw New Exception("Leer Campo Flotante " & ex.Message)
        End Try
    End Function

End Class
Public Class campotxtCadena
    Inherits campotxt

    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Me.TipoParametro = Me.ValoresInicialesTipoParametro(TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroEntero)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(NOMBRE_CAMPO, NOMBRE_LONGITUD)

        Me.relleno = tiporellenocampotxt.derecha
        Me.caracterrelleno = " "c
        Try

            With Me.ExtraeParametros(param, constantes)
                Me.Nombrecampo = .ITem_String(NOMBRE_CAMPO)
                Me.longitudcampo = .ITem_Integer(NOMBRE_LONGITUD)
            End With

        Catch ex As Exception
            Throw New Exception("En la definicion de Campo Cadena:" & ex.Message)
        End Try
    End Sub
    Public Overrides Sub Escribir(ByVal f As System.IO.StreamWriter, ByVal o As Object)
        Dim dato As String
        Dim s As String
        Try

            If o Is System.DBNull.Value Then
                s = Rellenadato("")
            Else
                Try
                    dato = CStr(o)
                Catch ex As Exception
                    Throw New Exception("Valor no es cadena") '¿entrará alguna vez?
                End Try
                s = Rellenadato(dato)
            End If
            f.Write(s)
        Catch ex As Exception
            Throw New Exception("Escribir Campo Cadena: " & ex.Message)
        End Try
    End Sub
    Public Overrides Function Leer(ByVal f As System.IO.StreamReader) As String
        Return v(Me.LeerCaracteres(f))
    End Function
End Class

Public Class campotxtFecha
    Inherits campotxt

    Dim cadenaformato As String

    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Const NOMBRE_FORMATO_FECHA As String = "Formato Fecha/hora"
        Me.TipoParametro = Me.ValoresInicialesTipoParametro(TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroCadena)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(NOMBRE_CAMPO, NOMBRE_FORMATO_FECHA)

        Me.relleno = tiporellenocampotxt.derecha
        Me.caracterrelleno = " "c
        Try

            With Me.ExtraeParametros(param, constantes)
                Me.Nombrecampo = .ITem_String(NOMBRE_CAMPO)
                cadenaformato = .ITem_String(NOMBRE_FORMATO_FECHA)
                Me.longitudcampo = cadenaformato.Length
            End With

        Catch ex As Exception
            Throw New Exception("En la definicion de Campo Fecha/hora:" & ex.Message)
        End Try
    End Sub

    Public Overrides Sub Escribir(ByVal f As System.IO.StreamWriter, ByVal o As Object)
        Dim dato As Date
        Dim s As String
        Try

            If o Is System.DBNull.Value Then
                s = Rellenadato("")
            Else
                Try
                    dato = CDate(o)
                Catch ex As Exception
                    Throw New Exception("Valor no es Fecha/hora")
                End Try
                s = dato.ToString(Me.cadenaformato)
            End If
            f.Write(s)
        Catch ex As Exception
            Throw New Exception("Escribir Campo Fecha/hora: " & ex.Message)
        End Try
    End Sub

    Public Overrides Function Leer(ByVal f As System.IO.StreamReader) As String
        'TODO definir la manera de leer fechas /horas
        'Simbolos aceptados :
        ' YYYY año (cuatro digitos) ¿alguna vez con dos?
        ' MM mes (dos digitos)
        ' dd dias 
        ' HH (horas)
        ' mm (minutos)
        ' ss segundos

        Return v(LeerCaracteres(f))
    End Function
End Class

Public Class Constante
    Inherits DefinicionFuncionTexto

    Public nombre As String
    Public valor As String

    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Const NOMBRECAMPO1 As String = "Nombre Constante"
        Const NOMBRECAMPO2 As String = "Valor Constante"

        Me.TipoParametro = Me.ValoresInicialesTipoParametro(TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroCadena)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(NOMBRECAMPO1, NOMBRECAMPO2)


        Try
            With Me.ExtraeParametros(param, constantes)
                nombre = .ITem_String(NOMBRECAMPO1)
                valor = .ITem_String(NOMBRECAMPO2)
            End With

            If nombre = "" Then
                Throw New Exception("El nombre de la constante es obligatorio " & param)
            End If
        Catch ex As Exception
            Throw New Exception("En la definición de CTE: " & ex.Message)
        End Try
    End Sub
End Class
Public Class mapeocampos
    Inherits DefinicionFuncionTexto

    Public campo1 As String
    Public campo2 As String


    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Const NOMBRECAMPO1 As String = "Nombre Primer Campo"
        Const NOMBRECAMPO2 As String = "Nombre Segundo Campo"

        Me.TipoParametro = Me.ValoresInicialesTipoParametro(TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroCadena)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(NOMBRECAMPO1, NOMBRECAMPO2)


        Try
            With Me.ExtraeParametros(param, constantes)
                campo1 = .ITem_String(NOMBRECAMPO1)
                campo2 = .ITem_String(NOMBRECAMPO2)
            End With

            If campo1 = "" Or campo2 = "" Then
                Throw New Exception("Deben especificarse dos campos en el mapeo: " & param)
            End If
        Catch ex As Exception
            Throw New Exception("En la definición de Mapeo: " & ex.Message)
        End Try
    End Sub
End Class

Public Class AutoMapeo
    Inherits DefinicionFuncionTexto

    Public auto As Boolean


    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Const AUTOMAPEO As String = "Automapeo"


        Me.TipoParametro = Me.ValoresInicialesTipoParametro(TipoParametroDefinicionFuncion.ParametroBooleano)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(AUTOMAPEO)


        Try
            With Me.ExtraeParametros(param, constantes)
                auto = .ITem_Bool(AUTOMAPEO)
            End With


        Catch ex As Exception
            Throw New Exception("En la definición de AutoMapeo: " & ex.Message)
        End Try
    End Sub
End Class

Public Class Coleccion_Constantes
    Inherits DefinicionFuncionTexto

    Const ETIQUETA_CTE As String = "CTE" & Util.PARENTESISI
    Public SEPARADOR As String = "$"

    Private col_constantes As ColecciondeString
    Public Sub New()
        MyBase.new()
        col_constantes = New ColecciondeString
    End Sub

    Public Sub New(ByVal s As String)
        Me.New()
        Me.Add(s)
    End Sub

    Public Sub Add(ByVal s As String)
        Dim tokens() As String = {ETIQUETA_CTE}
        Dim col As ColecciondeTokens
        Dim tn As Token_Nodo
        Dim i, il As Integer
        Dim contenido As String
        Dim m As Constante


        ' Me.col_constantes = New ColecciondeString

        col = Parser_TPR.SplitPrefijoMultiple(tokens, s)

        il = col.Count - 1

        For i = 0 To il
            Try
                tn = col.Item(i)
                contenido = Me.QuitaParentesisFinal(tn.contenido)

                Select Case tn.TOKEN
                    Case ETIQUETA_CTE
                        'Las nuevas constantes pueden hacer referencia a las anteriores
                        m = New Constante(contenido, Me)
                        Me.col_constantes.Add(normalizaNombre(m.nombre), m.valor)
                End Select
            Catch ex As Exception
                Throw New Exception("Definición Constantes: " & ex.Message)
            End Try
        Next
    End Sub
    Public Sub Combinar(ByVal c As ColecciondeString)
        Dim i, il As Integer
        If Not c Is Nothing Then
            il = c.Count - 1
            For i = 0 To il
                Me.col_constantes.Add(normalizaNombre(c.Indice(i)), c.Item(i))
            Next
        End If
    End Sub
    Private Function normalizaNombre(ByVal s As String) As String
        Dim temp As String
        temp = s.Trim
        If Not temp.StartsWith(SEPARADOR) Then
            temp = SEPARADOR & temp
        End If

        If Not temp.EndsWith(SEPARADOR) Then
            temp = temp & SEPARADOR
        End If
        Return temp
    End Function

    Public Function Reemplaza(ByRef cadena As String) As String
        Dim cambiosrealizados As Boolean = False
        Dim i, il As Integer

        With col_constantes
            il = .Count - 1

            For i = 0 To il
                If cadena.Contains(.Indice(i)) Then
                    cadena = cadena.Replace(.Indice(i), .Item(i))
                End If
            Next
        End With
        Return cadena

    End Function

End Class

Public Class Coleccion_mapeocampos
    Inherits DefinicionFuncionTexto

    Const ETIQUETA_M As String = "M" & Util.PARENTESISI
    Const ETIQUETA_AUTO As String = "AUTO" & Util.PARENTESISI

    Protected col_mapeos As ColecciondeString
    Protected automapeo As Boolean

    Public ReadOnly Property EsAutoMapeo() As Boolean
        Get
            Return automapeo
        End Get
    End Property

    Public ReadOnly Property Existe(ByVal s As String) As Boolean
        Get
            Return col_mapeos.Existe(s.ToUpper)
        End Get
    End Property
    Public ReadOnly Property Item(ByVal s As String) As String
        Get
            Return col_mapeos.Item(s.ToUpper)
        End Get
    End Property
    Public Sub New(ByVal s As String, ByVal constantes As Coleccion_Constantes)
        MyBase.new()
        Dim tokens() As String = {ETIQUETA_M, ETIQUETA_AUTO}
        Dim col As ColecciondeTokens
        Dim tn As Token_Nodo
        Dim i, il As Integer
        Dim contenido As String
        Dim m As mapeocampos
        Dim a As AutoMapeo

        Me.col_mapeos = New ColecciondeString

        col = Parser_TPR.SplitPrefijoMultiple(tokens, s)

        il = col.Count - 1

        For i = 0 To il
            Try
                tn = col.Item(i)
                contenido = Me.QuitaParentesisFinal(tn.contenido)

                Select Case tn.TOKEN
                    Case ETIQUETA_M
                        m = New mapeocampos(contenido, constantes)
                        Me.col_mapeos.Add(m.campo1.ToUpper, m.campo2.ToUpper)
                    Case ETIQUETA_AUTO
                        a = New AutoMapeo(contenido, constantes)
                        automapeo = a.auto
                End Select
            Catch ex As Exception
                Throw New Exception("Definición Mapeo Campos: " & ex.Message)
            End Try
        Next
    End Sub


    '    Protected Function ComprobarMapeo(ByVal c As campoMapeable, ByVal dt As DataTable, ByRef msj As String) As Boolean
    Protected Function ComprobarMapeo(ByVal nombrecampo As String, ByVal dt As DataTable, ByRef msj As String) As Boolean

        If Me.automapeo Then
            If Not dt.Columns.Contains(nombrecampo) Then
                msj = "Se ha especificado Auto Mapeo, y no se encuentra el campo: " & nombrecampo
                Return False
            End If
        Else
            If Not Me.Existe(nombrecampo) Then
                msj = "No se encuentra Mapeo para el campo:" & nombrecampo
                Return False
            End If
        End If

        Return True
    End Function


    Public Function ComprobarMapeo(ByVal campos As Coleccion_Campostxt, ByVal dt As DataTable, ByRef msj As String) As Boolean
        Dim i, il As Integer
        il = campos.Count - 1
        For i = 0 To il
            If Not ComprobarMapeo(campos.Item(i).Nombrecampo, dt, msj) Then
                Return False
            End If
        Next
        Return True
    End Function

    Public Function comprobarMapeo(ByVal campos As coleccion_camposexcel, ByVal dt As DataTable, ByRef msj As String) As Boolean
        Dim i As Integer

        For i = 0 To campos.Count - 1
            If Not comprobarMapeo(campos.Item(i).Nombrecampo, dt, msj) Then
                Return False
            End If
        Next
        Return True
    End Function

    Public Function comprobarMapeo(ByVal tuplas_cte As Coleccion_Tupla_Elemento_Cte, ByVal dt As DataTable, ByRef msj As String) As Boolean
        Dim i As Integer

        For i = 0 To tuplas_cte.Count - 1
            Return comprobarMapeo(tuplas_cte.Item(i).nombreelementotupla, dt, msj)
        Next i
        Return True
    End Function

    Public Function comprobarMapeo(ByVal tuplas_tit As Coleccion_Tupla_Elemento_Titulo, ByVal dt As DataTable, ByRef msj As String) As Boolean
        Dim i As Integer

        For i = 0 To tuplas_tit.Count - 1
            Return comprobarMapeo(tuplas_tit.Item(i).nombreelementotupla, dt, msj)
        Next i
        Return True
    End Function
    Public Function comprobarMapeo(ByVal tuplas_campo As coleccion_Tupla_Elemento_Campo, ByVal dt As DataTable, ByRef msj As String) As Boolean
        Dim i As Integer

        For i = 0 To tuplas_campo.Count - 1
            Return comprobarMapeo(tuplas_campo.Item(i).nombreelementotupla, dt, msj)
        Next i
        Return True
    End Function

    Public Function Mapea(ByVal nombrecampofichero As String) As String
        If automapeo Then
            Return nombrecampofichero
        Else
            Return Me.Item(nombrecampofichero)
        End If
    End Function

End Class
Public Class ColdeColeccion_mapeocampos
    Inherits coleeccion_object

    <System.Diagnostics.DebuggerStepThrough()> Public Overloads Sub Add(ByVal nombre As String, ByVal valor As Coleccion_mapeocampos)
        MyBase.Add(nombre, valor)
    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function Item(ByVal index As Integer) As Coleccion_mapeocampos
        Return DirectCast(MyBase.Item(index).valor, Coleccion_mapeocampos)
    End Function

    <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function ITem(ByVal nombre As String) As Coleccion_mapeocampos
        Return DirectCast(MyBase.Item(nombre).valor, Coleccion_mapeocampos)
    End Function
End Class


Public Class HojaExcel
    Public nombre As String
    Public titulos As coleccion_titulos_excel
    Public campos As coleccion_camposexcel

    ' Public ocultarduplicados As Boolean
    ' Public camporeferencia As String

    Public hoja As Microsoft.Office.Interop.Excel.Worksheet

    Public Sub New(ByVal nombrehoja As String)
        If nombrehoja = "" Then
            Throw New Exception("Definición de Hoja: No es posible un nombre de hoja vacío")
        Else
            nombre = nombrehoja
            titulos = New coleccion_titulos_excel
            campos = New coleccion_camposexcel
        End If
    End Sub

    Public Sub AjustarAnchoColumnas()
        'Conforme van escribiendo en las celdas, se va menteniendo una 
        'coleccion compartida con el ancho de las celdas.
        Dim i, il As Integer
        Dim t As campo_titulo_excel
        Dim celda1 As Object

        With Me.titulos
            il = .Count - 1
            For i = 0 To il
                t = Me.titulos.Item(i)
                If t.Numero_Celdas_Agrupadas = 1 Then
                    celda1 = hoja.Cells.Item(1, t.columna)
                    hoja.Range(celda1, celda1).EntireColumn.AutoFit()
                End If
            Next
        End With


    End Sub
End Class

Public Class Coleccion_HojasExcel
    Inherits coleeccion_object

    Public Overloads Sub Add(ByVal valor As HojaExcel)
        MyBase.Add(valor.nombre, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As HojaExcel
        Return DirectCast(MyBase.Item(index).valor, HojaExcel)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As HojaExcel
        Return DirectCast(MyBase.Item(nombre).valor, HojaExcel)
    End Function
End Class
Public MustInherit Class celda_excel
    Inherits campoMapeable
    Public cadenaformato As String
    Const JUST_DERECHA As String = "<JD>"
    Const JUST_IZQUIERDA As String = "<JI>"
    Const JUST_CENTRO As String = "<JC>"
    Const NEGRITA As String = "<N>"
    Const RESALTAR_FONDO As String = "<RF>"

    Protected tokens() As String = {JUST_DERECHA, JUST_IZQUIERDA, JUST_CENTRO, NEGRITA, RESALTAR_FONDO}


    Protected formato_justificacion As Microsoft.Office.Interop.Excel.XlHAlign
    Protected formato_negrita, formato_resaltarfondo As Boolean

    Protected multiplicador_tam_texto As Double ' para calcular el tamaño del texto (puntos) según el formato

    Public nombrehoja As String
    Public columna As String
    Protected nceldasagrupadas As Integer = 1
    Protected formatodenumero As String = "@" 'TODO lo vamos a crear como Texto, para que respete nuestro formato. En el futuro podremos afinar asignando a cada clase derivada un formato distinto (enteros,flotantes....)


    Public Property formato() As String
        Get
            Return cadenaformato
        End Get
        Set(ByVal value As String)
            cadenaformato = value
            formato_justificacion = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            formato_negrita = False
            formato_resaltarfondo = False

            multiplicador_tam_texto = 1.0

            For Each t As String In tokens
                If cadenaformato.Contains(t) Then
                    Select Case t
                        Case JUST_DERECHA
                            formato_justificacion = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                        Case JUST_IZQUIERDA
                            formato_justificacion = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                        Case JUST_CENTRO
                            formato_justificacion = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        Case NEGRITA
                            formato_negrita = True
                            multiplicador_tam_texto = 1.25
                        Case RESALTAR_FONDO
                            formato_resaltarfondo = True
                    End Select
                End If
            Next


        End Set
    End Property
    Public Property Numero_Celdas_Agrupadas() As Integer
        Get
            Return nceldasagrupadas
        End Get
        Set(ByVal value As Integer)
            If value < 1 Then
                Throw New Exception("Celdas Agrupadas: No es posible un valor menor que 1. " & value.ToString)
            End If
            nceldasagrupadas = value
        End Set

    End Property

    Public Property Fila() As Integer
        Get
            Return numerofila
        End Get
        Set(ByVal value As Integer)
            numerofila = value
            numerofilaactual = value
        End Set
    End Property
    Public numerofila As Integer
    Public numerofilaactual As Integer


    Public Sub ResetearFila()
        numerofilaactual = Fila
    End Sub
    Public Sub AddFila(ByVal n As Integer)
        numerofilaactual += n
    End Sub

    Protected Overridable Sub EscribeValor(ByVal s As String, ByVal h As HojaExcel)
        Dim celda1, celda2 As Object
        Dim hoja As Microsoft.Office.Interop.Excel.Worksheet = h.hoja
        Dim rango As Microsoft.Office.Interop.Excel.Range


        celda1 = hoja.Cells.Item(Me.numerofilaactual, Me.columna)

        If Me.nceldasagrupadas = 1 Then
            rango = hoja.Range(celda1, celda1)
        Else
            celda2 = hoja.Cells.Item(Me.numerofilaactual, SumaColumnasExcel(Me.columna, nceldasagrupadas))
            rango = hoja.Range(celda1, celda2)
            rango.Merge()
        End If
        With rango
            .NumberFormat = formatodenumero
            .Value = s
            .HorizontalAlignment = formato_justificacion
            .Characters.Font.Bold = formato_negrita
            If formato_resaltarfondo Then
                .Interior.ColorIndex = 15
                .Interior.Pattern = 1 ' SOLIDO
            End If
        End With

    End Sub





End Class

Public Class Excel_ocultar_duplicados
    Inherits DefinicionFuncionTexto

    Public nombrehoja As String
    Public ocultarduplicados As Boolean
    Public camporeferencia As String
    Public campoaocultar As String

    Protected Const cte_NOMBRE_HOJA As String = "Nombre Hoja"
    Protected Const cte_OCULTAR As String = "Ocultar Duplicados"
    Protected Const cte_CAMPO_REFERENCIA As String = "Campo de Referencia"
    Protected Const cte_CAMPO_A_OCULTAR As String = "Campo a Ocultar"

    Public Sub New(ByVal param As String, ByVal c As Coleccion_Constantes)
        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroBooleano, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_HOJA, cte_OCULTAR, cte_CAMPO_REFERENCIA, cte_CAMPO_A_OCULTAR)

        With Me.ExtraeParametros(param, c)
            nombrehoja = .ITem_String(cte_NOMBRE_HOJA).ToUpper
            ocultarduplicados = .ITem_Bool(cte_OCULTAR)
            camporeferencia = .ITem_String(cte_CAMPO_REFERENCIA)
            campoaocultar = .ITem_String(cte_CAMPO_A_OCULTAR)
        End With
    End Sub
End Class
Public MustInherit Class campo_excel
    Inherits celda_excel

    Protected Const cte_NOMBRE_HOJA As String = "Nombre Hoja"
    Protected Const cte_NOMBRE_CAMPO As String = "Nombre Campo"
    Protected Const cte_DESCRIPCION_CAMPO As String = "Descripcion Campo"
    Protected Const cte_NOMBRE_COLUMNA As String = "Nombre Columna"
    Protected Const cte_NOMBRE_FORMATO As String = "Formato"



    Protected Nombre_Campo As String
    Protected Descripcion_Campo As String

    Protected formato_interno As String

    'Para la logica de ocultación de los campos (que no aparezcan campos repetidos)
    Protected UltimoValorCampo As String
    Protected Ocultado_linea_actual As Boolean
    Public ocultarduplicados As Boolean
    Public camporeferencia_ocultar As String

    Public MustOverride Sub Escribe(ByVal dr As DataRow, ByVal nombrecampodeldr As String, ByVal hoja As HojaExcel)


    Protected Sub EscribeQuitandoRepeticiones(ByVal s As String, ByVal hoja As HojaExcel)
        Dim cad As String


        If Me.ocultarduplicados Then
            If s = UltimoValorCampo Then
                If Me.camporeferencia_ocultar = "" Then
                    Ocultado_linea_actual = True
                ElseIf Me.camporeferencia_ocultar = Me.Nombrecampo Then
                    Ocultado_linea_actual = True
                Else
                    Ocultado_linea_actual = hoja.campos.Item(Me.camporeferencia_ocultar).Ocultado_linea_actual
                End If
            Else
                Ocultado_linea_actual = False
            End If
        Else
            Ocultado_linea_actual = False
        End If

        UltimoValorCampo = s
        If Ocultado_linea_actual Then
            cad = ""
        Else
            cad = s
        End If


        MyBase.EscribeValor(cad, hoja)
    End Sub
    Protected Overridable Sub LeerDefinicion(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Const P1 As String = "<FORMATO"
        Const P2 As String = ">"
        Const expresionregular As String = P1 & ".+" & P2

        Dim col As System.Text.RegularExpressions.MatchCollection
        Me.TipoParametro = Me.ValoresInicialesTipoParametro(TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroCadena, TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_HOJA, cte_NOMBRE_CAMPO, cte_DESCRIPCION_CAMPO, cte_NOMBRE_COLUMNA, cte_NOMBRE_FORMATO)


        With Me.ExtraeParametros(param, constantes)
            Me.nombrehoja = .ITem_String(cte_NOMBRE_HOJA).ToUpper
            Me.Nombrecampo = .ITem_String(cte_NOMBRE_CAMPO)
            Me.Descripcion_Campo = .ITem_String(cte_DESCRIPCION_CAMPO)
            Me.columna = .ITem_String(cte_NOMBRE_COLUMNA)
            Me.formato = .ITem_String(cte_NOMBRE_FORMATO).ToUpper

            'En formato, vamos a tener una cadena de formato con nuestros
            'marcadors <N> negrita, <JC> justificacion centro ...
            'Además para los flotantes, vamos a tener 
            '<FORMATO xxxxxxxx>
            'donde xxxxxxxxxx es el formato que acepta la clausula
            ' .tostring("xxxxxxxxx")

            col = Regex.Matches(Me.formato, expresionregular)
            If col.Count > 0 Then
                With col.Item(0).Value
                    Me.formato_interno = .Substring(P1.Length, .Length - P1.Length - P2.Length)
                End With
            Else
                Me.formato_interno = ""
            End If

            Me.ocultarduplicados = False
        End With
    End Sub


End Class

Public Class campo_excel_entero
    Inherits campo_excel


    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)

        Try
            Me.LeerDefinicion(param, constantes)
        Catch ex As Exception
            Throw New Exception("En la definicion de Campo Cadena:" & ex.Message)
        End Try
    End Sub
    Public Overrides Sub Escribe(ByVal dr As System.Data.DataRow, ByVal nombrecampodeldr As String, ByVal hoja As HojaExcel)
        Dim valor As String
        If dr.Item(nombrecampodeldr) Is System.DBNull.Value Then
            valor = ""
        Else
            valor = CLng(dr.Item(nombrecampodeldr)).ToString(Me.formato_interno)
        End If
        Me.EscribeQuitandoRepeticiones(valor, hoja)
    End Sub

End Class

Public Class campo_excel_cadena
    Inherits campo_excel


    Public Overrides Sub Escribe(ByVal dr As System.Data.DataRow, ByVal nombrecampodeldr As String, ByVal hoja As HojaExcel)
        Dim valor As String
        If dr.Item(nombrecampodeldr) Is System.DBNull.Value Then
            valor = ""
        Else
            valor = CStr(dr.Item(nombrecampodeldr))
        End If
        Me.EscribeQuitandoRepeticiones(valor, hoja)
    End Sub

    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)

        Try
            Me.LeerDefinicion(param, constantes)
        Catch ex As Exception
            Throw New Exception("En la definicion de Campo Cadena:" & ex.Message)
        End Try
    End Sub
End Class

Public Class campo_excel_flotante
    Inherits campo_excel

    Public Overrides Sub Escribe(ByVal dr As System.Data.DataRow, ByVal nombrecampodeldr As String, ByVal hoja As HojaExcel)
        Dim valor As String
        If dr.Item(nombrecampodeldr) Is System.DBNull.Value Then
            valor = ""
        Else
            valor = CSng(dr.Item(nombrecampodeldr)).ToString(Me.formato_interno)
        End If

        Me.EscribeQuitandoRepeticiones(valor, hoja)
    End Sub

    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Try
            Me.LeerDefinicion(param, constantes)
        Catch ex As Exception
            Throw New Exception("En la definicion de Campo Cadena:" & ex.Message)
        End Try
    End Sub


End Class

Public Class campo_excel_fecha
    Inherits campo_excel

    Public Overrides Sub Escribe(ByVal dr As System.Data.DataRow, ByVal nombrecampodeldr As String, ByVal hoja As HojaExcel)
        Dim f As String
        Dim valor As String

        If dr.Item(nombrecampodeldr) Is System.DBNull.Value Then
            valor = ""
        Else
            If formato_interno = "" Then
                f = Util.FORMATO_FECHA_DMA
            Else
                f = formato_interno
            End If
            valor = CDate(dr.Item(nombrecampodeldr)).ToString(f)
        End If
        Me.EscribeQuitandoRepeticiones(valor, hoja)
    End Sub

    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Try
            Me.LeerDefinicion(param, constantes)
        Catch ex As Exception
            Throw New Exception("En la definicion de Campo Fecha:" & ex.Message)
        End Try
    End Sub
End Class
Public Class campo_excel_hora
    Inherits campo_excel

    Public Overrides Sub Escribe(ByVal dr As System.Data.DataRow, ByVal nombrecampodeldr As String, ByVal hoja As HojaExcel)
        Dim f As String
        If formato_interno = "" Then
            f = Util.FORMATO_HORA_HM
        Else
            f = formato_interno
        End If

        Me.EscribeQuitandoRepeticiones(CDate(dr.Item(nombrecampodeldr)).ToString(f), hoja)
    End Sub

    Public Sub New(ByVal param As String, ByVal constantes As Coleccion_Constantes)
        Try
            Me.LeerDefinicion(param, constantes)
        Catch ex As Exception
            Throw New Exception("En la definicion de Campo Hora:" & ex.Message)
        End Try
    End Sub
End Class


Public Class coleccion_camposexcel
    Inherits coleeccion_object
    Public Overloads Sub Add(ByVal valor As campo_excel)
        MyBase.Add(valor.Nombrecampo, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As campo_excel
        Return DirectCast(MyBase.Item(index).valor, campo_excel)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As campo_excel
        Return DirectCast(MyBase.Item(nombre).valor, campo_excel)
    End Function

End Class
Public Class campo_titulo_excel
    Inherits celda_excel

    Public contenido As String

    Public Sub New(ByVal s As String, ByVal c As Coleccion_Constantes)
        'buscamos el nombre y el contenido
        Const cte_NOMBRE_HOJA As String = "Nombre Hoja"
        Const NOMBRE_CAMPO As String = "Nombre"
        Const NOMBRE_CONTENIDO As String = "Contenido"
        Const NOMBRE_FILA As String = "Fila"
        Const NOMBRE_COLUMNA As String = "Columna"
        Const NOMBRE_NUMERO_CELDAS As String = "Numero celdas Agrupadas"
        Const NOMBRE_FORMATO As String = "Formato"

        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroEntero, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroEntero, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_HOJA, NOMBRE_CAMPO, NOMBRE_CONTENIDO, NOMBRE_COLUMNA, NOMBRE_FILA, NOMBRE_NUMERO_CELDAS, NOMBRE_FORMATO)

        With Me.ExtraeParametros(s, c)
            nombrehoja = .ITem_String(cte_NOMBRE_HOJA).ToUpper
            Nombrecampo = .ITem_String(NOMBRE_CAMPO)
            contenido = .ITem_String(NOMBRE_CONTENIDO)
            columna = .ITem_String(NOMBRE_COLUMNA)
            Fila = .ITem_Integer(NOMBRE_FILA)
            Me.Numero_Celdas_Agrupadas = .ITem_Integer(NOMBRE_NUMERO_CELDAS)
            Me.formato = .ITem_String(NOMBRE_FORMATO).ToUpper
        End With
    End Sub

    Public Sub Escribe(ByVal hoja As HojaExcel)
        'Si el contenido tiene varias líneas (separadas por vbcrlf)
        'lo escribimos en varias filas consecutivas
        Dim s, partes() As String

        partes = Split(Me.contenido, vbCrLf)
        For Each s In partes
            Me.EscribeValor(s, hoja)
            Me.numerofilaactual += 1
        Next

    End Sub
End Class

Public Class coleccion_titulos_excel
    Inherits coleeccion_object

    Public Overloads Sub Add(ByVal valor As campo_titulo_excel)
        MyBase.Add(valor.Nombrecampo, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As campo_titulo_excel
        Return DirectCast(MyBase.Item(index).valor, campo_titulo_excel)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As campo_titulo_excel
        Return DirectCast(MyBase.Item(nombre).valor, campo_titulo_excel)
    End Function

End Class



Public Class Coleccion_Tupla_Elemento_Cte
    Inherits coleeccion_object

    Public Overloads Sub Add(ByVal valor As Tupla_Elemento_CTE)
        MyBase.Add(valor.nombreelementotupla, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Tupla_Elemento_CTE
        Return DirectCast(MyBase.Item(index).valor, Tupla_Elemento_CTE)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Tupla_Elemento_CTE
        Return DirectCast(MyBase.Item(nombre).valor, Tupla_Elemento_CTE)
    End Function
End Class
Public MustInherit Class Tupla_Elemento_CTE
    Inherits DefinicionFuncionTexto


    Public nombretupla As String
    Public nombreelementotupla As String

    Protected Const cte_NOMBRE_TUPLA As String = "NOMBRETUPLA"
    Protected cte_NOMBRE_ELEMENTO_TUPLA As String = "NOMBRE ELEMENTO TUPLA"
    Protected cte_vALOR As String = "VALOR"

    Public MustOverride Sub AsignaValor(ByVal dr As DataRow, ByVal nombrecampo As String, ByVal contadorlocal As Integer, ByVal contadorglobal As Integer)
End Class

Public Class Tupla_ELEMENTO_CTE_Entero
    Inherits Tupla_Elemento_CTE

    Dim valor As Integer

    Public Sub New(ByVal param As String, ByVal cte As Coleccion_Constantes)

        'Falta el nombre y el tipo del valor constante.
        'Puede ser entero, cadena, fecha, flotante

        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroEntero)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_TUPLA, cte_NOMBRE_ELEMENTO_TUPLA, cte_vALOR)

        Try
            With Me.ExtraeParametros(param, cte)
                'dependiendo del tipo haremos una cosa u otra
                nombretupla = .ITem_String(cte_NOMBRE_TUPLA)
                nombreelementotupla = .ITem_String(cte_NOMBRE_ELEMENTO_TUPLA)
                valor = .ITem_Integer(cte_vALOR)
            End With
        Catch ex As Exception
            Throw New Exception("Definición Elemento Tupla Constante " & ex.Message)
        End Try

    End Sub
    Public Overrides Sub AsignaValor(ByVal dr As System.Data.DataRow, ByVal nombrecampo As String, ByVal contadorlocal As Integer, ByVal contadorglobal As Integer)
        dr(nombrecampo) = valor
    End Sub

End Class

Public Class Tupla_ELEMENTO_CTE_Cadena
    Inherits Tupla_Elemento_CTE

    Dim valor As String

    Public Sub New(ByVal param As String, ByVal cte As Coleccion_Constantes)

        'Falta el nombre y el tipo del valor constante.
        'Puede ser entero, cadena, fecha, flotante

        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_TUPLA, cte_NOMBRE_ELEMENTO_TUPLA, cte_vALOR)

        Try
            With Me.ExtraeParametros(param, cte)
                'dependiendo del tipo haremos una cosa u otra
                nombretupla = .ITem_String(cte_NOMBRE_TUPLA)
                nombreelementotupla = .ITem_String(cte_NOMBRE_ELEMENTO_TUPLA)
                valor = .ITem_String(cte_vALOR)
            End With
        Catch ex As Exception
            Throw New Exception("Definición Elemento Tupla Constante " & ex.Message)
        End Try

    End Sub
    Public Overrides Sub AsignaValor(ByVal dr As System.Data.DataRow, ByVal nombrecampo As String, ByVal contadorlocal As Integer, ByVal contadorglobal As Integer)
        Const PTR_DOSENTEROS As String = "[0-9][0-9]"
        Const PATRON_IDLOCAL As String = "\$IDLOCAL" & PTR_DOSENTEROS & "\$"
        Const PATRON_IDGLOBAL As String = "\$IDGLOBAL" & PTR_DOSENTEROS & "\$"

        Dim etiquetaidlocal, etiquetaidglobal As String
        Dim valoraescribir As String


        valoraescribir = valor
        Do
            etiquetaidlocal = Regex.Match(valoraescribir, PATRON_IDLOCAL).Value
            If etiquetaidlocal <> "" Then
                valoraescribir = valoraescribir.Replace(etiquetaidlocal, contadorlocal.ToString(New String("0"c, CInt(Regex.Match(etiquetaidlocal, PTR_DOSENTEROS).Value))))
            End If
            etiquetaidglobal = Regex.Match(valoraescribir, PATRON_IDGLOBAL).Value
            If etiquetaidglobal <> "" Then
                valoraescribir = valoraescribir.Replace(etiquetaidglobal, contadorglobal.ToString(New String("0"c, CInt(Regex.Match(etiquetaidglobal, PTR_DOSENTEROS).Value))))
            End If
        Loop While etiquetaidlocal <> "" Or etiquetaidglobal <> ""
        dr(nombrecampo) = valoraescribir
    End Sub
End Class

Public Class Tupla_ELEMENTO_CTE_flotante
    Inherits Tupla_Elemento_CTE

    Dim valor As Single

    Public Sub New(ByVal param As String, ByVal cte As Coleccion_Constantes)

        'Falta el nombre y el tipo del valor constante.
        'Puede ser entero, cadena, fecha, flotante

        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroFlotante)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_TUPLA, cte_NOMBRE_ELEMENTO_TUPLA, cte_vALOR)

        Try
            With Me.ExtraeParametros(param, cte)
                'dependiendo del tipo haremos una cosa u otra
                nombretupla = .ITem_String(cte_NOMBRE_TUPLA)
                nombreelementotupla = .ITem_String(cte_NOMBRE_ELEMENTO_TUPLA)
                valor = .ITem_Single(cte_vALOR)
            End With
        Catch ex As Exception
            Throw New Exception("Definición Elemento Tupla Constante Flotante" & ex.Message)
        End Try

    End Sub
    Public Overrides Sub AsignaValor(ByVal dr As System.Data.DataRow, ByVal nombrecampo As String, ByVal contadorlocal As Integer, ByVal contadorglobal As Integer)
        dr(nombrecampo) = valor
    End Sub
End Class
Public Class Tupla_ELEMENTO_CTE_Fecha
    Inherits Tupla_Elemento_CTE

    Dim valor As Date

    Public Sub New(ByVal param As String, ByVal cte As Coleccion_Constantes)

        'Falta el nombre y el tipo del valor constante.
        'Puede ser entero, cadena, fecha, flotante

        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroFecha)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_TUPLA, cte_NOMBRE_ELEMENTO_TUPLA, cte_vALOR)

        Try
            With Me.ExtraeParametros(param, cte)
                'dependiendo del tipo haremos una cosa u otra
                nombretupla = .ITem_String(cte_NOMBRE_TUPLA)
                nombreelementotupla = .ITem_String(cte_NOMBRE_ELEMENTO_TUPLA)
                valor = .ITem_Date(cte_vALOR)
            End With
        Catch ex As Exception
            Throw New Exception("Definición Elemento Tupla Constante Fecha" & ex.Message)
        End Try

    End Sub
    Public Overrides Sub AsignaValor(ByVal dr As System.Data.DataRow, ByVal nombrecampo As String, ByVal contadorlocal As Integer, ByVal contadorglobal As Integer)
        dr(nombrecampo) = valor
    End Sub
End Class

Public Class Coleccion_Tupla_Elemento_Titulo
    Inherits coleeccion_object

    Public Overloads Sub Add(ByVal valor As Tupla_Elemento_Titulo)
        MyBase.Add(valor.nombreelementotupla, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Tupla_Elemento_Titulo
        Return DirectCast(MyBase.Item(index).valor, Tupla_Elemento_Titulo)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Tupla_Elemento_Titulo
        Return DirectCast(MyBase.Item(nombre).valor, Tupla_Elemento_Titulo)
    End Function
End Class
Public Class Tupla_Elemento_Titulo
    Inherits Tupla_Elemento_CTE


    Protected Const cte_NOMBRE_COLUMNA As String = "Columna"
    Protected Const cte_NOMBRE_FILA As String = "Fila"
    Public fila As Long
    Public columna As String

    Public hoja As Microsoft.Office.Interop.Excel.Worksheet

    Public Sub New()
        'no debería usarse, es para que las clases que heredan de esta
        'puedan llamar a un constructor vacio 
    End Sub
    Public Sub New(ByVal param As String, ByVal ctes As Coleccion_Constantes)
        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroEntero)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_TUPLA, cte_NOMBRE_ELEMENTO_TUPLA, cte_NOMBRE_COLUMNA, cte_NOMBRE_FILA)

        Try
            With Me.ExtraeParametros(param, ctes)
                'dependiendo del tipo haremos una cosa u otra
                nombretupla = .ITem_String(cte_NOMBRE_TUPLA)
                nombreelementotupla = .ITem_String(cte_NOMBRE_ELEMENTO_TUPLA)
                columna = .ITem_String(cte_NOMBRE_COLUMNA)
                fila = .ITem_Integer(cte_NOMBRE_FILA)
            End With
        Catch ex As Exception
            Throw New Exception("Definición Elemento Tupla Titulo " & ex.Message)
        End Try
    End Sub

    Public Overrides Sub AsignaValor(ByVal dr As System.Data.DataRow, ByVal nombrecampo As String, ByVal contadorlocal As Integer, ByVal contadorglobal As Integer)
        Dim celda1 As Object
        Dim rango As Microsoft.Office.Interop.Excel.Range

        celda1 = hoja.Cells.Item(Me.fila, Me.columna)

        rango = hoja.Range(celda1, celda1)
        dr(nombrecampo) = rango.Value
    End Sub
    Public Function Estavacio() As Boolean
        Dim celda1 As Object
        Dim rango As Microsoft.Office.Interop.Excel.Range

        celda1 = hoja.Cells.Item(Me.fila, Me.columna)

        rango = hoja.Range(celda1, celda1)

        Return CStr(rango.Value) = ""
    End Function
End Class

Public Class coleccion_Tupla_Elemento_Campo
    Inherits coleeccion_object

    Public Overloads Sub Add(ByVal valor As Tupla_Elemento_Campo)
        MyBase.Add(valor.nombreelementotupla, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Tupla_Elemento_Campo
        Return DirectCast(MyBase.Item(index).valor, Tupla_Elemento_Campo)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Tupla_Elemento_Campo
        Return DirectCast(MyBase.Item(nombre).valor, Tupla_Elemento_Campo)
    End Function
End Class
Public Class Tupla_Elemento_Campo
    Inherits Tupla_Elemento_Titulo


    Public filainicial As Long
    Public columnainicial As String
    Public incrementofila As Integer
    Public incrementocolumna As Integer

    Protected Const CTE_NOMBRE_INCREMENTO_FILA As String = "Incremento fila"
    Protected Const CTE_NOMBRE_INCREMENTO_COLUMNA As String = "Incremento Columna"

    Public Sub New(ByVal param As String, ByVal ctes As Coleccion_Constantes)
        MyBase.New() 'por eso la clase base tiene un constructor vacio

        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, _
                                                            DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, _
                                                            DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, _
                                                            DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroEntero, _
                                                            DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroEntero, _
                                                            DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroEntero)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_TUPLA, cte_NOMBRE_ELEMENTO_TUPLA, cte_NOMBRE_COLUMNA, _
                                                                        cte_NOMBRE_FILA, CTE_NOMBRE_INCREMENTO_COLUMNA, CTE_NOMBRE_INCREMENTO_FILA)

        Try
            With Me.ExtraeParametros(param, ctes)
                'dependiendo del tipo haremos una cosa u otra
                nombretupla = .ITem_String(cte_NOMBRE_TUPLA)
                nombreelementotupla = .ITem_String(cte_NOMBRE_ELEMENTO_TUPLA)
                columnainicial = .ITem_String(cte_NOMBRE_COLUMNA)
                filainicial = .ITem_Integer(cte_NOMBRE_FILA)
                incrementocolumna = .ITem_Integer(CTE_NOMBRE_INCREMENTO_COLUMNA) + 1 'la funcion de suma columnas necesita uno mas
                incrementofila = .ITem_Integer(CTE_NOMBRE_INCREMENTO_FILA)
            End With
            columna = columnainicial
            fila = filainicial
        Catch ex As Exception
            Throw New Exception("Definición Elemento Tupla Campo " & ex.Message)
        End Try
    End Sub

    Public Overrides Sub AsignaValor(ByVal dr As System.Data.DataRow, ByVal nombrecampo As String, ByVal contadorlocal As Integer, ByVal contadorglobal As Integer)

        MyBase.AsignaValor(dr, nombrecampo, contadorlocal, contadorglobal)
        Me.fila += Me.incrementofila
        Me.columna = Util.SumaColumnasExcel(Me.columna, Me.incrementocolumna)
    End Sub

End Class


Public Class Coleccion_tupla_lecturaExcel
    Inherits coleeccion_object

    Public Overloads Sub Add(ByVal valor As Tupla_lecturaExcel)
        MyBase.Add(valor.nombretupla, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Tupla_lecturaExcel
        Return DirectCast(MyBase.Item(index).valor, Tupla_lecturaExcel)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Tupla_lecturaExcel
        Return DirectCast(MyBase.Item(nombre).valor, Tupla_lecturaExcel)
    End Function
End Class
Public Class Tupla_lecturaExcel
    Inherits DefinicionFuncionTexto

    'Una tupla de lectura está compuesta por 

    Public valores_constantes As New Coleccion_Tupla_Elemento_Cte
    Public valores_titulos As New Coleccion_Tupla_Elemento_Titulo
    Public valores_campos As New coleccion_Tupla_Elemento_Campo

    Public nombretupla As String
    Public nombrehoja As String

    Public Sub New(ByVal param As String, ByVal ctes As Coleccion_Constantes)
        Const cte_NOMBRE_HOJA As String = "NOMBREHOJA"
        Const cte_NOMBRE_TUPLA As String = "NOMBRETUPLA"
        Me.TipoParametro = Me.ValoresInicialesTipoParametro(DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena, DefinicionFuncionTexto.TipoParametroDefinicionFuncion.ParametroCadena)
        Me.DescripcionParametro = Me.ValoresInicialesDescripcionParametro(cte_NOMBRE_HOJA, cte_NOMBRE_TUPLA)

        Try
            With Me.ExtraeParametros(param, ctes)
                nombrehoja = .ITem_String(cte_NOMBRE_HOJA).ToUpper
                nombretupla = .ITem_String(cte_NOMBRE_TUPLA).ToUpper
            End With
        Catch ex As Exception
            Throw New Exception("Definicion Tupla Escritura Excel : " & ex.Message)
        End Try
    End Sub
End Class

Public Class Exception_Proceso_ejecucion_TPR_cancelado_por_usuario
    Inherits Exception
    Public Sub New(ByVal msg As String)
        MyBase.New(msg)
    End Sub
End Class
Public Class Ejecucion_TPR

    Public Enum NIVEL_INFORMACION_EVENTO
        NIVEL1
        NIVEL2
    End Enum

    Public Event Inicio(ByVal nivel As NIVEL_INFORMACION_EVENTO, ByVal maximo As Integer)
    Public Event Avance(ByVal nivel As NIVEL_INFORMACION_EVENTO, ByVal a As Integer)
    Public Event Fin(ByVal nivel As NIVEL_INFORMACION_EVENTO)

    Public Const CUATROPUNTOS As String = "::"
    Public Const ETIQUETA_TABLA As String = "TABLA" & CUATROPUNTOS
    Public Const ETIQUETA_COMANDO As String = "COMANDO" & CUATROPUNTOS
    Public Const ETIQUETA_CONEXION_SQL As String = "CONEXIONSQL" & CUATROPUNTOS
    Public Const ETIQUETA_TRANSFERIR As String = "TRANSFERIR" & CUATROPUNTOS
    Public Const ETIQUETA_INCLUDE As String = "INCLUDE" & CUATROPUNTOS
    Public Const ETIQUETA_CONEXION_TEXTO As String = "CONEXIONTEXTO" & CUATROPUNTOS
    Public Const ETIQUETA_MAPEO_CAMPOS As String = "MAPEO" & CUATROPUNTOS
    Public Const ETIQUETA_CONSTANTE As String = "CONSTANTE" & CUATROPUNTOS
    Public Const ETIQUETA_CONEXIONMYSQL As String = "CONEXIONMYSQL" & CUATROPUNTOS
    Public Const ETIQUETA_CONEXION_EXCEL As String = "CONEXIONEXCEL" & CUATROPUNTOS
    Const NOEXISTE As Integer = -1


    Private tokens() As String = {ETIQUETA_TABLA, ETIQUETA_COMANDO, ETIQUETA_CONEXION_SQL, ETIQUETA_TRANSFERIR, ETIQUETA_INCLUDE, ETIQUETA_CONEXION_TEXTO, ETIQUETA_MAPEO_CAMPOS, ETIQUETA_CONSTANTE, ETIQUETA_CONEXION_EXCEL, ETIQUETA_CONEXIONMYSQL}

    'Public Sub GeneraXMLdeDataset(ByVal conexion As ConexionSQL, ByVal informe As String)
    '    Dim f, consulta As String
    '    Dim t As DataTable


    '    f = Globales.directorio_informes & informe & ".tpr"
    '    consulta = LeeFicheroConsulta(f)
    '    t = conexion.SelectSQL(consulta)
    '    t.WriteXmlSchema(Globales.directorio_informes & informe & ".xml")

    '    conexion.Libera(t)

    'End Sub

    Private enprogreso As Boolean
    Friend cancelar As Boolean
    Public Sub CancelarProgreso()
        If enprogreso Then
            cancelar = True
        End If
    End Sub

    Friend Sub InformaEventoHijo_Inicio(ByVal nivel As NIVEL_INFORMACION_EVENTO, ByVal maximo As Integer)
        RaiseEvent Inicio(nivel, maximo)
    End Sub

    Friend Sub InformaEventoHijo_Avance(ByVal nivel As NIVEL_INFORMACION_EVENTO, ByVal a As Integer)
        RaiseEvent Avance(nivel, a)
    End Sub
    Friend Sub InformaEventoHijo_Fin(ByVal nivel As NIVEL_INFORMACION_EVENTO)
        RaiseEvent Fin(nivel)
    End Sub


    Public Shared Function Serializa_Filtro(ByVal f As ColecciondeString) As String
        Dim r As String = VACIO
        Dim i, fin As Integer

        fin = f.Count - 1

        For i = 0 To fin
            Util.AñadeSinoVacio(r, Y, f.Item(i))
        Next

        Return r


    End Function
    Public Shared Function Serializa_filtro(ByVal filtro As ColecciondeString, ByVal fragmentofrom As String) As String
        Dim fin, i As Integer
        Dim f As String = VACIO
        Dim al As String


        fin = filtro.Count - 1
        For i = 0 To fin
            al = extraeParte1(filtro.Indice(i))
#If Not COMPACT_FRAMEWORK Then
            If al = VACIO OrElse fragmentofrom.ToUpper.Contains(" " & al & " ") Then
#Else
            If al = VACIO OrElse fragmentofrom.ToCharArray Like ("* " & al & "* ") Then
#End If


                Util.AñadeSinoVacio(f, Y, filtro.Item(i))
            End If
        Next

        Return f
    End Function


    Public Shared Function Codifica_filtro(ByVal f As ColecciondeString) As String
        'Necesito algo como 
        'ALIASTABLA.NOMBRECAMPO , FILTRO 
        'EL FILTRO PUEDE SER COMPLEJO 

        ' hex(alias.nombrecampo)-hex(FILTRO);hex(alias.nombrecampo)-hex(FILTRO)....
        Dim i, fin As Integer
        Dim result As String = VACIO

        If f Is Nothing Then
            Return VACIO
        End If

        fin = f.Count - 1

        For i = 0 To fin
            Util.AñadeSinoVacio(result, PUNTOYCOMA, PasaraHex(f.Indice(i)) & "-" & PasaraHex(f.Item(i)))
        Next

        Return result
    End Function

    Public Shared Function extraeParte2(ByVal s As String) As String
        'PARTE1.PARTE2 (puede set Tabla.Campo ó conexion.Tabla
        'EJ: Tabla.Campo -> Devolvemos Campo
        'campo -> devolvemos campo
        Dim i As Integer
        Dim temp As String = s.Trim.ToUpper
        i = temp.LastIndexOf("."c)
        If i <> NOEXISTE Then
            If i < s.Length - 1 Then
                Return temp.Substring(i + 1)
            Else
                Return VACIO
            End If
        Else
            Return s
        End If
    End Function
    Public Shared Function extraeParte1(ByVal s As String) As String
        'PARTE1.PARTE2 (puede set Tabla.Campo ó conexion.Tabla ó ficherotxt.mapeo ó ficheroexcel.hoj1.mapeo
        'EJ: Tabla.Campo -> Devolvemos Tabla
        'ficheroexcel.hoja1.mapeo -> devolvemos ficheroexcel.hoja1
        'campo -> devolvemos ""
        Dim i As Integer
        Dim temp As String = s.Trim.ToUpper
        i = temp.LastIndexOf("."c)
        If i <> NOEXISTE Then
            Return temp.Substring(0, i)
        Else
            Return VACIO
        End If
    End Function

    Public Shared Function Decodifica_filtro(ByVal s As String) As ColecciondeString
        Dim f As New ColecciondeString
        Dim pfiltro(), parte() As String
        Dim i As Integer

        If s = VACIO Then
            Return Nothing
        End If
        pfiltro = Split(s, PUNTOYCOMA)
        For i = 0 To UBound(pfiltro)
            parte = Split(pfiltro(i), "-")
            f.Add(PasaraCadena(parte(0)), PasaraCadena(parte(1)))
        Next

        Return f
    End Function

    Public Function EjecutaTPR(ByVal conexion As ConexionSQL, ByVal ruta_archivo_tpr As String, ByVal filtro As ColecciondeString, ByVal ficheroconexionesglobales As String, ByVal ficheroconexionesempresa As String, ByVal expr_externa As ColecciondeString, Optional orden As String = VACIO) As DataSet
        Dim consulta As String
        Dim multitabla As Boolean
        Dim con_especifica As ConexionSQL
        Dim conexiones As Colecciondeconexion_de_informe
        Dim existeconexionbase As Boolean

        Try
            enprogreso = True
            If conexion Is Nothing Then
                existeconexionbase = False
                con_especifica = Nothing
            Else
                existeconexionbase = True
                con_especifica = conexion.Clonar()
            End If

            consulta = LeeFicheroConsulta(ruta_archivo_tpr)
            If orden <> VACIO Then consulta = consulta & " " & orden
            If consulta <> VACIO Then
                multitabla = False
                For Each t As String In tokens
                    If InStr(consulta, t) > 0 Then
                        multitabla = True
                        Exit For
                    End If
                Next

                conexiones = New Colecciondeconexion_de_informe
                If multitabla Then

                    If existeconexionbase Then
                        'Usado normalmente en SAP
                        conexiones.Add("", New conexion_de_informe(con_especifica))
                        'Añadimos las conexiones de SBO_COMMON.CON (si existe)
                        'Añadimos las conexiones de [NOMBREBDACTUAL].CON (si existe)                        
                    End If

                    AñadeINclude(consulta, "#automatico1", ficheroconexionesempresa)
                    AñadeINclude(consulta, "#automatico2", ficheroconexionesglobales)

                    MontaMultitabla(conexiones, LeeSentenciasConInclude(consulta, ruta_archivo_tpr), filtro, expr_externa)
                Else
                    If existeconexionbase Then
                        RaiseEvent Inicio(NIVEL_INFORMACION_EVENTO.NIVEL1, 1)
                        MontaMonoTabla(con_especifica, consulta, filtro)
                        RaiseEvent Fin(NIVEL_INFORMACION_EVENTO.NIVEL1)
                    Else
                        Throw New Exception("No es posible ejecutar un TPR simple sin conexion por defecto")
                    End If
                End If

            Else
                Throw New Exception("No se encuentra definicion de Informe")
            End If

        Catch cancelado As Exception_Proceso_ejecucion_TPR_cancelado_por_usuario
            Throw New Exception(cancelado.Message)
        Catch ex As Exception
            Throw New Exception("Error en el informe: " & ruta_archivo_tpr & " " & ex.Message)
        Finally
            enprogreso = False
        End Try

        If existeconexionbase Then
            Return con_especifica.dbDataset
        Else
            If conexiones.Count = 0 Then
                Return Nothing
            Else
                With conexiones.Item(0)
                    Select Case .Tipo
                        Case tipo_conexiones_de_informe.SQL
                            Return .con_sql.dbDataset
                        Case tipo_conexiones_de_informe.MYSQL
                            '   Return .con_mysql.dbdataset
                        Case Else
                            Return Nothing
                    End Select
                End With
            End If
        End If

    End Function

    Private Function LeeSentenciasConInclude(ByVal cad_sql_base As String, ByVal ruta_archivo_tpr As String) As Colecciondesentencia
        Dim ficherosyaprocesados As New ColecciondeString
        Dim directorio_base_tpr As String

        ficherosyaprocesados = New ColecciondeString
        ficherosyaprocesados.Add(ruta_archivo_tpr, ruta_archivo_tpr)
        directorio_base_tpr = System.IO.Path.GetDirectoryName(ruta_archivo_tpr)

        Return LSCI(cad_sql_base, directorio_base_tpr, ficherosyaprocesados)
    End Function

    Private Function LSCI(ByVal cad_sql As String, ByVal directorio_base_tpr As String, ByVal ficherosyaprocesados As ColecciondeString) As Colecciondesentencia
        Dim sentencias, sen_aux, salida As Colecciondesentencia
        Dim ss, saux As sentencia
        Dim nombrefichero As String
        Dim indice, indice_aux As String

        sentencias = LeeSentencias(cad_sql)
        salida = New Colecciondesentencia

        While sentencias.Count > 0
            ss = sentencias.Item(0)
            indice = sentencias.Indice(0)
            If ss.TOKEN = ETIQUETA_INCLUDE Then

                If directorio_base_tpr = "" Then
                    nombrefichero = ss.contenido.Trim
                Else
                    nombrefichero = directorio_base_tpr & "/" & ss.contenido.Trim
                End If



                If Not ficherosyaprocesados.Existe(nombrefichero) Then
                    ficherosyaprocesados.Add(nombrefichero, nombrefichero)
                    sen_aux = LSCI(LeeFicheroConsulta(nombrefichero), directorio_base_tpr, ficherosyaprocesados)
                    While sen_aux.Count > 0
                        saux = sen_aux.Item(0)
                        indice_aux = sen_aux.Indice(0)
                        salida.Add(indice & "-" & indice_aux, saux)
                        sen_aux.Remove(0)
                    End While
                End If
            Else
                salida.Add(indice, ss)
            End If
            sentencias.Remove(0)
        End While
        Return salida
    End Function
    Private Sub MontaMonoTabla(ByVal conexion As ConexionSQL, ByVal consulta As String, ByVal filtro As ColecciondeString)
        consulta = MontaConsulta(Util.SQL_SustCadenas(consulta), consulta, filtro, False)
        Try
            conexion.SelectSQL(consulta, ConexionSQL.NOMBRE_TABLA_EXPORTACION_ESQUEMA_SQL)
        Catch EX As Exception
            Throw New Exception("Consulta: " & consulta & " " & EX.Message)
        End Try
    End Sub
    Public Function MontaConsulta(ByVal base_mayuscula As String, ByVal base As String, ByVal filtro As ColecciondeString, ByVal aplicaalias As Boolean) As String
        Dim result, temp As String
        Dim f As String = VACIO

        If filtro Is Nothing OrElse filtro.Count = 0 Then
            Return base
        End If

        temp = base_mayuscula
        If temp.Contains(sql_UNION) Or temp.Contains(sql_UNION_ALL) Then
            result = AplicaConsultaUnion(temp, base, filtro, aplicaalias)
        Else
            result = AplicaFiltroConsultasimple(temp, base, filtro, aplicaalias)
        End If
        Return result
    End Function

    Private Function AplicaFiltroConsultasimple(ByVal base_mayuscula As String, ByVal base As String, ByVal filtro As ColecciondeString, ByVal aplicaalias As Boolean) As String
        Dim p_select, p_from, p_where, p_groupby, p_having, p_orderby, plw As Integer
        Dim f As String
        Dim fin As Integer
        Dim result As String
        Dim contenido_from As String
        Dim posminima, p1 As Integer
        'Dim sen As SentenciaSQL

        With base_mayuscula
            p_select = .IndexOf(SQL_SELECT)
            p_from = .IndexOf(SQL_FROM)
            p_where = .IndexOf(SQL_WHERE)
            p_orderby = .IndexOf(sql_ORDERBY)
            p_groupby = .IndexOf(sql_GROUPBY)
            p_having = .IndexOf(SQL_HAVING)
        End With

        If aplicaalias Then
            If p_select <> NOEXISTE Then
                If p_from <> NOEXISTE Then
                    fin = Parser_TPR.Minimo_Multiple(p_where, p_orderby, p_groupby, p_having)
                    If fin = NOEXISTE Then
                        fin = base_mayuscula.Length
                    End If
                    contenido_from = " " & base.Substring(p_from, fin - p_from) & " "
                Else
                    contenido_from = VACIO
                End If
            Else
                contenido_from = VACIO
            End If

            f = Serializa_Filtro(filtro, contenido_from)
        Else
            f = Serializa_Filtro(filtro)
        End If
        If f = VACIO Then
            Return base
        End If

        f = entreparen(f)

        posminima = Parser_TPR.Minimo_Multiple(p_groupby, p_having, p_orderby)
        p1 = posminima + 1
        If p_where > NOEXISTE Then
            'ya existe un where, añado como ( lo que habia antes) AND ( expr )
            plw = p_where + Len(SQL_WHERE)

            If posminima > NOEXISTE Then
                result = Left(base, plw) & entreparen(Mid(base, plw, p1 - plw)) & _
                         Y & f & Mid(base, p1)
            Else
                result = Left(base, plw) & entreparen(Mid(base, plw)) & Y & f
            End If
        Else
            'no existe, lo añado con WHERE expr
            'ya existe un where, añado como AND ( expr )

            If posminima > NOEXISTE Then
                result = Left(base, p1) & SQL_WHERE & f & Mid(base, p1)
            Else
                result = base & SQL_WHERE & f
            End If
        End If
        Return result
    End Function

    Private Function AplicaConsultaUnion(ByVal base_mayuscula As String, ByVal cad_sql As String, ByVal filtro As ColecciondeString, ByVal aplicaalias As Boolean) As String
        'dividimos por UNION
        Const Constante_ALL As String = "ALL " ' No la ponemos en conexion porque no debe tener el espacio inicial
        Dim pos As Integer
        Dim s, sbase, result As String
        Dim posactual As Integer
        Dim numeroespaciosiniciales As Integer


        result = VACIO
        pos = base_mayuscula.IndexOf(sql_UNION)

        While pos > NOEXISTE
            If result = VACIO Then
                'Rellenamos tambien la inicial 
                s = cad_sql.Substring(0, pos)
                sbase = base_mayuscula.Substring(0, pos)

                result = AplicaFiltroConsultasimple(" " & sbase & " ", " " & s & " ", filtro, aplicaalias)
            End If

            'Toda UNION tiene al menos dos partes
            posactual = pos + sql_UNION.Length
            pos = base_mayuscula.IndexOf(sql_UNION, pos + 1)
            If pos > NOEXISTE Then
                s = cad_sql.Substring(posactual, pos - posactual)
                sbase = base_mayuscula.Substring(posactual, pos - posactual)
            Else
                s = cad_sql.Substring(posactual)
                sbase = base_mayuscula.Substring(posactual)
            End If

            If LTrim(sbase).StartsWith(Constante_ALL) Then
                numeroespaciosiniciales = 0
                For Each c As Char In sbase
                    If c = " "c Then
                        numeroespaciosiniciales += 1
                    Else
                        Exit For
                    End If
                Next
                s = s.Substring(numeroespaciosiniciales + Len(Constante_ALL))
                sbase = sbase.Substring(numeroespaciosiniciales + Len(Constante_ALL))

                result &= sql_UNION_ALL
            Else
                result &= sql_UNION
            End If

            result &= AplicaFiltroConsultasimple(" " & sbase & " ", " " & s & " ", filtro, aplicaalias)

        End While

        Return result


    End Function


    Public Function LeeFicheroConsulta(ByVal n As String, Optional ByVal limpiarmientraslee As Boolean = True) As String
        Dim ts As System.IO.StreamReader = Nothing
        Dim result As String = VACIO
        Dim l As String

        Try
            If Not System.IO.File.Exists(n) Then
                Return VACIO
            End If
            ts = System.IO.File.OpenText(n)

            If limpiarmientraslee Then
                While Not ts.EndOfStream
                    l = ts.ReadLine.Trim
                    If Not l.StartsWith("--") Then
                        result &= " " & l
                    End If
                End While
                result = SQL_NormalizaExpresion(result)
            Else
                result = ts.ReadToEnd()
            End If
        Catch ex As Exception
            result = VACIO
        Finally
            If Not ts Is Nothing Then
                ts.Close()
            End If
        End Try

        Return result
    End Function

    'IDEA, utilizar las tablas temporales de SQL-SERVER

    'TABLA::nombrequesea::select c.Cardcode, C.CardName
    'into #misclientes
    'from OCRD C

    'TABLA::nombrequesea::select * from #misclientes

    'La primera no devuelve nada en el recordset XML, la segunda lo recupera.
    'En el primer caso no haría falta establecer el nombre de la tabla en el segundo sí.


    Private Sub MontaMultitabla(ByVal conexiones As Colecciondeconexion_de_informe, ByVal col_sentencias As Colecciondesentencia, ByVal filtro As ColecciondeString, ByVal expr_externas As ColecciondeString)
        Dim sentencia, sentenciabase As String
        Dim nt As String
        Dim dt As DataTable
        Dim destino_en_dataset As Boolean
        Dim tablatemporal As Boolean

        Dim ds As sentencia
        Dim i, il As Integer
        Dim j, jconex As Integer
        Dim es_necesario_transaccion As Boolean
        Dim base_mayusculas As String
        Dim nombretabla, nombreconexion As String
        Dim numero_conexiones As Integer = 1
        Dim mapeos As New ColdeColeccion_mapeocampos

        Dim tipoconexionactual As tipo_conexiones_de_informe
        Dim conexionactual_sql As ConexionSQL
        'Dim conexionactual_mysql As ConexionMySQL

        Dim conexionactual_texto As ConexionTexto
        Dim datasetglobal As New DataSet
        Dim conexionactual_excel As ConexionExcel
        Dim listatablastemporales As New ColecciondeString
        Dim factual As Date = Now()
        Dim constantes As New Coleccion_Constantes



        constantes.Combinar(TPR.CalculaExpresionesInternas(factual))
        constantes.Combinar(expr_externas)
        'TABLA::pepito::
        Try
            RaiseEvent Inicio(NIVEL_INFORMACION_EVENTO.NIVEL1, col_sentencias.Count)

            numero_conexiones = conexiones.Count
            If numero_conexiones = 0 Then
                datasetglobal = New DataSet()
            Else
                With conexiones.Item(0)
                    Select Case .Tipo
                        Case tipo_conexiones_de_informe.MYSQL
                            '     datasetglobal = .con_mysql.dbDataset
                        Case tipo_conexiones_de_informe.SQL
                            datasetglobal = .con_sql.dbDataset
                        Case Else
                            datasetglobal = New DataSet()
                    End Select
                End With

            End If


            il = col_sentencias.Count - 1
            es_necesario_transaccion = False

            'Una vez que leemos las conexiones, mapeos, etc...  no hace
            'falta volverlas a leer, por eso las eliminamos
            i = 0
            While i <= il
                ds = col_sentencias.Item(i)
                Select Case ds.TOKEN
                    Case ETIQUETA_COMANDO
                        es_necesario_transaccion = True
                        i = i + 1
                    Case ETIQUETA_CONEXION_SQL, ETIQUETA_CONEXION_TEXTO, ETIQUETA_MAPEO_CAMPOS, ETIQUETA_CONSTANTE, ETIQUETA_CONEXIONMYSQL, ETIQUETA_CONEXION_EXCEL
                        Select Case ds.TOKEN
                            Case ETIQUETA_CONEXION_SQL
                                conexiones.Add(ds.Nombre.ToUpper, New conexion_de_informe(New ConexionSQL(ds.contenido, datasetglobal)))
                                numero_conexiones += 1

                            Case ETIQUETA_CONEXION_TEXTO
                                conexiones.Add(ds.Nombre.ToUpper, New conexion_de_informe(New ConexionTexto(ds.contenido, constantes)))
                                numero_conexiones += 1
                            Case ETIQUETA_CONEXIONMYSQL
                                ' conexiones.Add(ds.Nombre.ToUpper, New conexion_de_informe(New ConexionMySQL(ds.contenido, datasetglobal)))
                                ' numero_conexiones += 1
                            Case ETIQUETA_MAPEO_CAMPOS
                                mapeos.Add(ds.Nombre.ToUpper, New Coleccion_mapeocampos(ds.contenido, constantes))
                            Case ETIQUETA_CONSTANTE
                                'Rellanamos la expresionexterna
                                constantes.Add(ds.contenido)
                            Case ETIQUETA_CONEXION_EXCEL
                                conexiones.Add(ds.Nombre.ToUpper, New conexion_de_informe(New ConexionExcel(Me, ds.contenido, constantes)))
                                numero_conexiones += 1


                        End Select
                        col_sentencias.Remove(i)
                        il = il - 1
                        RaiseEvent Avance(NIVEL_INFORMACION_EVENTO.NIVEL1, 1)
                    Case Else
                        i = i + 1
                End Select
            End While

            jconex = numero_conexiones - 1
            If es_necesario_transaccion Then
                For j = 0 To jconex
                    With conexiones.Item(j)
                        Select Case .Tipo
                            Case tipo_conexiones_de_informe.SQL
                                .con_sql.ComienzaTransaccion()
                            Case tipo_conexiones_de_informe.MYSQL
                                '     .con_mysql.ComienzaTransaccion()
                        End Select
                    End With
                Next
            End If



            For i = 0 To il
                If cancelar Then
                    Throw New Exception("Cancelado Por el usuario")
                End If
                conexionactual_sql = Nothing
                conexionactual_texto = Nothing
                ' conexionactual_mysql = Nothing
                conexionactual_excel = Nothing
                ds = col_sentencias.Item(i)
                nombreconexion = extraeParte1(ds.Nombre)
                nombretabla = extraeParte2(ds.Nombre)
                If conexiones.Existe(nombreconexion) Then
                    tipoconexionactual = conexiones.Item(nombreconexion).Tipo
                    With conexiones.Item(nombreconexion)
                        tipoconexionactual = .Tipo
                        Select Case tipoconexionactual
                            Case tipo_conexiones_de_informe.SQL
                                conexionactual_sql = .con_sql
                            Case tipo_conexiones_de_informe.TEXTO
                                conexionactual_texto = .con_texto
                            Case tipo_conexiones_de_informe.MYSQL
                                'conexionactual_mysql = .con_mysql
                            Case tipo_conexiones_de_informe.EXCEL
                                conexionactual_excel = .con_excel
                            Case Else
                                Throw New Exception("Tipo de conexion aun no definida. Nombre:" & nombreconexion)
                        End Select
                    End With

                Else
                    Throw New Exception("Conexion no encontrada: " & nombreconexion)
                End If
                tablatemporal = nombretabla.ToUpper.StartsWith("#TEMP")
                'PROCESAMOS CADA DIRECTIVA   
                Select Case ds.TOKEN
                    Case ETIQUETA_TABLA
                        Select Case tipoconexionactual
                            Case tipo_conexiones_de_informe.SQL, tipo_conexiones_de_informe.MYSQL

                                base_mayusculas = Util.SQL_SustCadenas(ds.contenido)

                                If tablatemporal Then
                                    sentenciabase = AñadeInto(tipoconexionactual, base_mayusculas, ds.contenido, nombretabla)
                                    destino_en_dataset = False
                                Else
                                    sentenciabase = ds.contenido
                                    destino_en_dataset = datasetglobal.Tables.Contains(nombretabla)
                                End If

                                If destino_en_dataset Then
                                    nt = VACIO
                                Else
                                    nt = nombretabla
                                End If

                                sentencia = MontaConsulta(base_mayusculas, sentenciabase, filtro, True)

                                If tablatemporal Then
                                    Try
                                        'Guardamos la tabla con la conexion ( CONEXION3.#TEMPORAL3)
                                        listatablastemporales.Add(ds.Nombre, ds.Nombre)
                                        Select Case tipoconexionactual
                                            Case tipo_conexiones_de_informe.SQL
                                                conexionactual_sql.Ejecuta(sentencia)
                                            Case tipo_conexiones_de_informe.MYSQL
                                                ' conexionactual_mysql.Ejecuta(sentencia)
                                        End Select


                                    Catch ex As Exception
                                        Throw New Exception("Error en la consulta: " & ds.TOKEN & ds.Nombre & vbCrLf & sentencia & " " & ex.Message)
                                    End Try
                                Else
                                    dt = Nothing
                                    Try
                                        Select Case tipoconexionactual
                                            Case tipo_conexiones_de_informe.SQL
                                                dt = conexionactual_sql.SelectSQL(sentencia, nt)
                                            Case tipo_conexiones_de_informe.MYSQL
                                                ' dt = conexionactual_mysql.SelectSQL(sentencia, nt)
                                        End Select

                                    Catch ex As Exception
                                        Throw New Exception("Error en la consulta: " & ds.TOKEN & ds.Nombre & vbCrLf & sentencia & " " & ex.Message)
                                    End Try

                                    If destino_en_dataset Then
                                        Try
                                            datasetglobal.Tables(nombretabla).Merge(dt)
                                        Catch ex As Exception
                                            Throw New Exception("Error en la consulta (para agregar filas a una tabla existente debe tener la misma estructura): " & ds.TOKEN & " " & ds.Nombre & vbCrLf & sentencia & " " & ex.Message)
                                        End Try
                                        Select Case tipoconexionactual
                                            Case tipo_conexiones_de_informe.SQL
                                                conexionactual_sql.Libera(dt)
                                            Case tipo_conexiones_de_informe.MYSQL
                                                '  conexionactual_mysql.Libera(dt)
                                        End Select
                                    End If
                                End If
                            Case Else
                                Throw New Exception("Solo es posible ejecutar una consulta sobre una conexion SQL." & ds.TOKEN & ds.Nombre)
                        End Select

                    Case ETIQUETA_COMANDO

                        Select Case tipoconexionactual
                            Case tipo_conexiones_de_informe.SQL
                                Try
                                    conexionactual_sql.Ejecuta(ds.contenido)
                                Catch ex As Exception
                                    Throw New Exception("Error en el comando: " & ds.TOKEN & ds.Nombre & ds.contenido & " " & ex.Message)
                                End Try
                            Case tipo_conexiones_de_informe.MYSQL
                                Try
                                    'conexionactual_mysql.Ejecuta(ds.contenido)
                                Catch ex As Exception
                                    Throw New Exception("Error en el comando: " & ds.TOKEN & ds.Nombre & ds.contenido & " " & ex.Message)
                                End Try
                            Case Else
                                Throw New Exception("Solo es posible ejecutar un comando sobre una conexion (SQL/MySQL)." & ds.TOKEN & ds.Nombre)

                        End Select


                    Case ETIQUETA_TRANSFERIR
                        Dim conexiondestino As String
                        Dim nombredestino As String 'puede ser el nombre de una tabla o un mapeo
                        Dim tipoconexiondestino As tipo_conexiones_de_informe
                        Dim nombrehojaexcel As String = String.empty

                        'Varios casos:

                        ' 1º ConexionSQL - ConexionSQL
                        ' 2º ConexionSQL - conexionTexto
                        ' 3º ConexionTexto - ConexionSQL
                        ' 4º ConexionTexto -Contexion Texto

                        Try
                            conexiondestino = extraeParte1(ds.contenido)
                            nombredestino = extraeParte2(ds.contenido)

                            If conexiondestino.Contains(".") Then
                                'Por si ds.contenido es  conexionexcel.hoja2.mapeo
                                nombrehojaexcel = extraeParte2(conexiondestino)
                                conexiondestino = extraeParte1(conexiondestino)
                            End If


                            If Not conexiones.Existe(conexiondestino) Then
                                Throw New Exception("Conexion no encontrada: " & nombreconexion)
                            End If

                            tipoconexiondestino = conexiones.Item(conexiondestino).Tipo
                            dt = Nothing
                            Select Case tipoconexionactual
                                Case tipo_conexiones_de_informe.SQL, tipo_conexiones_de_informe.MYSQL
                                    Select Case tipoconexionactual
                                        Case tipo_conexiones_de_informe.SQL
                                            dt = conexionactual_sql.SelectSQL(Util.sql_SELECT_a_FROM & nombretabla)
                                        Case tipo_conexiones_de_informe.MYSQL
                                            '      dt = conexionactual_mysql.SelectSQL(Util.sql_SELECT_a_FROM & nombretabla)
                                    End Select

                                    Select Case tipoconexiondestino
                                        Case tipo_conexiones_de_informe.SQL
                                            'PROBLEMA, la tabla destino debe existir y tener la misma estructura.
                                            Try
                                                conexiones.Item(conexiondestino).con_sql.BulkCopy(dt, nombredestino)
                                            Catch ex As Exception
                                                Throw New Exception(nombretabla & " -> " & nombredestino & " " & ex.Message)
                                            End Try
                                        Case tipo_conexiones_de_informe.MYSQL
                                            Try
                                                '   conexiones.Item(conexiondestino).con_mysql.BulkCopy(dt, nombredestino)
                                            Catch ex As Exception
                                                Throw New Exception(nombretabla & " -> " & nombredestino & " " & ex.Message)
                                            End Try
                                        Case tipo_conexiones_de_informe.TEXTO
                                            'El nombre del mapeo es el de nombredestino
                                            If mapeos.Existe(nombredestino) Then
                                                conexiones.Item(conexiondestino).con_texto.Escribe(dt, mapeos.Item(nombredestino))
                                            Else
                                                Throw New Exception("No se encuentra el mapeo de campos: " & nombredestino)
                                            End If

                                        Case tipo_conexiones_de_informe.EXCEL
                                            If nombrehojaexcel = VACIO Then
                                                Throw New Exception("Debe especificarse el nombre de la hoja excel")
                                            End If


                                            If mapeos.Existe(nombredestino) Then
                                                conexiones.Item(conexiondestino).con_excel.Escribe(dt, mapeos.Item(nombredestino), nombrehojaexcel)

                                            Else
                                                Throw New Exception("No se encuentra el mapeo de campos: " & nombredestino)
                                            End If

                                    End Select
                                    Select Case tipoconexionactual
                                        Case tipo_conexiones_de_informe.SQL
                                            conexionactual_sql.Libera(dt)
                                        Case tipo_conexiones_de_informe.MYSQL
                                            '  conexionactual_mysql.Libera(dt)
                                    End Select

                                Case tipo_conexiones_de_informe.TEXTO
                                    Select Case tipoconexiondestino
                                        Case tipo_conexiones_de_informe.SQL
                                            If mapeos.Existe(nombretabla) Then
                                                conexiones.Item(nombreconexion).con_texto.Lee(conexiones.Item(conexiondestino).con_sql, nombredestino, mapeos.Item(nombretabla))
                                            Else
                                                Throw New Exception("No se encuentra el mapeo de campos: " & nombretabla)
                                            End If
                                        Case tipo_conexiones_de_informe.TEXTO
                                            '¿?¿?¿?¿?¿?¿?¿?
                                    End Select

                                Case tipo_conexiones_de_informe.EXCEL
                                    Select Case tipoconexiondestino
                                        Case tipo_conexiones_de_informe.SQL
                                            Throw New Exception("No se encuentra la opción de excel: " & nombredestino & ".")
                                        Case Else
                                            Throw New Exception("No se ha definido TRANSFERIR entre Excel y la conexion " & conexiondestino)
                                    End Select



                            End Select

                        Catch ex As Exception_Proceso_ejecucion_TPR_cancelado_por_usuario
                            Throw New Exception_Proceso_ejecucion_TPR_cancelado_por_usuario(ex.Message)
                        Catch ex As Exception
                            Throw New Exception("Error al TRANSFERIR: " & ex.Message)
                        End Try
                End Select
                RaiseEvent Avance(NIVEL_INFORMACION_EVENTO.NIVEL1, 1)
            Next


            If es_necesario_transaccion Then
                For j = 0 To jconex
                    With conexiones.Item(j)
                        Select Case .Tipo
                            Case tipo_conexiones_de_informe.SQL
                                .con_sql.TerminaTransaccion(True)
                            Case tipo_conexiones_de_informe.MYSQL
                                ' .con_mysql.TerminaTransaccion(True)
                        End Select
                    End With
                Next
            End If

            For j = 0 To jconex
                With conexiones.Item(j)
                    Select Case .Tipo
                        Case tipo_conexiones_de_informe.SQL
                            .con_sql.CerrarConexion()
                        Case tipo_conexiones_de_informe.MYSQL
                            ' .con_mysql.CerrarConexion()
                        Case tipo_conexiones_de_informe.EXCEL
                            .con_excel.Cerrar()
                    End Select
                End With
            Next

        Catch ex As Exception_Proceso_ejecucion_TPR_cancelado_por_usuario

            For j = 0 To jconex
                With conexiones.Item(j)
                    Select Case .Tipo
                        Case tipo_conexiones_de_informe.SQL
                            If es_necesario_transaccion Then
                                .con_sql.TerminaTransaccion(False)
                            End If
                        Case tipo_conexiones_de_informe.EXCEL
                            .con_excel.CerrarAbortando()
                    End Select
                End With
            Next
            Throw New Exception_Proceso_ejecucion_TPR_cancelado_por_usuario(ex.Message)

        Catch ex As Exception
            If es_necesario_transaccion Then
                For j = 0 To jconex
                    With conexiones.Item(j)
                        Select Case .Tipo
                            Case tipo_conexiones_de_informe.SQL
                                .con_sql.TerminaTransaccion(False)
                            Case tipo_conexiones_de_informe.EXCEL
                                .con_excel.CerrarAbortando()

                            Case tipo_conexiones_de_informe.MYSQL
                                '   .con_mysql.TerminaTransaccion(False)
                        End Select
                    End With
                Next
            End If

            Throw New Exception(ex.Message)
        Finally
            'Borramos las tablas temporales

            With listatablastemporales
                While .Count > 0
                    nombreconexion = extraeParte1(listatablastemporales.Item(0))
                    nombretabla = extraeParte2(listatablastemporales.Item(0))

                    With conexiones.Item(nombreconexion)
                        Select Case .Tipo
                            Case tipo_conexiones_de_informe.SQL
                                If .con_sql.ExisteTablaSQL(nombretabla) Then
                                    .con_sql.Ejecuta("DROP TABLE " & nombretabla)
                                End If
                            Case tipo_conexiones_de_informe.MYSQL
                                '   .con_mysql.Ejecuta("DROP TEMPORARY TABLE IF EXISTS " & nombretabla.Substring(1))

                        End Select

                    End With

                    .Remove(0)
                End While
            End With

            RaiseEvent Fin(NIVEL_INFORMACION_EVENTO.NIVEL1)

        End Try
    End Sub


    Public Function DEBUG_EjecutaTPR(ByVal conexion As ConexionSQL, ByVal ruta_archivo_tpr As String, ByVal filtro As ColecciondeString, ByVal ficheroconexionesglobales As String, ByVal ficheroconexionesempresa As String, ByVal expr_externa As ColecciondeString) As String
        Dim consulta As String
        Dim multitabla As Boolean
        Dim con_especifica As ConexionSQL
        Dim conexiones As Colecciondeconexion_de_informe

        Dim cadena_resultado As String = String.empty

        Try
            con_especifica = conexion.Clonar()
            consulta = LeeFicheroConsulta(ruta_archivo_tpr)
            If consulta <> VACIO Then
                multitabla = False
                For Each t As String In tokens
                    If InStr(consulta, t) > 0 Then
                        multitabla = True
                        Exit For
                    End If
                Next


                If multitabla Then
                    conexiones = New Colecciondeconexion_de_informe
                    conexiones.Add("", New conexion_de_informe(con_especifica))
                    'Añadimos las conexiones de SBO_COMMON.CON (si existe)
                    'Añadimos las conexiones de [NOMBREBDACTUAL].CON (si existe)
                    LeeConexion(ficheroconexionesglobales, conexiones, con_especifica.dbDataset)
                    LeeConexion(ficheroconexionesempresa, conexiones, con_especifica.dbDataset)

                    cadena_resultado = DEBUG_MontaMultitabla(conexiones, LeeSentenciasConInclude(consulta, ruta_archivo_tpr), filtro, expr_externa)
                Else
                    cadena_resultado = DEBUG_MontaMonoTabla(con_especifica, consulta, filtro)
                End If

            Else
                Throw New Exception("No se encuentra definicion de Informe")
            End If


        Catch ex As Exception
            Throw New Exception("Error en el informe: " & ruta_archivo_tpr & " " & ex.Message)
        End Try

        Return cadena_resultado
    End Function

    Private Function DEBUG_MontaMonoTabla(ByVal conexion As ConexionSQL, ByVal consulta As String, ByVal filtro As ColecciondeString) As String
        consulta = MontaConsulta(Util.SQL_SustCadenas(consulta), consulta, filtro, False)
        Try
            Return consulta
        Catch EX As Exception
            Throw New Exception("Consulta: " & consulta & " " & EX.Message)
        End Try
    End Function

    Private Function DEBUG_MontaMultitabla(ByVal conexiones As Colecciondeconexion_de_informe, ByVal col_sentencias As Colecciondesentencia, ByVal filtro As ColecciondeString, ByVal expr_externas As ColecciondeString) As String
        Dim sentencia, sentenciabase As String
        Dim nt As String
        Dim dt As DataTable
        Dim destino_en_dataset As Boolean
        Dim tablatemporal As Boolean

        Dim ds As sentencia
        Dim i, il As Integer
        Dim j, jconex As Integer
        ' Dim es_necesario_transaccion As Boolean
        Dim base_mayusculas As String
        Dim nombretabla, nombreconexion As String
        Dim numero_conexiones As Integer = 1
        Dim mapeos As New ColdeColeccion_mapeocampos

        Dim tipoconexionactual As tipo_conexiones_de_informe
        Dim conexionactual_sql, conexion0 As ConexionSQL
        Dim conexionactual_texto As ConexionTexto
        Dim listatablastemporales As New ColecciondeString
        Dim factual As Date = Now()
        Dim expr_internas As New ColecciondeString
        Dim constantes As New Coleccion_Constantes

        Dim cadena_resultado As String = String.empty



        expr_internas = TPR.CalculaExpresionesInternas(factual)
        constantes.Combinar(expr_internas)
        constantes.Combinar(expr_externas)

        'TABLA::pepito::
        Try
            numero_conexiones = conexiones.Count
            If numero_conexiones > 0 Then
                conexion0 = conexiones.Item("").con_sql
            Else
                Return ""
            End If

            il = col_sentencias.Count - 1


            'Una vez que leemos las conexiones, mapeos, etc...  no hace
            'falta volverlas a leer, por eso las eliminamos
            i = 0
            While i <= il
                ds = col_sentencias.Item(i)
                Select Case ds.TOKEN
                    Case ETIQUETA_COMANDO
                        'es_necesario_transaccion = True
                        i = i + 1
                    Case ETIQUETA_CONEXION_SQL, ETIQUETA_CONEXION_TEXTO, ETIQUETA_MAPEO_CAMPOS, ETIQUETA_CONSTANTE
                        Select Case ds.TOKEN
                            Case ETIQUETA_CONEXION_SQL
                                conexiones.Add(ds.Nombre.ToUpper, New conexion_de_informe(New ConexionSQL(ds.contenido, conexion0.dbDataset)))
                                numero_conexiones += 1

                            Case ETIQUETA_CONEXION_TEXTO
                                conexiones.Add(ds.Nombre.ToUpper, New conexion_de_informe(New ConexionTexto(ds.contenido, constantes)))
                                numero_conexiones += 1
                            Case ETIQUETA_MAPEO_CAMPOS
                                mapeos.Add(ds.Nombre.ToUpper, New Coleccion_mapeocampos(ds.contenido, constantes))
                            Case ETIQUETA_CONSTANTE
                                constantes.Add(ds.contenido)

                        End Select
                        col_sentencias.Remove(i)
                        il = il - 1
                    Case Else
                        i = i + 1
                End Select
            End While

            'jconex = numero_conexiones - 1
            'If es_necesario_transaccion Then
            '    For j = 0 To jconex
            '        With conexiones.Item(j)
            '            If .Tipo = tipo_conexiones_de_informe.SQL Then
            '                .con_sql.ComienzaTransaccion()
            '            End If
            '        End With
            '    Next
            'End If

            For i = 0 To il
                conexionactual_sql = Nothing
                conexionactual_texto = Nothing
                ds = col_sentencias.Item(i)
                nombreconexion = extraeParte1(ds.Nombre)
                nombretabla = extraeParte2(ds.Nombre)
                If conexiones.Existe(nombreconexion) Then
                    tipoconexionactual = conexiones.Item(nombreconexion).Tipo
                    Select Case tipoconexionactual
                        Case tipo_conexiones_de_informe.SQL
                            conexionactual_sql = conexiones.Item(nombreconexion).con_sql
                        Case tipo_conexiones_de_informe.TEXTO
                            conexionactual_texto = conexiones.Item(nombreconexion).con_texto
                    End Select

                Else
                    Throw New Exception("Conexion no encontrada: " & nombreconexion)
                End If
                tablatemporal = nombretabla.ToUpper.StartsWith("#TEMP")
                'PROCESAMOS CADA DIRECTIVA   
                Select Case ds.TOKEN
                    Case ETIQUETA_TABLA
                        If tipoconexionactual = tipo_conexiones_de_informe.SQL Then
                            base_mayusculas = Util.SQL_SustCadenas(ds.contenido)

                            If tablatemporal Then
                                sentenciabase = AñadeInto(tipoconexionactual, base_mayusculas, ds.contenido, nombretabla)
                                destino_en_dataset = False
                            Else
                                sentenciabase = ds.contenido
                                'En el caso de varios tipos de conexions habrá que hacer un case por tipo
                                destino_en_dataset = conexionactual_sql.dbDataset.Tables.Contains(nombretabla)
                            End If

                            If destino_en_dataset Then
                                nt = VACIO
                            Else
                                nt = nombretabla
                            End If

                            sentencia = MontaConsulta(base_mayusculas, sentenciabase, filtro, True)

                            If tablatemporal Then
                                Try
                                    'Guardamos la tabla con la conexion ( CONEXION3.#TEMPORAL3)
                                    listatablastemporales.Add(ds.Nombre, ds.Nombre)

                                    conexionactual_sql.Ejecuta(sentencia)
                                    cadena_resultado &= Linea("-- CONSULTA: " & ds.Nombre)
                                    cadena_resultado &= Linea(sentencia)
                                    cadena_resultado &= Linea("-- fin CONSULTA : " & ds.Nombre)


                                Catch ex As Exception
                                    Throw New Exception("Error en la consulta: " & ds.TOKEN & ds.Nombre & vbCrLf & sentencia & " " & ex.Message)
                                End Try
                            Else
                                Try
                                    dt = conexionactual_sql.SelectSQL(sentencia, nt)
                                    cadena_resultado &= Linea("-- CONSULTA: " & ds.Nombre)
                                    cadena_resultado &= Linea("-- El resultado de la consulta  se enviará a Crystal Reports")
                                    cadena_resultado &= Linea(sentencia)
                                    cadena_resultado &= Linea("-- fin CONSULTA : " & ds.Nombre)

                                Catch ex As Exception
                                    Throw New Exception("Error en la consulta: " & ds.TOKEN & ds.Nombre & vbCrLf & sentencia & " " & ex.Message)
                                End Try

                                If destino_en_dataset Then
                                    Try
                                        conexionactual_sql.dbDataset.Tables(nombretabla).Merge(dt)
                                        cadena_resultado &= Linea("--Se añade a la tabla existente " & dt.TableName)
                                    Catch ex As Exception
                                        Throw New Exception("Error en la consulta (para agregar filas a una tabla existente debe tener la misma estructura): " & ds.TOKEN & " " & ds.Nombre & vbCrLf & sentencia & " " & ex.Message)
                                    End Try

                                    conexionactual_sql.Libera(dt)
                                End If
                            End If

                        Else
                            Throw New Exception("Solo es posible ejecutar una consulta sobre una conexion SQL." & ds.TOKEN & ds.Nombre)
                        End If
                    Case ETIQUETA_COMANDO
                        If tipoconexionactual = tipo_conexiones_de_informe.SQL Then
                            Try
                                ' conexionactual_sql.Ejecuta(ds.contenido)
                                cadena_resultado &= Linea("--COMANDO " & ds.Nombre)
                                cadena_resultado &= Linea(ds.contenido)
                                cadena_resultado &= Linea("-- fin COMANDO " & ds.Nombre)
                            Catch ex As Exception
                                Throw New Exception("Error en el comando: " & ds.TOKEN & ds.Nombre & ds.contenido & " " & ex.Message)
                            End Try
                        Else
                            Throw New Exception("Solo es posible ejecutar un comando sobre una conexion SQL." & ds.TOKEN & ds.Nombre)
                        End If

                    Case ETIQUETA_TRANSFERIR
                        Dim conexiondestino As String
                        Dim nombredestino As String 'puede ser el nombre de una tabla o un mapeo
                        Dim tipoconexiondestino As tipo_conexiones_de_informe

                        'Varios casos:

                        ' 1º ConexionSQL - ConexionSQL
                        ' 2º ConexionSQL - conexionTexto
                        ' 3º ConexionTexto - ConexionSQL
                        ' 4º ConexionTexto -Contexion Texto
                        Try
                            conexiondestino = extraeParte1(ds.contenido)
                            nombredestino = extraeParte2(ds.contenido)

                            If Not conexiones.Existe(conexiondestino) Then
                                Throw New Exception("Conexion no encontrada: " & nombreconexion)
                            End If

                            tipoconexiondestino = conexiones.Item(conexiondestino).Tipo

                            Select Case tipoconexionactual
                                Case tipo_conexiones_de_informe.SQL
                                    dt = conexionactual_sql.SelectSQL(Util.sql_SELECT_a_FROM & nombretabla)
                                    Select Case tipoconexiondestino
                                        Case tipo_conexiones_de_informe.SQL
                                            'PROBLEMA, la tabla destino debe existir y tener la misma estructura.
                                            Try
                                                ' conexiones.Item(conexiondestino).con_sql.BulkCopy(dt, nombredestino)
                                                cadena_resultado &= Linea("--TRANSFERIR entre conexiones SQL-SQL: " & nombretabla & " -> " & nombredestino)
                                            Catch ex As Exception
                                                Throw New Exception(nombretabla & " -> " & nombredestino & " " & ex.Message)
                                            End Try
                                            conexionactual_sql.Libera(dt)
                                        Case tipo_conexiones_de_informe.TEXTO
                                            'El nombre del mapeo es el de nombredestino
                                            If mapeos.Existe(nombredestino) Then
                                                'conexiones.Item(conexiondestino).con_texto.Escribe(dt, mapeos.Item(nombredestino))
                                                cadena_resultado &= Linea("--TRANSFERIR entre  SQL-TEXTO: " & nombretabla & " -> " & nombredestino & "." & conexiondestino)
                                            Else
                                                Throw New Exception("No se encuentra el mapeo de campos: " & nombredestino)
                                            End If

                                    End Select

                                Case tipo_conexiones_de_informe.TEXTO
                                    Select Case tipoconexiondestino
                                        Case tipo_conexiones_de_informe.SQL
                                            If mapeos.Existe(nombretabla) Then
                                                'conexiones.Item(nombreconexion).con_texto.Lee(conexiones.Item(conexiondestino).con_sql, nombredestino, mapeos.Item(nombretabla))
                                                cadena_resultado &= Linea("--TRANSFERIR entre  TEXTO-SQL: " & nombretabla & " -> " & nombredestino & "." & conexiondestino)
                                            Else
                                                Throw New Exception("No se encuentra el mapeo de campos: " & nombretabla)
                                            End If
                                        Case tipo_conexiones_de_informe.TEXTO
                                            '¿?¿?¿?¿?¿?¿?¿?
                                    End Select
                            End Select
                        Catch ex As Exception
                            Throw New Exception("Error al TRANFERIR: " & ex.Message)
                        End Try
                End Select
            Next


            'If es_necesario_transaccion Then
            '    For j = 0 To jconex
            '        With conexiones.Item(j)
            '            If .Tipo = tipo_conexiones_de_informe.SQL Then
            '                .con_sql.TerminaTransaccion(False)
            '            End If
            '        End With
            '    Next
            'End If

            For j = 0 To jconex
                With conexiones.Item(j)
                    If .Tipo = tipo_conexiones_de_informe.SQL Then
                        .con_sql.dbConnection.Close()
                    End If
                End With
            Next

        Catch ex As Exception
            'If es_necesario_transaccion Then
            '    For j = 0 To jconex
            '        With conexiones.Item(j)
            '            If .Tipo = tipo_conexiones_de_informe.SQL Then
            '                .con_sql.TerminaTransaccion(False)
            '            End If
            '        End With
            '    Next
            'End If

            Throw New Exception(ex.Message)
        Finally
            'Borramos las tablas temporales

            With listatablastemporales
                While .Count > 0
                    nombreconexion = extraeParte1(listatablastemporales.Item(0))
                    nombretabla = extraeParte2(listatablastemporales.Item(0))

                    With conexiones.Item(nombreconexion).con_sql

                        cadena_resultado = Linea("DROP TABLE " & nombretabla) & cadena_resultado
                        'If .ExisteTablaSQL(nombretabla) Then
                        '    .Ejecuta("DROP TABLE " & nombretabla)
                        'End If
                    End With

                    .Remove(0)
                End While
            End With

        End Try

        Return cadena_resultado
    End Function

    Private Function LeeSentencias(ByVal cad As String) As Colecciondesentencia
        Dim pos As Integer
        Dim cs As Colecciondesentencia
        Dim resultado, listaelegida As New Colecciondesentencia
        Dim conjunto_parcial As New ColdeColdesentencia

        Dim ds As sentencia
        Dim t As String
        Dim i, p, il As Integer
        Dim quedaalguno As Boolean
        Dim s As String



        Dim numerotokens As Integer
        Dim minimo, posbase As Integer

        numerotokens = tokens.Length

        For Each t In tokens
            pos = cad.IndexOf(t)
            cs = New Colecciondesentencia
            While pos <> NOEXISTE
                'marcamos la posicion
                ds = New sentencia
                ds.TOKEN = t
                'ds.cadsql =  (aun no podemos saber el final de la cadena
                'ds.Nombre= (aun no podemos saber el final de la cadena
                ds.posicion = pos
                cs.Add(pos.ToString, ds)
                pos = cad.IndexOf(t, pos + 1)
            End While
            conjunto_parcial.Add(cs)
        Next

        'Ahora reordenamos 
        il = conjunto_parcial.Count - 1

        quedaalguno = True
        While quedaalguno
            quedaalguno = False
            minimo = -1
            ds = Nothing
            listaelegida = Nothing
            For i = 0 To il 'recorremos las listas de todos los tokens
                cs = conjunto_parcial.Item(i)
                If cs.Count > 0 Then
                    ds = cs.Item(0)
                    If minimo = -1 Or minimo > ds.posicion Then
                        listaelegida = cs
                        minimo = ds.posicion
                    End If
                    quedaalguno = True
                End If
            Next
            If minimo > -1 Then
                resultado.Add(minimo.ToString, listaelegida.Item(0))
                listaelegida.Remove(0)
            End If
        End While

        conjunto_parcial.Clear()

        il = resultado.Count - 1
        For i = 0 To il
            ds = resultado.Item(i)
            posbase = ds.posicion + ds.TOKEN.Length
            If i = il Then   'No hay siguiente
                s = cad.Substring(posbase)
            Else 'existe un siguiente
                s = cad.Substring(posbase, resultado.Item(i + 1).posicion - posbase)
            End If
            p = InStr(s, CUATROPUNTOS)
            If p > 0 Then
                ds.contenido = " " & s.Substring(p + CUATROPUNTOS.Length - 1)
                ds.Nombre = s.Substring(0, p - 1)
            Else
                Throw New Exception("Falta Delimitador (::) en :" & s)
            End If
        Next
        Return resultado
    End Function

    Private Sub AñadeINclude(ByRef consulta As String, ByVal nombreinclude As String, ByVal nombreyrutafichero As String)
        If nombreyrutafichero <> "" AndAlso System.IO.File.Exists(nombreyrutafichero) Then
            consulta = " " & ETIQUETA_INCLUDE & nombreinclude & CUATROPUNTOS & nombreyrutafichero & " " & consulta
        End If
    End Sub
    Private Sub LeeConexion(ByVal nombreyrutafichero As String, ByVal conexiones As Colecciondeconexion_de_informe, ByVal dbdataset_existente As Data.DataSet)
        Dim cadsql As String
        Dim i, il As Integer
        Dim col As Colecciondesentencia
        Dim ds As sentencia

        If System.IO.File.Exists(nombreyrutafichero) Then
            cadsql = LeeFicheroConsulta(nombreyrutafichero, True)
            col = LeeSentencias(cadsql)
            il = col.Count - 1

            For i = 0 To il
                ds = col.Item(i)
                If ds.TOKEN = ETIQUETA_CONEXION_SQL Then 'IGNORAMOS LO QUE NO SEA CONEXION
                    conexiones.Add(ds.Nombre.ToUpper, New conexion_de_informe(New ConexionSQL(ds.contenido, dbdataset_existente)))
                End If
            Next
        End If

    End Sub
    Private Function AñadeInto(ByVal tipoconexionactual As tipo_conexiones_de_informe, ByRef base_mayuscula As String, ByVal s As String, ByVal nombretabla As String) As String
        'OJO que tiene un EFECTO COLATERAL EN base_mayuscula, también la cambia para que sea igual
        Dim p_from As Integer

        'temp =   Util.SQL_SustCadenas(s)
        Select Case tipoconexionactual
            Case tipo_conexiones_de_informe.SQL
                p_from = base_mayuscula.IndexOf(SQL_FROM)

                If p_from <> NOEXISTE Then
                    base_mayuscula = base_mayuscula.Substring(0, p_from) & Util.SQL_INTO & nombretabla & base_mayuscula.Substring(p_from)
                    Return s.Substring(0, p_from) & Util.SQL_INTO & nombretabla & s.Substring(p_from)
                Else
                    base_mayuscula = base_mayuscula & Util.SQL_INTO & nombretabla
                    Return s & Util.SQL_INTO & nombretabla


                End If

            Case tipo_conexiones_de_informe.MYSQL
                base_mayuscula = "CREATE TEMPORARY TABLE " & nombretabla.Substring(1) & " " & base_mayuscula
                Return "CREATE TEMPORARY TABLE " & nombretabla.Substring(1) & " " & s
            Case Else
                Throw New Exception("Añade Insert Into " & nombretabla & " conexion no soportada")
        End Select



    End Function
    'Private Function Esconsultainterna(ByVal s As String, ByVal conexion As ConexionSQL) As Boolean
    '    Dim fragmento_from As String

    '    fragmento_from = extrae_from(s)

    '    For Each dt As DataTable In conexion.dbDataset.Tables
    '        If fragmento_from.IndexOf(" " & dt.TableName.ToUpper & " ") <> NOEXISTE Then
    '            Return True
    '        End If
    '    Next

    '    Return False

    '    'En Los casos "raros", como que no existe select, o no exista from,
    '    'devolvemos false, para que se ejecute en el servidor sql server, y devolverá un mensaje de error

    'End Function

    'Private Function extrae_from(ByVal s As String) As String
    '    Dim temp As String
    '    Dim p_select, p_from, p_where, p_groupby, p_having, p_orderby, p_union, f As Integer

    '    'temp = Util.SQL_SustCadenas(s)
    '    temp = s

    '    p_union = temp.IndexOf(sql_UNION)
    '    If p_union = NOEXISTE Then
    '        p_select = temp.IndexOf(SQL_SELECT)
    '        p_from = temp.IndexOf(SQL_FROM)
    '        p_where = temp.IndexOf(SQL_WHERE)
    '        p_orderby = temp.IndexOf(sql_ORDERBY)
    '        p_groupby = temp.IndexOf(sql_GROUPBY)
    '        p_having = temp.IndexOf(SQL_HAVING)

    '        If p_select <> NOEXISTE Then
    '            If p_from <> NOEXISTE Then
    '                If p_where <> NOEXISTE And p_where > p_from Then
    '                    f = p_where
    '                ElseIf p_orderby <> NOEXISTE And p_orderby > p_from Then
    '                    f = p_orderby
    '                ElseIf p_groupby <> NOEXISTE And p_groupby > p_from Then
    '                    f = p_groupby
    '                Else
    '                    f = temp.Length
    '                End If
    '                Return " " & temp.Substring(p_from, f - p_from) & " "
    '            Else
    '                Return VACIO
    '            End If
    '        Else
    '            Return VACIO
    '        End If
    '    Else
    '        Return extrae_from(" " & temp.Substring(0, p_union) & " ")
    '    End If
    'End Function

    Private Class sentencia
        Public TOKEN As String
        Public Nombre As String
        Public contenido As String
        'Esta es la posicion en la que se encuentra el token (......TOKEN::NOMBRE::SENTENCIA)
        Public posicion As Integer
    End Class

    Private Class Colecciondesentencia
        Inherits coleeccion_object


        <System.Diagnostics.DebuggerStepThrough()> Public Overloads Sub Add(ByVal nombre As String, ByVal valor As sentencia)
            MyBase.Add(nombre, valor)
        End Sub

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function Item(ByVal index As Integer) As sentencia
            Return DirectCast(MyBase.Item(index).valor, sentencia)
        End Function

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function ITem(ByVal nombre As String) As sentencia
            Return DirectCast(MyBase.Item(nombre).valor, sentencia)
        End Function
    End Class
    Private Class ColdeColdesentencia
        Inherits coleeccion_object


        <System.Diagnostics.DebuggerStepThrough()> Public Overloads Sub Add(ByVal valor As Colecciondesentencia)
            MyBase.Add((MyBase.Count + 1).ToString, valor)
        End Sub

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function Item(ByVal index As Integer) As Colecciondesentencia
            Return DirectCast(MyBase.Item(index).valor, Colecciondesentencia)
        End Function

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function ITem(ByVal nombre As String) As Colecciondesentencia
            Return DirectCast(MyBase.Item(nombre).valor, Colecciondesentencia)
        End Function
    End Class

    Enum tipo_conexiones_de_informe
        SQL
        'Cuando vayamos añadiendo mas tipos de conexiones las colocamos aqui
        TEXTO
        EXCEL
        MYSQL
    End Enum
    Private Class conexion_de_informe
        Public Tipo As tipo_conexiones_de_informe
        Public con_sql As ConexionSQL
        Public con_texto As ConexionTexto
        Public con_excel As ConexionExcel
        'Public con_mysql As ConexionMySQL
        'Cuando vayamos añadiendo mas tipos de conexiones las colocamos aqui

        '--------

        Public Sub New(ByVal v_con As ConexionSQL)
            Tipo = tipo_conexiones_de_informe.SQL
            con_sql = v_con
        End Sub

        Public Sub New(ByVal v_con As ConexionTexto)
            Tipo = tipo_conexiones_de_informe.TEXTO
            con_texto = v_con
        End Sub

        'Public Sub New(ByVal v_con As ConexionMySQL)
        '    Tipo = tipo_conexiones_de_informe.MYSQL
        '    con_mysql = v_con
        'End Sub

        Public Sub New(ByVal v_con As ConexionExcel)
            Tipo = tipo_conexiones_de_informe.EXCEL
            con_excel = v_con
        End Sub


    End Class
    Private Class Colecciondeconexion_de_informe
        Inherits coleeccion_object


        <System.Diagnostics.DebuggerStepThrough()> Public Overloads Sub Add(ByVal nombre As String, ByVal valor As conexion_de_informe)
            MyBase.Add(nombre, valor)
        End Sub

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function Item(ByVal index As Integer) As conexion_de_informe
            Return DirectCast(MyBase.Item(index).valor, conexion_de_informe)
        End Function

        <System.Diagnostics.DebuggerStepThrough()> Public Shadows Function ITem(ByVal nombre As String) As conexion_de_informe
            Return DirectCast(MyBase.Item(nombre).valor, conexion_de_informe)
        End Function
    End Class
End Class

