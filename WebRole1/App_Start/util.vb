Option Strict Off
Imports System.Data
Imports System.Web.UI.Page
Imports System.Web
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Diagnostics.Debug
Imports System.Text



Public Module Util
    Public DIRECTORIOTEMPORAL As String
    Public Const sSecretKey As String = ",.l;@A9&"

    Public Const _Y_ As String = " AND "
    Public Const _O_ As String = " OR "
    Public Const _NOT_ As String = " NOT "
    Public Const _EXISTS_ As String = " EXISTS "
    Public Const _NOT_EXISTS_ As String = _NOT_ & _EXISTS_
    Public Const SQL_SELECT As String = " SELECT "
    Public Const SQL_WHERE As String = " WHERE "
    Public Const SQL_FROM As String = " FROM "
    Public Const SQL_HAVING As String = " HAVING "
    Public Const sql_ORDERBY As String = " ORDER BY "
    Public Const sql_GROUPBY As String = " GROUP BY "

    Public Const sql_INNER_JOIN As String = " INNER JOIN "
    Public Const sql_LEFT_OUTER_JOIN As String = " LEFT OUTER JOIN "
    Public Const sql_RIGHT_OUTER_JOIN As String = " RIGHT OUTER JOIN "
    Public Const sql_ON As String = " ON "
    Public Const sql_AS As String = " AS "
    Public Const sql_UNION As String = " UNION "
    Public Const sql_UNION_ALL As String = sql_UNION & "ALL "
    Public Const SQL_NULL As String = " NULL "
    Public Const SQL_LIKE As String = " LIKE "
    Public Const SQL_NOT_LIKE As String = " NOT LIKE "
    Public Const SQL_DATEADD As String = " DATEADD "
    Public Const SQL_CONVERT As String = " CONVERT "
    Public Const SQL_VARCHAR As String = " VARCHAR "

    Public Const _set_ As String = " SET "
    Public Const sql_UPDATE As String = " UPDATE "
    Public Const SQL_INTO As String = " INTO "
    Public Const SQL_INSERT_INTO As String = " INSERT INTO "
    Public Const _VALUES_ As String = " VALUES "
    Public Const SQL_DELETE As String = " DELETE "
    Public Const SQL_DELETE_FROM As String = " DELETE FROM "

    Public Const PUNTOYCOMA As String = ";"
    Public Const CORCHETEI As String = "["
    Public Const CORCHETED As String = "]"
    Public Const PARENTESISI As String = "("
    Public Const PARENTESISD As String = ")"
    Public Const PUNTO As String = "."
    Public Const COMA As String = ","
    Public Const VACIO As String = ""
    Public Const ASTERISCO As String = "*"
    Public Const PORCENTAJE As String = "%"
    Public Const sql_SELECT_a_FROM As String = SQL_SELECT & ASTERISCO & SQL_FROM
    Public Const sql_SELECT_COUNT_from As String = SQL_SELECT & "COUNT(*)" & SQL_FROM

    Public Const SQL_DISTINCT As String = " DISTINCT "
    Public Const SQL_TOP As String = " TOP "
    Public Const SQL_ASC As String = " ASC "
    Public Const SQL_DESC As String = " DESC "



    Public Enum tipofecha
        ddMMyyyy = 103
        yyyyddMMmmss = 120
        yyyyddMMmmssmi = 121
    End Enum

    Enum alineacion
        izq = 1
        der = 2
    End Enum

    Public Const FORMATO_FECHA_DMA As String = "dd/MM/yyyy"
    Public Const FORMATO_HORA_HMS As String = "HH:mm:ss"
    Public Const FORMATO_HORA_HM As String = "HH:mm"
    Public Const ESTILO_FECHA As Integer = 103
    Public Const ESTILO_HORA As Integer = 108
    Public Const Y As String = " and "
    Public Const O As String = " or  "
    Public Const ASCEN As String = " ASC "
    Public Const DESCEN As String = " DESC"

    Public Function CambiarColorControlesScript(colorNuevo1 As String, colorNuevo2 As String) As String
        Dim script1 As String = "
            <script>
                document.addEventListener('DOMContentLoaded', function() {
                    const targetColor1 = 'rgb(37, 33, 73)'; // Primer color objetivo en RGB (equivalente a #252149)
                    const targetColor2 = 'rgb(39, 185, 198)'; // Segundo color objetivo en RGB (equivalente a #27b9c6)
                    const newColor1 = '" & colorNuevo1 & "'; // Primer color nuevo en hexadecimal
                    const newColor2 = '" & colorNuevo2 & "'; // Segundo color nuevo en hexadecimal
                    
                    function replaceColors(element) {
                        const computedStyle = getComputedStyle(element);
                        
                        // Verificar y reemplazar colores para el primer color objetivo
                        if (computedStyle.backgroundColor === targetColor1) {
                            element.style.setProperty('background-color', newColor1, 'important');
                        }
                        if (computedStyle.color === targetColor1) {
                            element.style.setProperty('color', newColor1, 'important');
                        }
                        if (computedStyle.borderColor === targetColor1) {
                            element.style.setProperty('border-color', newColor1, 'important');
                        }

                        // Verificar y reemplazar colores para el segundo color objetivo
                        if (computedStyle.backgroundColor === targetColor2) {
                            element.style.setProperty('background-color', newColor2, 'important');
                        }
                        if (computedStyle.color === targetColor2) {
                            element.style.setProperty('color', newColor2, 'important');
                        }
                        if (computedStyle.borderColor === targetColor2) {
                            element.style.setProperty('border-color', newColor2, 'important');
                        }
                    }

                    // Recorrer todos los elementos de la página y aplicar reemplazo
                    document.querySelectorAll('*').forEach(replaceColors);
                    document.body.style.visibility = 'visible'; // Mostrar la página después de los cambios
                });
            </script>"

        Return script1

    End Function
    Public Sub CambiarColorControles(ctrl As Control, colorprimario As String, colorsecundario As String)

        ' Recorrer todos los controles de la página
        For Each c As Control In ctrl.Controls
            If TypeOf c Is TextBox Then
                ' Asignar el color de fondo y de texto con formato hexadecimal
                DirectCast(c, TextBox).BackColor = System.Drawing.ColorTranslator.FromHtml(colorprimario)
                DirectCast(c, TextBox).ForeColor = System.Drawing.ColorTranslator.FromHtml("#000000") ' Color negro
            ElseIf TypeOf c Is Button Then
                ' Asignar el color de fondo y de texto con formato hexadecimal
                DirectCast(c, Button).BackColor = System.Drawing.ColorTranslator.FromHtml(colorprimario) ' Color verde
                DirectCast(c, Button).ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF")
            ElseIf TypeOf c Is LinkButton Then
                DirectCast(c, LinkButton).BackColor = System.Drawing.ColorTranslator.FromHtml(colorsecundario) ' Color verde
                DirectCast(c, LinkButton).ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF")

            End If

            ' Si el control tiene hijos, llamamos recursivamente
            If c.HasControls() Then
                CambiarColorControles(c, colorprimario, colorsecundario)
            End If
        Next
    End Sub



    Public Function SumaColumnasExcel(ByVal columna As String, ByVal valor As Integer) As String
        Dim i As Integer
        Dim temp As String
        temp = columna
        For i = 1 To valor - 1
            temp = SumaColumnasExcel(temp)
        Next
        Return temp
    End Function
    'Public Sub EncryptFile(ByVal sInputFilename As String, _
    '                   ByVal sOutputFilename As String, _
    '                   ByVal sKey As String)

    '    Dim fsInput As New FileStream(sInputFilename, _
    '                                FileMode.Open, FileAccess.Read)
    '    Dim fsEncrypted As New FileStream(sOutputFilename, _
    '                                FileMode.Create, FileAccess.Write)

    '    Dim DES As New DESCryptoServiceProvider()

    '    'Establecer la clave secreta para el algoritmo DES.
    '    'Se necesita una clave de 64 bits y IV para este proveedor
    '    DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey)

    '    'Establecer el vector de inicialización.
    '    DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)

    '    'crear cifrado DES a partir de esta instancia
    '    Dim desencrypt As ICryptoTransform = DES.CreateEncryptor()
    '    'Crear una secuencia de cifrado que transforma la secuencia
    '    'de archivos mediante cifrado DES
    '    Dim cryptostream As New CryptoStream(fsEncrypted, _
    '                                        desencrypt, _
    '                                        CryptoStreamMode.Write)

    '    'Leer el texto del archivo en la matriz de bytes
    '    Dim bytearrayinput(fsInput.Length - 1) As Byte
    '    fsInput.Read(bytearrayinput, 0, bytearrayinput.Length)
    '    'Escribir el archivo cifrado con DES
    '    cryptostream.Write(bytearrayinput, 0, bytearrayinput.Length)
    '    cryptostream.Close()
    'End Sub

    'Public Sub DecryptFile(ByVal sInputFilename As String, _
    '    ByVal sOutputFilename As String, _
    '    ByVal sKey As String)

    '    Dim DES As New DESCryptoServiceProvider()
    '    'Se requiere una clave de 64 bits y IV para este proveedor.
    '    'Establecer la clave secreta para el algoritmo DES.
    '    DES.Key() = ASCIIEncoding.ASCII.GetBytes(sKey)
    '    'Establecer el vector de inicialización.
    '    DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)

    '    'crear la secuencia de archivos para volver a leer el archivo cifrado
    '    Dim fsread As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)
    '    'crear descriptor DES a partir de nuestra instancia de DES
    '    Dim desdecrypt As ICryptoTransform = DES.CreateDecryptor()
    '    'crear conjunto de secuencias de cifrado para leer y realizar 
    '    'una transformación de descifrado DES en los bytes entrantes
    '    Dim cryptostreamDecr As New CryptoStream(fsread, desdecrypt, CryptoStreamMode.Read)
    '    'imprimir el contenido de archivo descifrado
    '    Dim fsDecrypted As New StreamWriter(sOutputFilename)
    '    fsDecrypted.Write(New StreamReader(cryptostreamDecr).ReadToEnd)
    '    fsDecrypted.Flush()
    '    fsDecrypted.Close()
    'End Sub
    Public Function QuitaAcentosHtml(ByVal Cadena As String) As String

        ' Cadena = Cadena.Replace("???", "")
        Cadena = Cadena.Replace("á", "&aacute")
        Cadena = Cadena.Replace("é", "&eacute")
        Cadena = Cadena.Replace("í", "&iacute")
        Cadena = Cadena.Replace("ó", "&oacute")
        Cadena = Cadena.Replace("ú", "&uacute")
        Cadena = Cadena.Replace("Á", "&aacute")
        Cadena = Cadena.Replace("É", "&eacute")
        Cadena = Cadena.Replace("Í", "&iacute")
        Cadena = Cadena.Replace("Ó", "&oacute")
        Cadena = Cadena.Replace("Ú", "&uacute")
        Cadena = Cadena.Replace("Ñ", "&Ntilde;")
        Cadena = Cadena.Replace("ñ", "&ntilde;")
        '  Cadena = Cadena.Replace("??", "TTT")


        Return Cadena
    End Function

    Public Function MD5EncryptPass(ByVal StrPass As String) As String
   
Dim Ue As New UnicodeEncoding()
        Dim ByteSourceText() As Byte = Ue.GetBytes(StrPass)
        Dim Md5 As New MD5CryptoServiceProvider()
        Dim ByteHash() As Byte = Md5.ComputeHash(ByteSourceText)
        Return Convert.ToBase64String(ByteHash)

    End Function
    Public Sub CryptFile(ByVal password As String, ByVal _
            in_file As String, ByVal out_file As String, ByVal _
            encrypt As Boolean)
        ' Create input and output file streams.
        Using in_stream As New FileStream(in_file, _
            FileMode.Open, FileAccess.Read)
            Using out_stream As New FileStream(out_file, _
                FileMode.Create, FileAccess.Write)
                ' Encrypt/decrypt the input stream into the
                ' output stream.
                CryptStream(password, in_stream, out_stream, _
                    encrypt)
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Función para eliminar todos los caracteres especiales, 
    ''' excepto ("@", "-" , "."y ",").
    ''' </summary>
    ''' <param name="campo">Nombre del campo que queremos evaluar.</param>
    ''' <returns>Devuelve el texto sin caracteres especiales.</returns>
    Public Function EliminarCaracteresEspeciales(campo As String) As String
        ' Si es un caracter invalido lo reemplazamos
        ' con un espacio en blanco.
        Try
            'campo = (Regex.Replace(campo, "[^\w\.,@-]", ""))
            'campo = campo.Replace("<", "")
            'campo = campo.Replace(">", "")
            'campo = campo.Replace("/", "")
            'campo = campo.Replace("\", "")
            'campo = campo.Replace("@", "")

            Return campo
            ' Si ocurre un timeout convertimos a string.empty
        Catch e As RegexMatchTimeoutException
            Return String.Empty
        End Try
    End Function


    Public Sub Encripta(ByVal password As String, ByVal _
 in_file As String)
        Dim out_file As String
        out_file = in_file & "_"
        CryptFile(password, in_file, out_file, True)
        File.Copy(out_file, in_file, True)
        File.Delete(out_file)
    End Sub
    Public Sub Desencripta_Pon_mismo_nombre(ByVal password As String, ByVal _
        in_file As String)
        Dim out_file As String
        out_file = in_file & "_"
        CryptFile(password, in_file, out_file, False)
        File.Copy(out_file, in_file, True)
        File.Delete(out_file)
    End Sub
    Public Sub Desencripta_Pon_Distinto_nombre(ByVal password As String, ByVal _
        in_file As String)
        Dim out_file As String
        out_file = in_file & "_"
        'CryptFile(password, in_file, out_file, False)
        CryptFile(password, in_file, out_file, False)
        'File.Copy(out_file, in_file, True)

    End Sub
    Public Function Sella(ByVal RutaSellador As String, ByVal RutaSello As String, ByVal RutaFicheroaSellar As String) As Boolean
        Try
            'Dim proceso As System.Diagnostics.Process
            'proceso = System.Diagnostics.Process.Start(RutaSellador, RutaFicheroaSellar & " stamp " & RutaSello & " output " & RutaFicheroaSellar & "_")
            'proceso.WaitForExit()
            'proceso.Close()

            Dim proProcess As Diagnostics.Process
            Dim inStartInfo As New Diagnostics.ProcessStartInfo
            inStartInfo.WindowStyle = Diagnostics.ProcessWindowStyle.Hidden
            inStartInfo.FileName = RutaSellador
            inStartInfo.Arguments = RutaFicheroaSellar & " stamp " & RutaSello & " output " & RutaFicheroaSellar & "_"
            proProcess = Diagnostics.Process.Start(inStartInfo)
            proProcess.WaitForExit()
            proProcess = Nothing
            inStartInfo = Nothing



            IO.File.Delete(RutaFicheroaSellar)
            IO.File.Move(RutaFicheroaSellar & "_", RutaFicheroaSellar)
            Return True
        Catch ex As Exception
            MsgBox("Error:" & ex.Message)
            Return False
        End Try

    End Function
    Public Sub BorraDirectoriotemporal_usuario(Optional ByVal d As String = "")
        Dim ficheros() As String
        Dim directorios() As String
        Dim f As String
        Dim df As String
        If d = "" Then
            d = DIRECTORIOTEMPORAL
        End If
        If System.IO.Directory.Exists(d) Then
            Try
                ficheros = System.IO.Directory.GetFiles(d)
                For Each f In ficheros
                    System.IO.File.Delete(f)
                Next
                directorios = System.IO.Directory.GetDirectories(d)
                For Each df In directorios
                    Util.BorraDirectoriotemporal(df)
                    Directory.Delete(df)
                Next

                'Directory.Delete(d)
            Catch e As Exception
                'hay algun fichero abierto, no podemos borrarlo
                'continuamos
            End Try
        End If
    End Sub
    Public Function CreaRutaNivel9(ByVal RutaIncial As String, ByVal IdCuadro As Long) As String

        Dim Carpetadestino9 As String
        Dim RutaCompuesta9 As String
        Dim i, j, k As Integer
        k = IdCuadro
        j = CLng(Math.Ceiling(k / 1000))

        If j = 0 Then j = 1

        Carpetadestino9 = j.ToString("000000000")
        RutaCompuesta9 = RutaIncial
        For i = 1 To 3
            RutaCompuesta9 = RutaCompuesta9 + Mid(Carpetadestino9, (3 * i) - 2, 3) + "\"
            If Not Directory.Exists(RutaCompuesta9) Then
                Directory.CreateDirectory(RutaCompuesta9)
            End If
        Next i

        CreaRutaNivel9 = RutaCompuesta9

    End Function

    Public Function SacaRutaNivel9(ByVal RutaIncial As String, ByVal IdCuadro As Long) As String

        Dim Carpetadestino9 As String
        Dim RutaCompuesta9 As String
        Dim i, j, k As Integer
        k = IdCuadro
        j = CLng(Math.Ceiling(k / 1000))

        If j = 0 Then j = 1

        Carpetadestino9 = j.ToString("000000000")
        RutaCompuesta9 = RutaIncial
        For i = 1 To 3
            RutaCompuesta9 = RutaCompuesta9 + Mid(Carpetadestino9, (3 * i) - 2, 3) + "\"
        Next i

        sacaRutaNivel9 = RutaCompuesta9

    End Function
    Public Sub CryptStream(ByVal password As String, ByVal _
    in_stream As Stream, ByVal out_stream As Stream, ByVal _
    encrypt As Boolean)
        ' Make an AES service provider.
        Dim aes_provider As New DESCryptoServiceProvider()

        ' Find a valid key size for this provider.
        Dim key_size_bits As Integer = 0
        For i As Integer = 1024 To 1 Step -1
            If (aes_provider.ValidKeySize(i)) Then
                key_size_bits = i
                Exit For
            End If
        Next i
        System.Diagnostics.Debug.Assert(key_size_bits > 0)
        Console.WriteLine("Key size: " & key_size_bits)

        ' Get the block size for this provider.
        Dim block_size_bits As Integer = aes_provider.BlockSize

        ' Generate the key and initialization vector.
        Dim key() As Byte = Nothing
        Dim iv() As Byte = Nothing
        Dim salt() As Byte = {&H0, &H0, &H1, &H2, &H3, &H4, _
            &H5, &H6, &HF1, &HF0, &HEE, &H21, &H22, &H45}
        MakeKeyAndIV(password, salt, key_size_bits, _
            block_size_bits, key, iv)

        ' Make the encryptor or decryptor.
        Dim crypto_transform As ICryptoTransform
        If (encrypt) Then
            crypto_transform = _
                aes_provider.CreateEncryptor(key, iv)
        Else
            crypto_transform = _
                aes_provider.CreateDecryptor(key, iv)
        End If

        ' Attach a crypto stream to the output stream.
        ' Closing crypto_stream sometimes throws an
        ' exception if the decryption didn't work
        ' (e.g. if we use the wrong password).
        Try
            Using crypto_stream As New CryptoStream(out_stream, _
                crypto_transform, CryptoStreamMode.Write)
                ' Encrypt or decrypt the file.
                Const block_size As Integer = 1024
                Dim buffer(block_size) As Byte
                Dim bytes_read As Integer
                Do
                    ' Read some bytes.
                    bytes_read = in_stream.Read(buffer, 0, _
                        block_size)
                    If (bytes_read = 0) Then Exit Do

                    ' Write the bytes into the CryptoStream.
                    crypto_stream.Write(buffer, 0, bytes_read)
                Loop
            End Using
        Catch ex As Exception
            iv = iv
        End Try

        crypto_transform.Dispose()
    End Sub
    Private Sub MakeKeyAndIV(ByVal password As String, ByVal _
    salt() As Byte, ByVal key_size_bits As Integer, ByVal _
    block_size_bits As Integer, ByRef key() As Byte, ByRef _
    iv() As Byte)
        Dim derive_bytes As New Rfc2898DeriveBytes(password, _
            salt, 1000)

        key = derive_bytes.GetBytes(key_size_bits / 8)
        iv = derive_bytes.GetBytes(block_size_bits / 8)
    End Sub
    Public Function SumaColumnasExcel(ByVal columna As String) As String
        'Columa viene de la forma "A" "B" "C" .... "Z" "AA" "AB" .... "BA" "BB"   "ZZ" 
        Dim i, l As Integer
        Dim c As Char
        Dim resultado As String = String.empty

        l = columna.Length - 1

        For i = l To 0 Step -1
            c = columna.Chars(i)

            If c = "Z" Then
                If i = 0 Then
                    resultado = "AA" & resultado
                    Exit For
                Else
                    resultado = "A" & resultado
                End If
            Else
                resultado = columna.Substring(0, i) & Chr(Asc(c) + 1) & resultado
                Exit For
            End If
        Next

        Return resultado
    End Function

    Public Sub CreaDirectorioTemporalAplicacion(Optional ByVal nombrefijo As String = "", Optional ByVal prefijo As String = "")
        Dim dirtemporal As String

        If nombrefijo = "" Then
            Do
                dirtemporal = System.IO.Path.GetTempPath & prefijo & Alea.CadenaAleatoria(5)
            Loop Until Not System.IO.Directory.Exists(dirtemporal)

            System.IO.Directory.CreateDirectory(dirtemporal)
        Else
            'Queremos el mismo nombre para cada ejecucion
            dirtemporal = System.IO.Path.GetTempPath & nombrefijo
            If Not System.IO.Directory.Exists(dirtemporal) Then
                System.IO.Directory.CreateDirectory(dirtemporal)
            End If
        End If

        DIRECTORIOTEMPORAL = dirtemporal & "\"

    End Sub
    Public Sub BorraDirectoriotemporal(Optional ByVal d As String = "")
        Dim ficheros() As String
        Dim f As String
        If d = "" Then
            d = DIRECTORIOTEMPORAL
        End If
        If System.IO.Directory.Exists(d) Then
            Try
                ficheros = System.IO.Directory.GetFiles(d)
                For Each f In ficheros
                    System.IO.File.Delete(f)
                Next
                'Directory.Delete(d)
            Catch e As Exception
                'hay algun fichero abierto, no podemos borrarlo
                'continuamos
            End Try
        End If
    End Sub
    Public Function CreaFicherotemporal(Optional ByVal prefijo As String = "", Optional ByVal extension As String = ".tmp") As String

        Dim nombrefichero As String
        Dim fs As System.IO.FileStream
        Const LONG_NOMBREFICHERO As Integer = 30
        Const MAX_LONG_NOMBREFICHERO As Integer = 25
        Dim s As String

        s = Left(prefijo, MAX_LONG_NOMBREFICHERO) & "~"
        Do
            nombrefichero = DIRECTORIOTEMPORAL & s & Alea.CadenaAleatoria(LONG_NOMBREFICHERO - s.Length) & extension
        Loop Until Not System.IO.File.Exists(nombrefichero)

        fs = System.IO.File.Create(nombrefichero)
        fs.Close()
        Return nombrefichero

    End Function
    Public Function ConcatenaDir(ByVal base As String, ByVal directorio As String) As String
        Dim t As String

        t = directorio.Trim
        If base.EndsWith("\") Then
            If t.StartsWith("\") Then
                t = t.Substring(1)
            End If
        Else
            If Not t.StartsWith("\") Then
                t = "\" + t
            End If
        End If

        If Not t.EndsWith("\") Then
            t = t + "\"
        End If
        Return base + t


    End Function
    Public Function MismoFichero(ByVal forigen As String, ByVal fdestino As String) As Boolean
        Dim iorigen, idestino As System.IO.FileInfo
        iorigen = New System.IO.FileInfo(forigen)
        idestino = New System.IO.FileInfo(fdestino)

        With iorigen
            If .Exists Then
                If idestino.Exists Then
                    Return .LastWriteTime = idestino.LastWriteTime AndAlso _
                           .Length = idestino.Length
                Else
                    'si el destino no existe, no es el mismo fichero
                    Return False
                End If
            Else
                'Si el origen no existe, no tenemos nada que comparar
                Throw New Exception("No se encuentra el fichero: " & forigen)
            End If
        End With
    End Function

    Public Sub BorraficheroSiExiste(ByVal n As String)
        If System.IO.File.Exists(n) Then System.IO.File.Delete(n)

    End Sub

    Public Function BuscaParametroEnLineaComandos(ByVal base As String, ByVal busqueda As String, ByRef valor As String) As Boolean
        Const guion As String = "-"
        Const espacio As String = " "
        Dim s As String
        Dim ini, fin As Integer
        s = espacio & guion & busqueda & espacio
        ini = InStr(base, s)
        If ini <= 0 Then
            s = guion & busqueda & espacio
            ini = InStr(base, s)
            If ini <> 1 Then
                Return False
            End If
        End If
        ini = ini + Len(s) - 1

        If ini > Len(base) Then
            'es el ultimo
            valor = ""
            Return True
        End If
        fin = InStr(ini, base, espacio & guion)
        If fin <= 0 Then
            fin = Len(base)
        End If
        valor = Trim(Mid(base, ini, fin - ini))
        Return True
    End Function

    Public Function LongaBytes(ByVal dato As Long) As Byte()
        Dim b(3) As Byte

        b(0) = CByte(dato And 255)
        b(1) = CByte((dato \ 256) And 255)
        b(2) = CByte((dato \ 65536) And 255)
        b(3) = CByte((dato \ 16777216) And 255)
        Return b
    End Function

    Public Function BytesaLong(ByVal b As Byte()) As Long
        Return CLng(b(0)) + 256 * CLng(b(1)) + 65536 * CLng(b(2)) + 16777216 * CLng(b(3))
    End Function

    Public Function IntaBytes(ByVal dato As Long) As Byte()
        Dim b(1) As Byte

        b(0) = CByte(dato And 255)
        b(1) = CByte((dato \ 256) And 255)
        Return b
    End Function

    Public Function BytesaInt(ByVal b As Byte()) As Integer
        Return CInt(b(0)) + 256 * CInt(b(1))
    End Function

    Public Function CadenaaBytes(ByVal cad As String) As Byte()
        '        Dim i As Integer
        '       Dim result As Byte()

        '      ReDim result(cad.Length - 1)

        '     For i = 0 To cad.Length - 1
        '    result(i) = CByte(Asc(cad.Chars(i)))
        '   Next
        '  Return result
        Dim encodingASCII As New System.Text.ASCIIEncoding
        Return encodingASCII.GetBytes(cad)

    End Function
    Public Function BytesaCadena(ByVal b As Byte()) As String
        Dim i As Integer
        Dim temp As String = String.empty

        For i = 1 To b.Length
            temp &= Chr(b(i))
        Next i
        Return temp
    End Function

    Public Function Existe(ByVal c As Collection, ByVal k As String) As Boolean

        Try
            If (c(k)) Is Nothing Then Return True
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub VaciaColleccion(ByVal c As Collection)
        While c.Count > 0
            c.Remove(1)
        End While
    End Sub

    Public Function Directorioaplicacion() As String
        Dim tmp As String = String.empty
        Dim i As Integer
#If COMPACT_FRAMEWORK Then
        tmp = System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase
#ElseIf APLICACION_DE_CONSOLA Then
        tmp = System.Reflection.Assembly.GetExecutingAssembly.Location
        'esto devuelve la ruta y el ejecutable       
#Else
        'tmp = System.Windows.Forms.Application.ExecutablePath
#End If
        i = tmp.LastIndexOf("\")
        If i <> -1 Then
            Return tmp.Substring(0, i + 1)
        Else
            Return "\"
        End If
    End Function
    Public Function FinaldeMes(ByVal m As Integer, ByVal y As Integer) As Integer
        Dim meses() As Integer = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}

        If m = 2 Then
            If (y Mod 4) = 0 Then
                'divisible por 4
                If (y Mod 100) = 0 Then
                    'divisible por 100, 
                    If (y Mod 400) = 0 Then
                        Return 29
                    Else
                        Return 28
                    End If
                Else
                    Return 29
                End If
            Else
                Return 28
            End If
        Else
            Return meses(m - 1)
            'par = ((m Mod 2) = 0)
            'If m < 8 Then
            '    If par Then
            '        Return 30
            '    Else
            '        Return 31
            '    End If
            'Else
            '    If par Then
            '        Return 31
            '    Else
            '        Return 30
            '    End If
            'End If
        End If
    End Function
    Public Function valorHora(ByVal valor_str As String, ByRef d As DateTime) As Boolean
        Dim hh, mm As Integer
        If valor_str Like "[0-9][0-9]:[0-9][0-9]" Then
            hh = CInt(valor_str.Substring(0, 2))
            mm = CInt(valor_str.Substring(3, 2))
        ElseIf valor_str Like "[0-9][0-9][0-9][0-9]" Then
            hh = CInt(valor_str.Substring(0, 2))
            mm = CInt(valor_str.Substring(2, 2))
        End If

        If hh < 24 And mm < 61 Then
            d = New DateTime(2000, 1, 1, hh, mm, 0, 0)
            Return True
        Else
            d = New DateTime(2000, 1, 1, 0, 0, 0, 0)
            Return False
        End If
    End Function
    Public Function ValorFecha(ByVal valor_str As String, ByRef d As Date) As Boolean
        Dim dd, mm, aaaa As Integer

        If valor_str Like "[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]" Then
            dd = CInt(valor_str.Substring(0, 2))
            mm = CInt(valor_str.Substring(3, 2))
            aaaa = CInt(valor_str.Substring(6, 4))

        ElseIf valor_str Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" Then
            dd = CInt(valor_str.Substring(0, 2))
            mm = CInt(valor_str.Substring(2, 2))
            aaaa = CInt(valor_str.Substring(4, 4))

        ElseIf valor_str Like "[0-9][0-9][0-9][0-9][0-9][0-9]" Then
            dd = CInt(valor_str.Substring(0, 2))
            mm = CInt(valor_str.Substring(2, 2))
            aaaa = 2000 + CInt(valor_str.Substring(4, 2))
        Else
            Return False
        End If

        If mm < 13 Then
            If dd <= Util.FinaldeMes(mm, aaaa) Then
                d = DateSerial(aaaa, mm, dd)
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function
    Public Sub AñadeSinoVacio(ByRef base As String, ByRef prefijo As String, ByRef cadena As String)
        If cadena <> "" Then
            If base = "" Then
                base = cadena
            Else
                base &= prefijo & cadena
            End If
        End If
    End Sub
    Public Sub AñadeCampo(ByRef campos As String, ByVal c As String, ByVal sep As String, ByVal tam As Integer, ByVal ali As alineacion)
        Dim n As String = String.empty
        Select Case sep
            Case " ", vbCrLf
                Dim l As Integer
                l = tam - Len(c)
                If l < 0 Then
                    l = 0
                End If
                n = Space(l)
            Case ""
                n = ""
        End Select
        If sep = vbCrLf Then
            campos &= sep
        End If
        If ali = alineacion.der Then
            campos &= n & c
        Else
            campos &= c & n
        End If
    End Sub

    Public Function HorasAsingle(ByVal h As Long, ByVal m As Long) As Single
        Return Math.Round(CSng(h) + CSng(m) / 60.0, 2)
    End Function

    Public Sub SingleAhoras(ByVal v As Single, ByRef h As Long, ByRef m As Long)
        h = Fix(v)
        m = 60 * (v - h)
    End Sub

    Public Function Valor(ByVal s As String) As Single
        Dim separadordecimal As String
        Dim separadormiles As String

        separadordecimal = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator()
        separadormiles = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator()

        s = s.Trim
        If s = "" Then
            Return ZERO_SNG
        Else
            If InStr(s, COMA) > 0 And InStr(s, PUNTO) > 0 Then
                'tenemos los dos. Lo dejamos igual
                Return CSng(s)
            Else
                Return CSng(Replace(Replace(s, COMA, separadordecimal), PUNTO, separadordecimal))
            End If
        End If
    End Function
    Public Function Valor_entero(ByVal s As String) As Integer
        s = s.Trim
        If s = "" Then
            Return ZERO
        Else
            Return CInt(Replace(s, ",", "."))
        End If
    End Function

    Public Function TextoOK(ByVal s As String) As Boolean
        'Devuelve Verdadero, cuando el texto es CADENA VACIA (o espacios)
        ' o cuando es Numerico
        If Not TextoVacio(s) Then
            If IsNumeric(s) Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function

    Public Function TextoVacio(ByRef s As String) As Boolean
        Return s.Trim = VACIO
    End Function

    Public Function TextOK_noVacio(ByRef s As String) As Boolean
        Return (Not TextoVacio(s)) AndAlso TextoOK(s)
    End Function

    Public Function Trim_bajonivel(ByVal s As String) As String
        Dim result As String = VACIO
        Dim cs(), c As Char
        Dim i, l As Integer
        Dim temp As String = VACIO

        cs = s.ToCharArray
        l = cs.Length - 1

        For i = 0 To l
            c = cs(i)
            If Asc(c) <= 32 Then
                temp &= Chr(32)
            Else
                result &= temp & c
                temp = VACIO
            End If
        Next
        Return result
    End Function

    Public Function Linea(ByRef s As String) As String
        Return s & vbCrLf
    End Function
    Public Function PasaraHex(ByRef s As String) As String
        Dim i As Integer
        Dim result As String
        Dim a As Integer

        result = ""
        For i = 1 To Len(s)
            a = Asc(Mid(s, i, 1))
            If a < 16 Then 'un digito
                result = result & "0"
            End If
            result = result & Hex(a)
        Next i
        PasaraHex = result
    End Function

    Public Function PasaraCadena(ByRef s As String) As String
        Dim i As Integer
        Dim result As String

        result = ""

        For i = 1 To Len(s) - 1 Step 2
            result = result & Chr(Hex2Dec(Mid(s, i, 1), Mid(s, i + 1, 1)))
        Next i
        PasaraCadena = result
    End Function

    Public Function Hex2Dec(ByVal v1 As String, ByVal v2 As String) As Byte
        Hex2Dec = CByte(Hex2Dec_unitario(v1) * 16) + Hex2Dec_unitario(v2)
    End Function

    Public Function Hex2Dec_unitario(ByVal s As String) As Byte
        If s Like "[0-9]" Then
            Hex2Dec_unitario = CByte(s)
        ElseIf s Like "[A-F]" Then
            Hex2Dec_unitario = CByte(10 + (Asc(s) - Asc("A")))
        Else
            Hex2Dec_unitario = 0
        End If
    End Function

    Public Function HoraSQL(ByVal d As Date) As String
        Return DatetimeSql(d, FORMATO_HORA_HMS, ESTILO_HORA)
    End Function

    Private Function DatetimeSql(ByVal d As Date, ByVal formato As String, ByVal estilo As Integer) As String
        Return "CONVERT(DATETIME," & v(Format(d, formato)) & "," & CStr(estilo) & ")"
    End Function

    Public Function NormalizaRuta(ByVal s As String) As String
        If Not s.EndsWith("\") Then
            Return s & "\"
        Else
            Return s
        End If
    End Function
    Public Function Formatea_nombre_fichero_carpeta(ByVal nombre As String) As String
        nombre = Replace(nombre, "/", "_")
        nombre = Replace(nombre, "\", "_")
        nombre = Replace(nombre, ":", "_")
        nombre = Replace(nombre, "*", "_")
        nombre = Replace(nombre, "?", "_")
        nombre = Replace(nombre, """", "_")
        nombre = Replace(nombre, "<", "_")
        nombre = Replace(nombre, ">", "_")
        nombre = Replace(nombre, "|", "_")
        Return nombre
    End Function
    Public Function Obtenbytes(ByVal n As String) As Byte()
        Dim fsImageFile As System.IO.FileStream
        Dim bytimagedata() As Byte


        fsImageFile = New System.IO.FileStream(n, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Try
            ReDim bytimagedata(CInt(fsImageFile.Length - 1))

            fsImageFile.Read(bytimagedata, 0, bytimagedata.Length)
            fsImageFile.Close()

        Finally
            fsImageFile.Close()
        End Try


        Return bytimagedata
    End Function

    Public Function NumeroALetras(ByVal numero As String, ByVal comoensap As Boolean) As String
        '********Declara variables de tipo cadena************
        Dim entero As String = String.empty
        Dim dec As String = String.empty
        Dim palabrasdecimales, flag As String
        Dim palabras As String = String.empty
        Dim e, c As String

        '********Declara variables de tipo entero***********
        Dim x, y As Integer

        flag = "N"

        '**********Número Negativo***********
        If Mid(numero, 1, 1) = "-" Then
            numero = Mid(numero, 2, numero.ToString.Length - 1).ToString
            palabras = "menos "
        End If

        '**********Si tiene ceros a la izquierda*************
        For x = 1 To numero.ToString.Length
            If Mid(numero, 1, 1) = "0" Then
                numero = Trim(Mid(numero, 2, numero.ToString.Length).ToString)
                If numero.ToString.Length = 0 Then palabras = ""
            Else
                Exit For
            End If
        Next

        '*********Dividir parte entera y decimal************
        For y = 1 To Len(numero)
            If Mid(numero, y, 1) = "." Then
                flag = "S"
            Else
                If flag = "N" Then
                    entero = entero + Mid(numero, y, 1)
                Else
                    dec = dec + Mid(numero, y, 1)
                End If
            End If
        Next y

        If Len(dec) = 1 Then dec = dec & "0"

        '**********proceso de conversión***********
        flag = "N"

        If Val(entero) = 1 Then
            e = "Euro"
        Else
            If comoensap Then
                e = "Euros"
            Else
                If entero Like "[0-9]000000" Then
                    e = "de Euros"
                Else
                    e = "Euros"
                End If
            End If

            ' habría que detectar el caso de los millones "redondos" 
            ' en español se dice " de Euros"

        End If

        If Val(dec) = 1 Then
            c = "Céntimo"
        Else
            c = "Céntimos"
        End If

        If Val(numero) <= 999999999 Then
            palabras &= conversion(entero, False)
            If palabras <> "" Then
                'poner en mayúsculas la primera letra
                palabras = palabras.Substring(0, 1).ToUpper & palabras.Substring(1)
            End If

            '**********Une la parte entera y la parte decimal*************
            If dec <> "00" Then
                palabrasdecimales = conversion(dec, True)
                If palabras <> "" Then
                    Return palabras & e & " y " & palabrasdecimales & c
                Else
                    Return palabrasdecimales.Substring(0, 1).ToUpper & palabrasdecimales.Substring(1) & c
                End If
            Else
                Return palabras & e
            End If

        Else
            Return ""
        End If
    End Function


    Private Function conversion(ByVal entero As String, ByVal estamosendecimales As Boolean) As String
        Dim flag As String
        Dim y As Integer
        Dim num As Integer
        Dim palabras As String = String.empty
        Dim longitud As Integer

        flag = "N"


        longitud = Len(entero)
        For y = longitud To 1 Step -1
            num = longitud - (y - 1)
            Select Case y
                Case 3, 6, 9
                    '**********Asigna las palabras para las centenas***********
                    Select Case Mid(entero, num, 1)
                        Case "1"
                            If Mid(entero, num + 1, 1) = "0" And Mid(entero, num + 2, 1) = "0" Then
                                palabras &= "cien "
                            Else
                                palabras &= "ciento "
                            End If
                        Case "2"
                            palabras &= "doscientos "
                        Case "3"
                            palabras &= "trescientos "
                        Case "4"
                            palabras &= "cuatrocientos "
                        Case "5"
                            palabras &= "quinientos "
                        Case "6"
                            palabras &= "seiscientos "
                        Case "7"
                            palabras &= "setecientos "
                        Case "8"
                            palabras &= "ochocientos "
                        Case "9"
                            palabras &= "novecientos "
                    End Select
                Case 2, 5, 8
                    '*********Asigna las palabras para las decenas************
                    Select Case Mid(entero, num, 1)
                        Case "1"
                            If Mid(entero, num + 1, 1) = "0" Then
                                flag = "S"
                                palabras &= "diez "
                            End If
                            If Mid(entero, num + 1, 1) = "1" Then
                                flag = "S"
                                palabras &= "once "
                            End If
                            If Mid(entero, num + 1, 1) = "2" Then
                                flag = "S"
                                palabras &= "doce "
                            End If
                            If Mid(entero, num + 1, 1) = "3" Then
                                flag = "S"
                                palabras &= "trece "
                            End If
                            If Mid(entero, num + 1, 1) = "4" Then
                                flag = "S"
                                palabras &= "catorce "
                            End If
                            If Mid(entero, num + 1, 1) = "5" Then
                                flag = "S"
                                palabras &= "quince "
                            End If
                            If Mid(entero, num + 1, 1) > "5" Then
                                flag = "N"
                                palabras &= "dieci"
                            End If
                        Case "2"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras &= "veinte "
                                flag = "S"
                            Else
                                palabras &= "veinti"
                                flag = "N"
                            End If
                        Case "3"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras &= "treinta "
                                flag = "S"
                            Else
                                palabras &= "treinta y "
                                flag = "N"
                            End If
                        Case "4"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras &= "cuarenta "
                                flag = "S"
                            Else
                                palabras &= "cuarenta y "
                                flag = "N"
                            End If
                        Case "5"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras &= "cincuenta "
                                flag = "S"
                            Else
                                palabras &= "cincuenta y "
                                flag = "N"
                            End If
                        Case "6"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras &= "sesenta "
                                flag = "S"
                            Else
                                palabras &= "sesenta y "
                                flag = "N"
                            End If
                        Case "7"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras &= "setenta "
                                flag = "S"
                            Else
                                palabras &= "setenta y "
                                flag = "N"
                            End If
                        Case "8"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras &= "ochenta "
                                flag = "S"
                            Else
                                palabras &= "ochenta y "
                                flag = "N"
                            End If
                        Case "9"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras &= "noventa "
                                flag = "S"
                            Else
                                palabras &= "noventa y "
                                flag = "N"
                            End If
                    End Select
                Case 1, 4, 7
                    '*********Asigna las palabras para las unidades*********
                    Select Case Mid(entero, num, 1)
                        Case "1"
                            If flag = "N" Then
                                If y = 1 Then
                                    If palabras.EndsWith("veinti") Then
                                        palabras &= "ún "
                                    Else
                                        If longitud = 1 Then
                                            palabras &= "un "
                                        Else
                                            ' If estamosendecimales Then
                                            palabras &= "un "
                                            ' Else
                                            '    palabras &= "uno "
                                            ' End If

                                        End If
                                    End If

                                Else
                                    If y = 4 And longitud = 4 Then

                                    Else
                                        If palabras.EndsWith("veinti") Then
                                            palabras &= "ún "
                                        Else
                                            palabras &= "un "
                                        End If
                                    End If
                                End If
                            Else
                                If y = 4 Then
                                    If longitud > 5 AndAlso Mid(entero, num - 2, 2) = "00" Then
                                        'nada de nada 
                                    ElseIf longitud = 4 Then
                                        ' nada de nada 
                                    Else
                                        'If longitud > 4 And Mid(entero, num - 1, 1) = "0" Then
                                        '    palabras &= "un "

                                        'End If
                                        palabras &= "un "
                                    End If
                                End If
                            End If
                        Case "2"
                            If flag = "N" Then
                                If palabras.EndsWith("veinti") Then
                                    palabras &= "dós "
                                Else
                                    palabras &= "dos "
                                End If
                            End If

                        Case "3"
                            If flag = "N" Then
                                If palabras.EndsWith("veinti") Then
                                    palabras &= "trés "
                                Else
                                    palabras &= "tres "
                                End If

                            End If

                        Case "4"
                            If flag = "N" Then palabras &= "cuatro "
                        Case "5"
                            If flag = "N" Then palabras &= "cinco "
                        Case "6"
                            If flag = "N" Then palabras &= "seis "
                        Case "7"
                            If flag = "N" Then palabras &= "siete "
                        Case "8"
                            If flag = "N" Then palabras &= "ocho "
                        Case "9"
                            If flag = "N" Then palabras &= "nueve "
                    End Select
            End Select

            '***********Asigna la palabra mil***************
            If y = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                Len(entero) <= 6) Then palabras &= "mil "
            End If

            '**********Asigna la palabra millón*************
            If y = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    palabras &= "millón "
                Else
                    palabras &= "millones "
                End If
            End If
        Next y

        Return palabras
    End Function

    Public Function QuitaCaracteresControl(ByVal s As String) As String
        Dim result As String = VACIO
        Dim ultimo As String = VACIO


        s = s.Trim
        For Each c As Char In s.ToCharArray
            If Char.IsControl(c) Then
                If ultimo <> " " Then
                    result &= " "
                    ultimo = " "
                End If
            Else
                result &= c
                ultimo = c
            End If
        Next
        Return result
    End Function

    Public Function SQL_NormalizaExpresion(ByVal s As String) As String
        Return " " & QuitaCaracteresControl(s) & " "
        'Los espacios son para que las palabras claves siempre tengan un espacio al final
    End Function

    Public Function SQL_SustCadenas(ByVal s As String) As String
        'Esto sirve para evitar cosas del tipo

        ' Select 'SELECT * FROM' & [FROM] FROM [Tabla WHERE]
        'Pasamos la expresion a mayusculas y despues pasamos a minuscula lo que esté
        'entre comillas, corchetes o parentesis.
        'La expresion resultante NO ES EQUIVALENTE a la origen (las constantes se pasan a minuscula)
        'pero sirve para analizar la aparición de las palabras claves (SELECT FROM WHERE...) y que no
        'nos confundamos con posibles constantes.

        'Se ha añadido que también pasamos a minusculas lo que está entre parentesis
        'por ejemplo, consultas anidadas.

        'SELECT c.campo1, (select max(c.campo2) from tabla2) as campo2 FROM tabla3

        Return SC(SC(SC(s.ToUpper, "'"c, "'"c, True, True), "["c, "]"c, True, False), "("c, ")"c, False, False)
    End Function

    Private Function SC(ByVal s As String, ByVal car_abre As String, ByVal car_cierra As String, ByVal doblecaracter_es_literal As Boolean, ByVal sustuirinterior As Boolean) As String
        Dim result As String = VACIO
        Dim i, total, abiertos As Integer
        Dim parcial As String = VACIO
        Dim c As String

        total = Len(s) - 1
        parcial = VACIO

        i = 0
        abiertos = 0

        While i <= total
            c = s(i)
            If abiertos > 0 Then 'Estamos dentro
                If c = car_cierra Then
                    If doblecaracter_es_literal Then
                        If i = total Then
                            'no puede haber otro es el de cierre
                            abiertos -= 1
                        Else
                            If s(i + 1) = car_cierra Then
                                'No lo tenemos en cuenta es doble
                                parcial = parcial & c
                                i += 1
                            Else
                                'Es otro caracter
                                abiertos -= 1
                            End If
                        End If
                    Else
                        abiertos -= 1
                    End If

                    If abiertos = 0 Then
                        If sustuirinterior Then
                            parcial = New String(" "c, parcial.Length)
                        End If
                        result &= car_abre & parcial & car_cierra

                        parcial = VACIO
                    Else
                        parcial &= c.ToLower
                    End If
                Else
                    If c = car_abre Then
                        abiertos += 1
                    End If
                    parcial &= c.ToLower
                End If
            Else ' abiertos = 0, estamos fuera
                If c = car_abre Then
                    abiertos += 1
                    result &= parcial
                    parcial = VACIO
                Else
                    parcial &= c
                End If
            End If
            i = i + 1
        End While

        If abiertos > 0 Then
            'chungo, aqui deberia dar error ya que la cadena no se cierra
            'dejamos los caracteres que están fuera de la cadena como estaban en el original
            result = result & car_abre & Microsoft.VisualBasic.Right(s, Len(parcial)).ToLower
        Else
            result = result & parcial
        End If
        Return result

    End Function


#If Not APLICACION_DE_CONSOLA Then

    'Public Function ValorActual(ByVal c As ComboBox) As ValorCombo

    '    Return DirectCast(c.SelectedItem, ValorCombo)

    'End Function
    Public Sub CompruebaValorNumericoTextBox(ByVal t As TextBox)
        '¡OJO! que cambia el textbox como el formato sea incorrecto
        Dim s As String
        s = t.Text.Replace(",", ".").Trim()
        If Not TextoOK(s) Then
            t.Text = ""
        End If
    End Sub

    'Public Sub AsignarImagen(ByVal t As ToolBar, ByVal i_imagen As Integer, ByVal i_boton As Integer)
    '    ' Presuponemos que la barra ya tiene asignada el imagelist
    '    t.Buttons(i_boton).ImageIndex = i_imagen
    'End Sub


    'Public Sub SeleccionaValorCombo(ByVal c As ComboBox, ByVal s As String)
    '    If TypeOf (c.DataSource) Is ValoresCombo Then
    '        For Each o As ValorCombo In c.Items
    '            If o.Valor = s Then
    '                c.SelectedItem = o
    '            End If
    '        Next
    '    Else
    '        For Each o As Object In c.Items
    '            If CStr(o) = s Then
    '                c.SelectedItem = o
    '            End If
    '        Next
    '    End If
    'End Sub


    Public Function MontaValorCombo(ByVal cod As Long, ByVal descripcion As String) As String
        Return cod.ToString & " " & descripcion
    End Function
    'Public Function ObtenCodigo(ByRef combo As ComboBox) As Integer
    '    Dim i As Integer, j As Integer
    '    Dim s As String

    '    i = combo.SelectedIndex
    '    If i <> -1 Then
    '        s = CStr(combo.SelectedItem)
    '        j = InStr(s, " ")
    '        Return CInt(Mid(s, 1, j - 1))
    '    Else
    '        Return 0
    '    End If

    'End Function


    'Public Function ObtenTexto(ByRef combo As ComboBox) As String
    '    Dim i As Integer, j As Integer
    '    Dim s As String

    '    i = combo.SelectedIndex
    '    If i <> -1 Then
    '        s = CStr(combo.SelectedItem)
    '        j = InStr(s, " ")
    '        Return Mid(s, j + 1)
    '    Else
    '        Return ""
    '    End If

    'End Function




#End If
#If Not COMPACT_FRAMEWORK Then
    Public Function CompruebaDerechosEscrituraenDirectorio(ByVal directorio As String) As Boolean
        Dim fichero As Integer
        Dim nombrefichero As String
        Dim cadena As String
        Try
            nombrefichero = directorio + CadenaAleatoria(6)
            fichero = FileSystem.FreeFile()

            cadena = Alea.CadenaAleatoria(256)

            FileSystem.FileOpen(fichero, nombrefichero, OpenMode.Output, OpenAccess.Write, OpenShare.LockReadWrite)
            FileSystem.WriteLine(fichero, cadena)
            FileSystem.FileClose(fichero)

            fichero = FileSystem.FreeFile()
            FileSystem.FileOpen(fichero, nombrefichero, OpenMode.Input, OpenAccess.Read, OpenShare.LockReadWrite)
            cadena = FileSystem.LineInput(fichero)
            FileSystem.FileClose(fichero)

            System.IO.File.Delete(nombrefichero)
        Catch ex As Exception
            Return False
        End Try
        Return True


    End Function
#End If

    Public Function v(ByVal s As String) As String
        Return "'" & Replace(s, "'", "''") & "'"
    End Function
    Public Function v_like(ByVal s As String) As String
        Return "'%" & s & "%'"
    End Function

    Public Function v2(ByVal s As String) As String
        Return COMILLA_DOBLE & Replace(s, COMILLA_DOBLE, COMILLA_DOBLE & COMILLA_DOBLE) & COMILLA_DOBLE
    End Function

    Public Function func(ByVal f As String, ByVal ParamArray param() As String) As String
        Dim s As String
        s = VACIO

        For i As Integer = LBound(param) To UBound(param)
            Util.AñadeSinoVacio(s, COMA, param(i))
        Next
        If s = VACIO Then
            Return f
        Else
            Return f & entreparen(s)
        End If

    End Function
    Public Function entreparen(ByVal s As String) As String
        Return PARENTESISI & s & PARENTESISD
    End Function
    Public Function FechaCrystalReport(ByRef d As Date) As String
        Return "DATE(" & Year(d) & "," & Month(d) & "," & Day(d) & ")"
    End Function
    Public Function FechaHoraCrystalReport(d As DateTime) As String
        Return "DateTime (" & Year(d) & "," & Month(d) & "," & Day(d) & "," & Hour(d) & "," & Minute(d) & "," & Second(d) & ")"
    End Function

    Public Function fechaAccess(ByVal d As DateTime) As String
        Return d.ToString("'Dateserial'(yyyy,MM,dd) + 'TimeSerial'(HH,mm,ss)")
    End Function
    Public Function fechasql(ByVal d As DateTime, Optional ByVal t As tipofecha = tipofecha.ddMMyyyy) As String

        ' COMPARACIONES DE FECHAS!!!!!
        ' hay veces que nos intersan llegar a los segundos y otras al día.

        Select Case t
            Case tipofecha.ddMMyyyy
                Return "CONVERT(datetime," & v(Format(d, "dd/MM/yyyy")) & ",103)"
            Case tipofecha.yyyyddMMmmss
                'yyyy-mm-dd hh:mi:ss  tipo 120
                Return "CONVERT(datetime," & v(Format(d, "yyyy-MM-dd HH:mm:ss")) & ",120)"
            Case tipofecha.yyyyddMMmmssmi
                ' aaaa-mm-dd hh:mi:ss.mmm tipo 121
                ' Return "CONVERT(datetime," & v(Format(d, "yyyy-MM-dd HH:mm:ss")) & ",120)"
                Return "CONVERT(datetime," & v(d.ToString("yyyy-MM-dd HH:mm:ss.fff")) & ",121)"
            Case Else
                Return ""
        End Select


    End Function

    Public Function rellenadato(ByVal valor As String, ByVal longitud As Integer, ByVal caracterrellenar As String, ByVal derecha As Boolean) As String
        If Len(valor) > longitud Then
            valor = Left(valor, longitud)
        Else
            Dim añade As String = StrDup(longitud - Len(valor), caracterrellenar)
            If derecha Then valor = valor & añade Else valor = añade & valor
        End If
        Return valor
    End Function




    Public Function DeCadenaaFecha(ByVal Cadena As String) As Date
        Return DateSerial(Mid(Cadena, 5, 4), Mid(Cadena, 3, 2), Mid(Cadena, 1, 2))
    End Function
    Public Function DeCadenaaFechaAAAMMDD(ByVal Cadena As String) As Date
        Return DateSerial(Mid(Cadena, 1, 4), Mid(Cadena, 5, 2), Mid(Cadena, 7, 2))
    End Function
    Public Function fechasql_numerico(ByVal nombrecampo As String) As String
        Return "CAST(10000 * YEAR(" & nombrecampo & ") + 100 * MONTH(" & nombrecampo & ") + DAY(" & nombrecampo & ") + 0.01 * DATEPART(hh, " & nombrecampo & ") + 0.0001 * DATEPART(mi," & nombrecampo & ") + 0.000001 * DATEPART(ss, " & nombrecampo & ") AS nvarchar)"
    End Function

    Public Function comparafecha(ByVal nombrecampo As String, ByVal f As Date) As String
        Dim sf As String
        sf = v(Format(f, "yyyyMMdd"))

        ' Return "Day(" & nombrecampo & ") =" & Microsoft.VisualBasic.DateAndTime.Day(f).ToString & _Y_ & "Month(" & nombrecampo & ")=" & Month(f).ToString & _Y_ & "Year(" & nombrecampo & ")=" & Year(f).ToString
        Return "(" & nombrecampo & " >= " & sf & _Y_ & nombrecampo & " < " & "dateadd(dd,1," & sf & "))"
    End Function

    Public Function doublesql(ByVal s As Double) As String
        Return Replace(Replace(CStr(s), PUNTO, VACIO), COMA, PUNTO)
    End Function

    Public Function QZI(ByVal s As String) As String 'Quita Zeros Izquierda
        Dim i, fin As Integer
        Dim result As String

        result = ""
        fin = s.Length
        For i = 0 To fin - 1
            If Not s.Substring(i, 1) = "0" Then
                result = s.Substring(i)
                Exit For
            End If
        Next
        Return result

    End Function
End Module

Public Class ValorCombo
    Private v As String
    Private d As String

    Public Sub New(ByVal value As String, ByVal display As String)
        v = value
        d = display
    End Sub

    Public ReadOnly Property Valor() As String
        Get
            Return v
        End Get
    End Property

    Public ReadOnly Property Mostrar() As String
        Get
            Return d
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return v & " - " & d
    End Function
End Class
Public Class parametros_iniciales
    Inherits coleeccion_object

    Public Overloads Sub Add_Long(ByVal nombre As String, ByVal valor As Long)
        MyBase.Add(nombre, valor)
    End Sub

    Public Overloads Sub Add_Date(ByVal nombre As String, ByVal valor As Date)
        MyBase.Add(nombre, valor)
    End Sub

    Public Overloads Sub Add_String(ByVal nombre As String, ByVal valor As String)
        MyBase.Add(nombre, valor)
    End Sub

    Public Overloads Sub Add_Short(ByVal nombre As String, ByVal valor As Long)
        MyBase.Add(nombre, valor)
    End Sub
    Public Overloads Sub Add_Integer(ByVal nombre As String, ByVal valor As Integer)
        MyBase.Add(nombre, valor)
    End Sub
    Public Overloads Sub Add_Bool(ByVal nombre As String, ByVal valor As Boolean)
        MyBase.Add(nombre, valor)
    End Sub

    Public Overloads Sub Add_Single(ByVal nombre As String, ByVal valor As Single)
        MyBase.Add(nombre, valor)
    End Sub

    Public Overloads Sub Add_Char(ByVal nombre As String, ByVal valor As Char)
        MyBase.Add(nombre, valor)
    End Sub

    Public Shadows Function ITem_Long(ByVal nombre As String) As Long
        Return CLng(MyBase.Item(nombre).valor)
    End Function

    Public Shadows Function ITem_Date(ByVal nombre As String) As Date
        Return CDate(MyBase.Item(nombre).valor)
    End Function

    Public Shadows Function ITem_String(ByVal nombre As String) As String
        Return CStr(MyBase.Item(nombre).valor)
    End Function

    Public Shadows Function ITem_Short(ByVal nombre As String) As Short
        Return CShort(MyBase.Item(nombre).valor)
    End Function

    Public Shadows Function ITem_Integer(ByVal nombre As String) As Integer
        Return CInt(MyBase.Item(nombre).valor)
    End Function
    Public Shadows Function ITem_Bool(ByVal nombre As String) As Boolean
        Return CBool(MyBase.Item(nombre).valor)
    End Function

    Public Shadows Function ITem_Single(ByVal nombre As String) As Single
        Return CSng(MyBase.Item(nombre).valor)
    End Function

    Public Shadows Function Item_Char(ByVal nombre As String) As Char
        Return CChar(MyBase.Item(nombre).valor)
    End Function
End Class
#If Not APLICACION_DE_CONSOLA Then
Public Class ValoresCombo
    Inherits ArrayList
    Public Sub New()
        MyBase.new()
    End Sub
    Public Sub New(ByVal t As data.DataTable, Optional ByVal ncolumnacodigo As Integer = 0, Optional ByVal ncolumnadescripcion As Integer = 1)
        MyBase.New()

        For Each r As data.DataRow In t.Rows
            Me.Añade(CStr(r(ncolumnacodigo)), CStr(r(ncolumnadescripcion)))
        Next
    End Sub
    Public Sub New(ByVal dr() As data.DataRow, Optional ByVal ncolumnacodigo As Integer = 0, Optional ByVal ncolumnadescripcion As Integer = 1)
        MyBase.New()
        For Each r As data.DataRow In dr
            Me.Añade(CStr(r(ncolumnacodigo)), CStr(r(ncolumnadescripcion)))
        Next
    End Sub

    Public Sub Añade(ByVal v As String, ByVal d As String)
        Me.Add(New ValorCombo(v, d))
    End Sub


    Public Sub AsignaAListControl(ByVal l As ListControl)
        'With l
        '    If Me.Count > 0 Then

        '        .DataSource = Me
        '        .DisplayMember = "Mostrar"
        '        .ValueMember = "Valor"
        '    Else
        '        .DataSource = Nothing
        '        .DisplayMember = ""
        '        .ValueMember = ""
        '    End If
        'End With
    End Sub
    'Public Sub AsignaAcombo(ByVal c As ComboBox)
    '    Me.AsignaAListControl(c)
    'End Sub


    Public Sub AsignaALista(ByVal l As ListBox)
        Me.AsignaAListControl(l)

    End Sub

    'Public Sub AsignaaCheckedListBox(ByVal l As CheckedListBox)
    '    Me.AsignaAListControl(l)
    'End Sub



End Class
#End If

Public Class informador
    'clase abastracta para poder informar en pantalla de los eventos del protocolo
    Public Sub New()

    End Sub
    Public Overridable Sub Informa(ByVal s As String)

    End Sub
End Class
Public Class MarcadeTiempo
    Private t As Decimal
    Private d As Date
    Private Const año As String = "yyyy"
    Private Const mes As String = "MM"
    Private Const dia As String = "dd"
    Private Const hora As String = "HH"
    Private Const minuto As String = "mm"
    Private Const segundo As String = "ss"
    Private Const milesimas As String = "fff"
    Private Const sep_fecha As String = "/"
    Private Const sep_horas As String = ":"
    Private Const sep_mil As String = "."
    ' Private Const cadenaformato As String = "yyyyMMddHHmmssfff"
    Private Const cadenaformato As String = año & mes & dia & hora & minuto & segundo & milesimas

    Property valor() As Decimal
        Get
            valor = t
        End Get
        Set(ByVal Value As Decimal)
            t = Value
        End Set
    End Property
    ReadOnly Property fecha() As Date
        Get
            Return d
        End Get
    End Property


    Public Sub New()
        d = Now()
        t = CDec(Format(d, cadenaformato))
    End Sub
    Public Sub New(ByVal fecha As Date)
        d = fecha
        t = CDec(Format(d, cadenaformato))
    End Sub
    Public Sub New(ByVal v As Decimal)
        t = v
        d = Now()
    End Sub
    Public Function Legible(ByVal v As Decimal) As String
        Dim s As String
        s = Format(v)
        Return E(s, dia) & sep_fecha & E(s, mes) & sep_fecha & E(s, año) & " " & E(s, hora) & sep_horas & E(s, minuto) & sep_horas & E(s, segundo) & sep_mil & E(s, milesimas)
    End Function

    Public Function Legible() As String
        Return Legible(t)
    End Function

    Private Function E(ByVal cadena As String, ByVal parte As String) As String
        Return Mid(cadena, InStr(cadenaformato, parte), Len(parte))
    End Function


End Class
Public Class FicheroINI
    Private Const CORCHETED As String = "]"
    Private Const CORCHETEI As String = "["
    Private Const PUNTOYCOMA As String = ";"
    Public coleccion As ColecciondeString

    Public Function LeerFicheroINI(ByVal nombrefichero As String) As Boolean
        Dim i As Integer
        Dim c, seccion, campo, valor As String
        Dim ts As System.IO.StreamReader
        Dim result As Boolean


        Try
            If Not System.IO.File.Exists(nombrefichero) Then
                Return False
            End If

            seccion = ""
            coleccion = New ColecciondeString


            ts = System.IO.File.OpenText(nombrefichero)

            c = ts.ReadLine()
            While Not c Is Nothing
                c = Trim(c)
                Select Case Left(c, 1)
                    Case PUNTOYCOMA
                        'Pasamos, es un comentario
                    Case CORCHETEI
                        i = InStr(1, c, CORCHETED)
                        If i > 2 Then
                            seccion = Trim(Mid(c, 2, i - 2)).ToUpper
                        End If
                    Case Else
                        'No es ni un comentario ni una seccion, tiene que ser un valor
                        i = InStr(1, c, IGUAL)
                        If i > 2 Then
                            campo = Trim(Mid(c, 1, i - 1)).ToLower
                            valor = Trim(Mid(c, i + 1))
                            coleccion.Add(seccion & campo, valor)
                        End If
                End Select
                c = ts.ReadLine
            End While
            ts.Close()
            ' f.Close

            result = True
        Catch ex As Exception
            Throw New Exception("Error leyendo fichero: " & nombrefichero & " " & ex.Message)
        End Try

        Return result


    End Function
    Public Function ObtenerValorINIEx(ByVal seccion As String, ByVal nombre As String) As String
        Dim k As String

        k = seccion.ToUpper & nombre.ToLower
        If Me.coleccion.Existe(k) Then
            Return Me.coleccion.Item(k)
        Else
            Throw New Exception("No se ha definido el Valor en fichero .INI " & seccion & " " & nombre)
        End If
    End Function

    Public Function ObtenerValorINI(ByVal seccion As String, ByVal nombre As String) As String
        Dim k As String

        k = seccion.ToUpper & nombre.ToLower
        If coleccion.Existe(k) Then
            Return Me.coleccion.Item(k)
        Else
            Return VACIO
        End If
    End Function

    Public Function ExisteValorINI(ByVal seccion As String, ByVal nombre As String) As Boolean
        Return Me.coleccion.Existe(seccion.ToUpper & nombre.ToLower)
    End Function
End Class

Public Class nodocoleccion
    Public nombre As String
    Public valor As Object

    Public Sub New()

    End Sub
    <System.Diagnostics.DebuggerStepThrough()> Public Sub New(ByVal n As String, ByVal v As Object)
        nombre = n
        valor = v
    End Sub

End Class

Public Class coleeccion_object
    Inherits CollectionBase

    <System.Diagnostics.DebuggerStepThrough()> Protected Sub Add(ByVal nombre As String, ByVal valor As Object)
        Dim n As nodocoleccion

        n = Me.Recupera(nombre)
        If n Is Nothing Then
            List.Add(New nodocoleccion(nombre, valor))
        Else
            n.valor = valor
            n.nombre = nombre
        End If
    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Public Sub Remove(ByVal index As Integer)
        If index < Count And index >= 0 Then
            List.RemoveAt(index)
        End If
    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Public Sub Remove(ByVal nombre As String)
        Dim c As nodocoleccion
        c = Me.Recupera(nombre)
        If Not c Is Nothing Then
            List.Remove(c)
        End If
    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Public Sub Remove_sincascada(ByVal index As Integer)
        If index < Count And index >= 0 Then
            Me.InnerList.RemoveAt(index)
        End If
    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Public Sub Remove_sincascada(ByVal nombre As String)
        Dim c As nodocoleccion
        c = Me.Recupera(nombre)
        If Not c Is Nothing Then
            Me.InnerList.Remove(c)
        End If
    End Sub


    <System.Diagnostics.DebuggerStepThrough()> Protected Function Item(ByVal index As Integer) As nodocoleccion
        Return InnerList.Item(index)
    End Function

    <System.Diagnostics.DebuggerStepThrough()> Protected Function Item(ByVal nombre As String) As nodocoleccion
        Dim c As nodocoleccion

        c = Me.Recupera(nombre)
        If Not c Is Nothing Then
            Return c
        Else
            Throw New Exception("No encontrado elemento: " & v(nombre) & " en coleccion")
        End If
    End Function


    <System.Diagnostics.DebuggerStepThrough()> Protected Function Recupera(ByVal nombre As String) As nodocoleccion
        For Each c As nodocoleccion In Me
            If c.nombre = nombre Then
                Return c
            End If
        Next
        Return Nothing
    End Function

    <System.Diagnostics.DebuggerStepThrough()> Public Function Indice(ByVal pos As Integer) As String
        Return DirectCast(InnerList.Item(pos), nodocoleccion).nombre
    End Function

    <System.Diagnostics.DebuggerStepThrough()> Public Function Existe(ByVal nombre As String) As Boolean
        For Each c As nodocoleccion In Me
            If c.nombre = nombre Then
                Return True
            End If
        Next
        Return False
    End Function


    <System.Diagnostics.DebuggerStepThrough()> Protected Overrides Sub OnClear()
        For Each c As nodocoleccion In Me.List
            If TypeOf c.valor Is coleeccion_object Then
                DirectCast(c.valor, coleeccion_object).Clear()
            End If
        Next
    End Sub


End Class


Public Class ColecciondeLong
    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As Long)
        MyBase.Add(nombre, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Long
        Return CLng(MyBase.Item(index).valor)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Long
        Return CLng(MyBase.Item(nombre).valor)
    End Function



End Class
Public Class ColecciondeDouble
    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal valor As Double)
        MyBase.Add(valor.ToString, valor)
    End Sub

    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As Double)
        MyBase.Add(nombre, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Double
        Return CDbl(MyBase.Item(index).valor)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Double
        Return CDbl(MyBase.Item(nombre).valor)
    End Function



End Class

Public Class ColecciondeInteger
    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal valor As Integer)
        MyBase.Add(valor.ToString, valor)
    End Sub

    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As Integer)
        MyBase.Add(nombre, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Integer
        Return CInt(MyBase.Item(index).valor)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Integer
        Return CInt(MyBase.Item(nombre).valor)
    End Function



End Class

Public Class ColecciondeString

    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As String)
        MyBase.Add(nombre, valor)
    End Sub

    Public Overloads Sub add(ByVal valor As String)
        MyBase.Add((MyBase.Count + 1).ToString, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As String
        Return CStr(MyBase.Item(index).valor)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As String
        Return CStr(MyBase.Item(nombre).valor)
    End Function


End Class

Public Class ColecciondeColeccionDeString
    Inherits coleeccion_object

    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As ColecciondeString)
        MyBase.Add(nombre, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As ColecciondeString
        Return DirectCast(MyBase.Item(index).valor, ColecciondeString)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As ColecciondeString
        Return DirectCast(MyBase.Item(nombre).valor, ColecciondeString)
    End Function
End Class
#If Not COMPACT_FRAMEWORK Then

Public Class parametros_conexion
    Inherits nodocoleccion

    Public usuario As String
    Public pass As String
    Public servidor As String
    Public bd As String
    Public dirrpt As String
    Public litrosdeposito As String

    ' estos son de EMASA, habria que quitarlos de aqui
    '¿esta clase es usada por otro proyecto(s)?
    Public dirimpor As String
    Public dirimporh As String
    Public direxpor As String
    Public direxporh As String

    Public dirsalidaexcel As String

    Public especial1 As Boolean ' Permitir de fecha actuacion posterior a la fecha actual 
End Class

Public Class Colecciondeparametros_conexion

    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As parametros_conexion)
        MyBase.Add(nombre, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As parametros_conexion
        Return DirectCast(MyBase.Item(index).valor, parametros_conexion)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As parametros_conexion
        Return DirectCast(MyBase.Item(nombre).valor, parametros_conexion)
    End Function


End Class



#If Not APLICACION_DE_CONSOLA Then
'Public Class seccion_progreso

'    Inherits nodocoleccion

'    Private _padre As seccion_progreso
'    Private _valoractual As Single

'    Private _maximo, _minimo As Single


'    Dim barraprogreso As ProgressBar
'    Dim etiqueta As Label
'    Dim Panelcontenedor As Panel

'    Dim _numhijos As Integer
'    Dim _texto As String

'    Public Sub New(ByVal f As Form, ByVal texto As String, ByVal minimo As Single, ByVal maximo As Single, Optional ByVal padre As seccion_progreso = Nothing)

'        _texto = texto
'        _numhijos = 0
'        _padre = padre
'        If Not _padre Is Nothing Then
'            _padre.Añadehijo()
'        End If
'        _valoractual = 0 'va a estar entre 0 y 100, siempre
'        _minimo = minimo
'        _maximo = maximo
'    End Sub

'    Public ReadOnly Property Identificador() As String
'        Get
'            Return _texto
'        End Get
'    End Property
'    Public Property ValorMaximo() As Single
'        Get
'            Return _maximo
'        End Get
'        Set(ByVal Value As Single)
'            _maximo = Value
'            Actualiza()
'        End Set
'    End Property
'    Public Property ValorActual() As Single
'        Get
'            Return _valoractual
'        End Get
'        Set(ByVal Value As Single)

'            _valoractual = Value
'            Actualiza()

'        End Set
'    End Property

'    Public Sub Actualiza()
'        If _valoractual >= _maximo Then
'            barraprogreso.Value = 100
'            If Not _padre Is Nothing Then
'                _padre.ValorActual += 1
'            End If
'            _valoractual = _maximo
'            barraprogreso.Value = 100
'        Else
'            barraprogreso.Value = Porcentaje
'        End If

'        Me.etiqueta.Text = _texto & " " & barraprogreso.Value.ToString & "%"
'    End Sub

'    Private Function CalculaPorcentaje() As Integer
'        Return CInt(100 * (_valoractual - _minimo) / (_maximo - _minimo))
'    End Function

'    Public Property Porcentaje() As Integer
'        Get
'            Return CalculaPorcentaje()
'        End Get
'        Set(ByVal Value As Integer)
'            'para que el porcentaje sea value, _valor tiene que ser.....
'            _valoractual = (Value * (_maximo - _minimo) + _minimo) / 100
'        End Set
'    End Property
'    Public Sub Avanza1()
'        ValorActual += 1
'    End Sub


'    Public Sub Avanza(ByVal n As Integer)
'        ValorActual += n
'    End Sub

'    Public Sub Fin()
'        ValorActual = _maximo
'    End Sub


'    Protected Friend Sub Añadehijo()
'        If _numhijos = 0 Then
'            _minimo = 0
'            _maximo = 0
'        End If
'        _numhijos += 1
'        _maximo += 1
'    End Sub

'    Public Sub CreaControles(ByVal f As Form, ByVal i As Integer)

'        Const ALTO As Integer = 64
'        Const ANCHO As Integer = 544
'        barraprogreso = New System.Windows.Forms.ProgressBar
'        Panelcontenedor = New System.Windows.Forms.Panel
'        etiqueta = New System.Windows.Forms.Label


'        f.SuspendLayout()
'        With barraprogreso
'            .Location = New System.Drawing.Point(8, 32)
'            .Name = "p_Principal"
'            .Size = New System.Drawing.Size(520, 23)
'            .TabIndex = 0
'            .Minimum = 0
'            .Maximum = 100
'            .Step = 1
'        End With

'        With etiqueta
'            .Text = _texto
'            .Location = New System.Drawing.Point(16, 8)
'            .Name = "l_principal"
'            .Size = New System.Drawing.Size(512, 23)
'            .TabIndex = 3
'        End With
'        '
'        'PanelPrincipal
'        '
'        With Panelcontenedor
'            .Controls.Add(etiqueta)
'            .Controls.Add(barraprogreso)
'            .Dock = System.Windows.Forms.DockStyle.Bottom
'            .Location = New System.Drawing.Point(0, 0)
'            .Name = "PanelPrincipal"
'            .Size = New System.Drawing.Size(ANCHO, ALTO)
'            .TabIndex = 3
'        End With

'        f.Controls.Add(Panelcontenedor)

'        Dim nuevoalto As Integer
'        nuevoalto = 104 + ALTO * i
'        If nuevoalto > f.Height Then
'            f.Height = nuevoalto
'        End If


'        f.ResumeLayout(True)

'    End Sub
'End Class
'Public Class Coleccion_seccion_progreso

'    Inherits coleeccion_object


'    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As seccion_progreso)
'        MyBase.Add(nombre, valor)
'    End Sub

'    Public Shadows Function Item(ByVal index As Integer) As seccion_progreso
'        Return DirectCast(MyBase.Item(index).valor, seccion_progreso)
'    End Function

'    Public Shadows Function ITem(ByVal nombre As String) As seccion_progreso
'        Return DirectCast(MyBase.Item(nombre).valor, seccion_progreso)
'    End Function


'End Class
'Public Module UTIL_INTERFAZGRAFICO

'    Public Function algunoseleccionado(ByVal g As DataGrid, ByVal numerofilas As Integer) As Boolean
'        Dim i As Integer
'        For i = 0 To numerofilas - 1
'            If g.IsSelected(i) Then
'                Return True
'            End If
'        Next
'        Return False
'    End Function

'    Public Function NuevaColumnadeTexto(ByVal Textocabecera As String, ByVal nombrecampo As String, ByVal ancho As Integer, Optional ByVal formato as string = string.empty) As DataGridTextBoxColumn
'        Dim cs As New System.Windows.Forms.DataGridTextBoxColumn

'        With cs
'            .Format = ""
'            .FormatInfo = Nothing
'            .HeaderText = Textocabecera
'            .MappingName = nombrecampo
'            .ReadOnly = True
'            .Width = ancho

'            .Format = formato



'        End With
'        Return cs
'    End Function

'    Public Function MeasureCorrectedTextWidth(ByVal objGraphics As Graphics, ByVal objFont As Font, ByVal sngWidth As Single, ByVal sngHeight As Single, ByVal sText As String) As Single

'        Dim sngSingleTextWidth As Single
'        Dim sngDoubleTextWidth As Single
'        Dim sngResult As Single = 0.0

'        If sText <> "" Then

'            ' The measurement routine (MeasureCharacterRanges) adds some extra pixels to the result, that we want to discard.
'            ' To do this, we meausure the string and the string duplicated, and the difference is the measure that we want.
'            ' That is:
'            ' A = X + C
'            ' B = 2X + C
'            ' Where A and B are known (the measures) and C is unknown. We are interested in X, which is X = B - A
'            sngSingleTextWidth = MeasureTextWidth(objGraphics, objFont, sngWidth, sngHeight, sText)
'            sngDoubleTextWidth = MeasureTextWidth(objGraphics, objFont, sngWidth * 2, sngHeight, sText & sText)

'            sngResult = sngDoubleTextWidth - sngSingleTextWidth

'        End If

'        Return sngResult

'    End Function

'    Public Function MeasureTextWidth(ByVal objGraphics As Graphics, ByVal objFont As Font, ByVal sngWidth As Single, ByVal sngHeight As Single, ByVal sText As String) As Single

'        Dim sngResult As Single
'        Dim colCharacterRanges(0) As CharacterRange
'        Dim colRegions(1) As Region
'        Dim objStringFormat As StringFormat
'        Dim objLayoutRectangleF As RectangleF
'        Dim objMeasureRectangleF As RectangleF

'        objStringFormat = New StringFormat

'        ' Allow enough width for the bold case
'        If objFont.Bold Then
'            sngWidth = sngWidth * 2
'        End If

'        objLayoutRectangleF = New RectangleF(0, 0, sngWidth, sngHeight)

'        colCharacterRanges(0) = New CharacterRange(0, sText.Length)
'        objStringFormat.SetMeasurableCharacterRanges(colCharacterRanges)
'        colRegions = objGraphics.MeasureCharacterRanges(sText, objFont, objLayoutRectangleF, objStringFormat)
'        objMeasureRectangleF = colRegions(0).GetBounds(objGraphics)
'        sngResult = objMeasureRectangleF.Width

'        Return sngResult

'    End Function

'    Private Function AnchodeTexto(ByVal s As String, ByVal fuente As Font, ByVal mx As Integer, ByVal my As Integer) As Integer
'        Dim imgbmp As Bitmap
'        Dim canvas As Graphics
'        Dim colorpixel As System.Drawing.Color
'        Dim x, y As Integer

'        imgbmp = New Bitmap(mx, my)
'        canvas = Graphics.FromImage(imgbmp)

'        canvas.FillRectangle(Brushes.White, New Rectangle(0, 0, my, my))
'        canvas.DrawString(s, fuente, Brushes.Black, 0, 50)

'        For x = mx - 1 To 0 Step -1
'            For y = my - 1 To 0 Step -1
'                colorpixel = imgbmp.GetPixel(x, y)
'                With colorpixel
'                    If .R <> 255 Or .G <> 255 Or .B <> 255 Then
'                        canvas.Dispose()
'                        imgbmp.Dispose()
'                        Return y
'                    End If
'                End With
'            Next
'        Next
'    End Function


'    Public Function ObtieneBMB(ByVal g As DataGrid) As BindingManagerBase
'        Return g.BindingContext(g.DataSource, g.DataMember)
'    End Function

'    Public Function NumeroRegistros(ByVal g As DataGrid) As Integer
'        Return ObtieneBMB(g).Count
'    End Function
'    'Public Function ValorActual(ByVal c As ComboBox) As ValorCombo

'    '    Return DirectCast(c.SelectedItem, ValorCombo)

'    'End Function




'End Module

'Public Class ToolTipCeldaGrid
'    'NOTA: Usar uno por FORMULARIO. 
'    Implements IDisposable

'    ' Dim emergente As ToolTip
'    '  Dim formulario_padre As Form
'    Dim formulariopadreactivo As Boolean

'    Public Sub New(ByVal fp As Form)
'        emergente = New ToolTip
'        emergente.InitialDelay = 300
'        formulariopadreactivo = True

'        formulario_padre = fp
'        AddHandler formulario_padre.Activated, AddressOf Me.formulario_padre_Activated
'        AddHandler formulario_padre.Deactivate, AddressOf Me.formulario_padre_Deactivate

'    End Sub

'    Public Sub AddGrid(ByVal g As DataGrid)
'        AddHandler g.MouseMove, AddressOf Me.OnMouseMove
'    End Sub
'    Public Sub addCombobox(ByVal c As ComboBox)
'        AddHandler c.MouseMove, AddressOf Me.OnmouseMoveCombobox
'    End Sub
'    Private Sub formulario_padre_Activated(ByVal sender As Object, ByVal e As System.EventArgs)
'        formulariopadreactivo = True
'    End Sub

'    Private Sub formulario_padre_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs)
'        formulariopadreactivo = False
'    End Sub
'    Public Sub OnmouseMoveCombobox(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
'        Dim c As ComboBox
'        Dim tiptext As String

'        If formulariopadreactivo Then
'            c = DirectCast(sender, ComboBox)
'            If Not c.DroppedDown Then
'                If c.SelectedIndex <> -1 Then
'                    tiptext = ValorActual(c).Mostrar
'                    If tiptext <> emergente.GetToolTip(c) Then
'                        If emergente.Active Then
'                            emergente.Active = False
'                        End If
'                        emergente.SetToolTip(c, tiptext)
'                        emergente.Active = True
'                    Else
'                        If Not emergente.Active Then
'                            emergente.Active = True
'                        End If
'                    End If
'                Else
'                    emergente.Active = False
'                End If
'            End If

'        End If
'    End Sub
'    Public Sub OnmouseMoveGrid(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
'        'Dim g As DataGrid

'        'Dim tipText As String
'        'Dim hti As DataGrid.HitTestInfo
'        'Dim bmb As BindingManagerBase

'        'If formulariopadreactivo Then
'        '    g = DirectCast(sender, DataGrid)
'        '    If Not (g.DataSource Is Nothing) Then
'        '        bmb = ObtieneBMB(g)
'        '        hti = g.HitTest(New Point(e.X, e.Y))
'        '        If hti.Row < bmb.Count AndAlso hti.Type = DataGrid.HitTestType.Cell Then
'        '            tipText = Trim(g(hti.Row, hti.Column).ToString())
'        '            If tipText <> emergente.GetToolTip(g) Then
'        '                If emergente.Active Then
'        '                    emergente.Active = False
'        '                End If
'        '                emergente.SetToolTip(g, tipText)
'        '                emergente.Active = True
'        '            Else
'        '                If Not emergente.Active Then
'        '                    emergente.Active = True
'        '                End If
'        '            End If
'        '        Else
'        '            emergente.Active = False
'        '        End If
'        '    End If
'        'End If
'    End Sub
'    Public Sub OnMouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
'        'Dim g As DataGrid

'        'Dim tipText As String
'        'Dim hti As DataGrid.HitTestInfo
'        'Dim bmb As BindingManagerBase

'        'g = DirectCast(sender, DataGrid)

'        'If formulariopadreactivo Then

'        '    If Not (g.DataSource Is Nothing) Then
'        '        bmb = ObtieneBMB(g)
'        '        hti = g.HitTest(New Point(e.X, e.Y))
'        '        If hti.Row < bmb.Count AndAlso hti.Type = DataGrid.HitTestType.Cell Then
'        '            tipText = Trim(g(hti.Row, hti.Column).ToString())
'        '            If tipText <> emergente.GetToolTip(g) Then
'        '                If emergente.Active Then
'        '                    emergente.Active = False
'        '                End If
'        '                emergente.SetToolTip(g, tipText)
'        '                emergente.Active = True
'        '            Else
'        '                If Not emergente.Active Then
'        '                    emergente.Active = True
'        '                End If
'        '            End If
'        '        Else
'        '            emergente.Active = False
'        '        End If
'        '    End If
'        'End If
'    End Sub


'    Public Sub Dispose() Implements System.IDisposable.Dispose
'        'emergente.Dispose()
'    End Sub
'End Class

Public Class DescripcionColumnaGrid
    Implements IDisposable

    Public cabecera As String
    Public campo As String
    Public ancho As Integer
    Public orden As Integer

    ' Dim columnatexto As DataGridTextBoxColumn

    Public Sub New(ByVal textocabecera As String, ByVal nombrecampo As String, ByVal anchocolumna As Integer, Optional ByVal formato As String = "")
        cabecera = textocabecera
        campo = nombrecampo
        ancho = anchocolumna

        '   columnatexto = NuevaColumnadeTexto(cabecera, campo, ancho, formato)
    End Sub

    'ReadOnly Property ColumnadeTexto() As DataGridTextBoxColumn
    '    Get
    '        Return columnatexto
    '    End Get

    'End Property

    Public Sub Dispose() Implements System.IDisposable.Dispose
        ' columnatexto.Dispose()
    End Sub
End Class

Public Class ColeccionDescripcionColumnaGrid
    Inherits CollectionBase
    Implements IDisposable

    Dim elgrid As DataGrid
    Dim anchobase As Integer
    Dim numerocolumnasvisibles As Integer
    Dim _v As DataView

    Public Sub New()
        MyBase.new()
    End Sub

    Public Sub Añade(ByVal d As DescripcionColumnaGrid)
        List.Add(d)
    End Sub

    Public Function Posicion(ByVal i As Integer) As DescripcionColumnaGrid
        Return DirectCast(List(i), DescripcionColumnaGrid)
    End Function


    Public Sub inicializaGrid(ByVal g As DataGrid, ByVal dv As DataView)
        Inicializa(g, dv)
    End Sub

    Private Function CreaVista(ByVal t As data.DataTable, ByVal filtro As String, ByVal ordenacion As String) As DataView
        Dim dv As DataView
        dv = New DataView(t)
        With dv
            .AllowDelete = False
            .AllowEdit = False
            .AllowNew = False
            .RowFilter = filtro
            .Sort = ordenacion
        End With

        Return dv
    End Function

    Public Sub inicializaGrid(ByVal g As DataGrid, ByVal t As data.DataTable, ByVal filtro As String, ByVal ordenacion As String)
        Inicializa(g, CreaVista(t, filtro, ordenacion))
    End Sub

    Private Sub Inicializa(ByVal g As DataGrid, ByVal dv As DataView)
        'Dim dgts As DataGridTableStyle
        'Dim dc As DescripcionColumnaGrid


        'AddHandler g.Resize, AddressOf Me.Resize
        'dgts = New DataGridTableStyle
        'numerocolumnasvisibles = 0
        'anchobase = 0
        'With dgts
        '    .HeaderForeColor = System.Drawing.SystemColors.ControlText
        '    .MappingName = dv.Table.TableName
        '    For Each dc In Me
        '        dc.orden = .GridColumnStyles.Add(dc.ColumnadeTexto)
        '        anchobase += dc.ancho
        '        If dc.ancho > 0 Then
        '            numerocolumnasvisibles += 1
        '        End If
        '    Next
        'End With

        'g.TableStyles.Clear()
        'g.TableStyles.Add(dgts)
        '_v = dv
        'g.DataSource = dv
        'g.RowHeadersVisible = False

        'RecalculaAnchoCOlumnas(g)
    End Sub

    Public ReadOnly Property vista() As Data.DataView
        Get
            Return _v
        End Get
    End Property

    Public Sub ActualizafuenteDedatos(ByVal g As DataGrid, ByVal t As data.DataTable, ByVal f As String, ByVal o As String)
        ActualizafuenteDedatos(g, CreaVista(t, f, o), f, o)
    End Sub

    Public Sub ActualizafuenteDeDatos(ByVal g As DataGrid, ByVal v As DataView, ByVal f As String, ByVal o As String)
        'v.RowFilter = f
        '_v = v
        'If o <> "" Then
        '    v.Sort = o
        'End If
        'g.DataSource = v
        'g.Refresh()
    End Sub
    Public Sub Resize(ByVal sender As Object, ByVal e As System.EventArgs)
        RecalculaAnchoCOlumnas(DirectCast(sender, DataGrid))
    End Sub

    Private Sub RecalculaAnchoCOlumnas(ByVal g As DataGrid)
        ''recalculamos el tamaño de las columnas
        'Dim anchototal As Integer
        'Dim dc As DescripcionColumnaGrid
        'Dim ancho As Integer
        'Const ANCHOCOLUMNA As Integer = 1
        'Dim tot As Integer

        ''quitamos la primera columna y las lineas de separacion
        'anchototal = g.Width - g.RowHeaderWidth - numerocolumnasvisibles * ANCHOCOLUMNA
        ''anchototal = g.Width
        'For Each dc In Me
        '    If dc.ancho > 0 Then
        '        ancho = Math.Floor(CSng(dc.ancho * anchototal) / CSng(anchobase))
        '        dc.ColumnadeTexto.Width = ancho

        '        tot += ancho
        '    End If
        'Next
    End Sub


    Public Sub Dispose() Implements System.IDisposable.Dispose
        For Each dc As DescripcionColumnaGrid In Me
            dc.Dispose()
        Next
    End Sub


    Public Function ObtenNumeroColumna(ByVal nombre As String) As Integer
        '    Dim gcs As DataGridColumnStyle
        '    Dim estilos As GridColumnStylesCollection
        '    estilos = g.TableStyles(0).GridColumnStyles
        '    For Each gcs In estilos
        '    If gcs.MappingName = nombre Then
        '    Return estilos.IndexOf(gcs)
        '    End If
        '    Next
        Dim resultado As Integer = 0
        For Each i As DescripcionColumnaGrid In List
            If i.campo = nombre Then
                Return List.IndexOf(i)
            End If
        Next

        Err.Raise(1, "ObtenNumeroColumna", "Columna no encontrada " & nombre)
        Return resultado
    End Function
End Class

Public Class ColecciondeControl
    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As Control)
        MyBase.Add(nombre, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Control
        Return CType(MyBase.Item(index).valor, Control)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Control
        Return CType(MyBase.Item(nombre).valor, Control)
    End Function

 


End Class




#End If

#End If



