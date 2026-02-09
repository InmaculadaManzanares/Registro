'	Crypto class
'	--
'	
Public Class Crypto
    Private Const TAMK As Integer = 255
    Private Const TAMK1 As Integer = TAMK + 1

    Private KE(TAMK) As Byte
    Private KE0(TAMK) As Byte
    Private KD(TAMK) As Byte
    Private KD0(TAMK) As Byte
    Private AKE As Byte
    Private AKD As Byte
    Private IKE As Integer
    Private IKD As Integer

    Public Sub New(ByVal passphrase As String)
        Me.Iniciar(passphrase)
    End Sub

    Public Sub Iniciar(ByVal passphrase As String)
        Dim i As Integer
        Dim clave As Byte()

        clave = GetCryptoKey(passphrase)
        For i = 0 To TAMK
            KE(i) = clave(i)
            KD(i) = KE(i)
            KE0(i) = clave(i) \ CByte(2)
            KD0(i) = KE0(i)
        Next i
        IKD = 0
        IKE = 0
        AKE = clave(4) ' Porque me da la gana
        AKD = AKE
        clave = Nothing
    End Sub
    Public Function CadenaaBytes(ByVal cad As String) As Byte()
        Dim encodingASCII As New System.Text.ASCIIEncoding
        Return encodingASCII.GetBytes(cad)
    End Function
    Public Function BytesaCadena(ByVal b As Byte()) As String
        Dim i As Integer
        Dim temp as string = string.empty

        For i = b.GetLowerBound(0) To b.GetUpperBound(0)
            temp &= Chr(b(i))
        Next i
        Return temp
    End Function
    Public Function Encriptar(ByVal entrada As String) As String
        Return Me.CodificarLegible(Me.Encriptar(Me.CadenaaBytes(entrada)))
    End Function
    Public Function desencriptar(ByVal entrada As String) As String
        Return Me.BytesaCadena(Me.Desencriptar(Me.DecodificardeLegible(entrada)))
    End Function
    Public Function Encriptar(ByRef entrada() As Byte) As Byte()
        Dim fin As Integer, i As Integer, jk As Integer, m As Integer
        Dim vant, v(), vi As Byte

        fin = entrada.GetUpperBound(0)
        ReDim v(fin)

        System.Buffer.BlockCopy(entrada, 0, v, 0, fin + 1)
        For i = 0 To fin
            vi = v(i)
            vant = Not vi
            v(i) = vi Xor (KE(IKE) Xor (Not (KE0(IKE) Xor AKE)))
            AKE = AKE Xor vi
            'v(i) = vi Xor vtemp
            m = Math.Max(IKE - 8, 0)
            For jk = IKE To m Step -1
                vant = vant >> 1
                KE0(jk) = KE0(jk) Xor vant
            Next jk
            IKE = (IKE + 1) Mod TAMK1
        Next i
        Return v
    End Function

    Public Function Desencriptar(ByRef entrada() As Byte) As Byte()
        Dim fin As Integer, i As Integer, jk As Integer, m As Integer
        Dim v(), vant, vi As Byte

        fin = entrada.GetUpperBound(0)
        ReDim v(fin)

        System.Buffer.BlockCopy(entrada, 0, v, 0, fin + 1)


        For i = 0 To fin
            v(i) = v(i) Xor (KD(IKD) Xor (Not (KD0(IKD) Xor AKD)))
            vi = v(i)
            AKD = AKD Xor vi
            vant = Not vi
            m = Math.Max(IKD - 8, 0)
            For jk = IKD To m Step -1
                vant = vant >> 1
                KD0(jk) = KD0(jk) Xor vant
            Next jk
            IKD = (IKD + 1) Mod TAMK1
        Next i
        Return v
    End Function
    Private Function GetCryptoKey(ByVal passphrase As String) As Byte()
        Dim i, j, lfrase, vant As Integer
        Dim result(), frasebyte() As Byte

        lfrase = passphrase.Length - 1
        frasebyte = Util.CadenaaBytes(passphrase)
        result = Hash(frasebyte)

        vant = 13 Mod (lfrase + 1)
        For i = 0 To TAMK - lfrase
            For j = 0 To lfrase
                result(i + j) = (result(i + j)) Xor frasebyte((i + j + vant) Mod (lfrase + 1))
                vant = (result(i + j)) Xor vant
            Next j
        Next i
        Return result
    End Function

    Private Function Hash(ByVal b As Byte()) As Byte()
        Dim base() As Byte
        'Dim temp() As Byte
        Dim i, indice, incremento, tamactual As Integer
        Dim cadena as string = string.empty

        'ReDim temp(TAMK)
        ReDim base(TAMK)

        System.Buffer.BlockCopy(b, 0, base, 0, b.Length)
        tamactual = b.Length

        Do
            For i = 0 To tamactual - 1
                cadena = cadena & base(i).ToString
            Next

            cadena = Mid(cadena & StrReverse(cadena), 1, TAMK)
            indice = cadena.Length

            'indice = 0
            'For i = 0 To tamactual - 1
            '    cadena = base(i).ToString
            '    For j = 0 To cadena.Length - 1
            '        temp(indice) = CByte(Asc(cadena.Chars(j)))
            '        indice = indice + 1
            '        If indice > TAMK Then
            '            Exit For
            '        End If
            '    Next j
            '    If indice > TAMK Then
            '        Exit For
            '    End If
            'Next i

            'For i = tamactual - 1 To 0 Step -1
            '    If indice > TAMK Then
            '        Exit For
            '    End If
            '    cadena = temp(i).ToString
            '    For j = cadena.Length - 1 To 0 Step -1
            '        temp(indice) = CByte(Asc(cadena.Chars(j)))
            '        indice = indice + 1
            '        If indice > TAMK Then
            '            Exit For
            '        End If
            '    Next j
            'Next i

            i = 0
            tamactual = 0
            While i < indice
                'cadena = ""
                'incremento = 2 - (tamactual Mod 2)
                'If i + incremento >= indice Then
                '    incremento = indice - 1 - i
                'End If
                'For j = i To i + incremento
                '    cadena = Chr(temp(j)) + cadena
                'Next j

                incremento = Math.Min((tamactual Mod 3) + 1, indice - i)
                base(tamactual) = CByte(CInt(Mid(cadena, i + 1, incremento)) Mod 256)
                tamactual += 1
                i += incremento
            End While
            reordena(base, tamactual)
        Loop While indice < TAMK
        'temp = Nothing
        Return base
    End Function


    Private Sub reordena(ByRef b As Byte(), ByVal maxindice As Integer)
        Dim temp As Byte
        Dim pizq, pder, iizq, ider, I As Integer
        
        I = maxindice + 1
        pizq = 0
        pder = maxindice
        While (pder > pizq)
            iizq = b(pizq) Mod I
            ider = b(pder) Mod I
            temp = b(iizq)
            b(iizq) = b(ider)
            b(ider) = temp
            pder = pder - 1
            pizq = pizq + 1
        End While
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Function CodificarLegible(ByVal origen() As Byte) As String
        Dim fin As Integer
        Dim i As Integer
        Dim s As String
        Dim v As String

        fin = UBound(origen)
        s = ""

        For i = 0 To fin
            v = Hex(origen(i))
            If Len(v) = 1 Then
                s = s & "0"
            End If
            s = s & v
        Next i
        Return s
    End Function

    Private Function DecodificardeLegible(ByVal origen As String) As Byte()
        Dim i As Integer, j As Integer
        Dim fin As Integer
        Dim destino() As Byte

        fin = Len(origen)
        ReDim destino((fin \ 2) - 1)
        j = 0

        For i = 1 To fin Step 2
            destino(j) = Hex2Dec(Mid(origen, i, 1), Mid(origen, i + 1, 1))
            j = j + 1
        Next i
        Return destino
    End Function

    Private Function Hex2Dec(ByVal v1 As String, ByVal v2 As String) As Byte
        Hex2Dec = Hex2Dec_unitario(v1) * CByte(16) + Hex2Dec_unitario(v2)
    End Function

    Private Function Hex2Dec_unitario(ByVal v As String) As Byte
        If v Like "[0-9]" Then
            Hex2Dec_unitario = CByte(v)
        ElseIf v Like "[A-F]" Then
            Hex2Dec_unitario = CByte(10) + (CByte(Asc(v) - Asc("A")))
        Else
            Hex2Dec_unitario = 0
        End If
    End Function


End Class
