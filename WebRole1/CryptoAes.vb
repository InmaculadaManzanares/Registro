Imports System
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

' AES-CBC + HMAC-SHA256 (Encrypt-then-MAC)
' Formato: [1 byte version][16 bytes IV][ciphertext...][32 bytes HMAC]
' HMAC se calcula sobre (version + IV + ciphertext)
Public NotInheritable Class CryptoAes

    Private Sub New()
    End Sub

    Private Const VERSION As Byte = 1
    Private Const IV_LEN As Integer = 16
    Private Const HMAC_LEN As Integer = 32

    ' masterKey: 32 bytes (256-bit) recomendado
    ' De masterKey derivamos:
    '  - encKey (32 bytes) para AES-256
    '  - macKey (32 bytes) para HMAC-SHA256
    Private Shared Sub DeriveKeys(masterKey As Byte(), ByRef encKey As Byte(), ByRef macKey As Byte())
        ' Derivación simple y auditable por hashing (suficiente para separar claves)
        ' encKey = SHA256("enc" || masterKey)
        ' macKey = SHA256("mac" || masterKey)
        Using sha As SHA256 = SHA256.Create()
            encKey = sha.ComputeHash(Concat(Encoding.UTF8.GetBytes("enc"), masterKey))
            macKey = sha.ComputeHash(Concat(Encoding.UTF8.GetBytes("mac"), masterKey))
        End Using
    End Sub

    Public Shared Function EncryptToBase64Url(plainText As String, masterKeyB64 As String) As String
        If plainText Is Nothing Then Throw New ArgumentNullException(NameOf(plainText))
        Dim masterKey As Byte() = Convert.FromBase64String(masterKeyB64)
        If masterKey.Length < 32 Then Throw New ArgumentException("La clave maestra debe ser Base64 de al menos 32 bytes.", NameOf(masterKeyB64))

        Dim encKey As Byte() = Nothing
        Dim macKey As Byte() = Nothing
        DeriveKeys(masterKey, encKey, macKey)

        Dim iv As Byte() = New Byte(IV_LEN - 1) {}
        Using rng As New RNGCryptoServiceProvider()
            rng.GetBytes(iv)
        End Using

        Dim cipher As Byte()
        Using aes As Aes = Aes.Create()
            aes.Mode = CipherMode.CBC
            aes.Padding = PaddingMode.PKCS7
            aes.KeySize = 256
            aes.Key = encKey
            aes.IV = iv

            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write)
                    Dim pt As Byte() = Encoding.UTF8.GetBytes(plainText)
                    cs.Write(pt, 0, pt.Length)
                    cs.FlushFinalBlock()
                    cipher = ms.ToArray()
                End Using
            End Using
        End Using

        ' payload sin HMAC: version + iv + cipher
        Dim header As Byte() = New Byte(1 + IV_LEN + cipher.Length - 1) {}
        header(0) = VERSION
        Buffer.BlockCopy(iv, 0, header, 1, IV_LEN)
        Buffer.BlockCopy(cipher, 0, header, 1 + IV_LEN, cipher.Length)

        Dim tag As Byte()
        Using h As New HMACSHA256(macKey)
            tag = h.ComputeHash(header)
        End Using

        Dim full As Byte() = New Byte(header.Length + tag.Length - 1) {}
        Buffer.BlockCopy(header, 0, full, 0, header.Length)
        Buffer.BlockCopy(tag, 0, full, header.Length, tag.Length)

        Return Base64UrlEncode(full)
    End Function

    Public Shared Function DecryptFromBase64Url(tokenB64Url As String, masterKeyB64 As String) As String
        If String.IsNullOrWhiteSpace(tokenB64Url) Then Throw New ArgumentNullException(NameOf(tokenB64Url))
        Dim masterKey As Byte() = Convert.FromBase64String(masterKeyB64)
        If masterKey.Length < 32 Then Throw New ArgumentException("La clave maestra debe ser Base64 de al menos 32 bytes.", NameOf(masterKeyB64))

        Dim encKey As Byte() = Nothing
        Dim macKey As Byte() = Nothing
        DeriveKeys(masterKey, encKey, macKey)

        Dim full As Byte() = Base64UrlDecode(tokenB64Url)
        If full.Length < 1 + IV_LEN + 1 + HMAC_LEN Then
            Throw New CryptographicException("Token inválido (tamaño insuficiente).")
        End If

        Dim version As Byte = full(0)
        If version <> version Then
            Throw New CryptographicException("Versión de token no soportada.")
        End If

        Dim headerLen As Integer = full.Length - HMAC_LEN
        Dim header As Byte() = New Byte(headerLen - 1) {}
        Dim tag As Byte() = New Byte(HMAC_LEN - 1) {}

        Buffer.BlockCopy(full, 0, header, 0, headerLen)
        Buffer.BlockCopy(full, headerLen, tag, 0, HMAC_LEN)

        ' Verificar HMAC (integridad)
        Dim expected As Byte()
        Using h As New HMACSHA256(macKey)
            expected = h.ComputeHash(header)
        End Using

        If Not FixedTimeEquals(tag, expected) Then
            Throw New CryptographicException("Token inválido (HMAC no coincide).")
        End If

        Dim iv As Byte() = New Byte(IV_LEN - 1) {}
        Buffer.BlockCopy(header, 1, iv, 0, IV_LEN)

        Dim cipherLen As Integer = headerLen - 1 - IV_LEN
        Dim cipher As Byte() = New Byte(cipherLen - 1) {}
        Buffer.BlockCopy(header, 1 + IV_LEN, cipher, 0, cipherLen)

        Using aes As Aes = Aes.Create()
            aes.Mode = CipherMode.CBC
            aes.Padding = PaddingMode.PKCS7
            aes.KeySize = 256
            aes.Key = encKey
            aes.IV = iv

            Using ms As New MemoryStream(cipher)
                Using cs As New CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Read)
                    Using sr As New StreamReader(cs, Encoding.UTF8)
                        Return sr.ReadToEnd()
                    End Using
                End Using
            End Using
        End Using
    End Function

    Private Shared Function FixedTimeEquals(a As Byte(), b As Byte()) As Boolean
        If a Is Nothing OrElse b Is Nothing OrElse a.Length <> b.Length Then Return False
        Dim diff As Integer = 0
        For i As Integer = 0 To a.Length - 1
            diff = diff Or (a(i) Xor b(i))
        Next
        Return diff = 0
    End Function

    Private Shared Function Concat(a As Byte(), b As Byte()) As Byte()
        Dim r As Byte() = New Byte(a.Length + b.Length - 1) {}
        Buffer.BlockCopy(a, 0, r, 0, a.Length)
        Buffer.BlockCopy(b, 0, r, a.Length, b.Length)
        Return r
    End Function

    Private Shared Function Base64UrlEncode(data As Byte()) As String
        Dim s As String = Convert.ToBase64String(data)
        s = s.TrimEnd("="c).Replace("+"c, "-"c).Replace("/"c, "_"c)
        Return s
    End Function

    Private Shared Function Base64UrlDecode(s As String) As Byte()
        Dim b64 As String = s.Replace("-"c, "+"c).Replace("_"c, "/"c)
        Select Case (b64.Length Mod 4)
            Case 2 : b64 &= "=="
            Case 3 : b64 &= "="
        End Select
        Return Convert.FromBase64String(b64)
    End Function

End Class

