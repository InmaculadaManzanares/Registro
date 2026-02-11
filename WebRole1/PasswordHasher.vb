Imports System
Imports System.Security.Cryptography

Public Module PasswordHasher

    ' En .NET 4.5 (PBKDF2-HMAC-SHA1). Sube iteraciones según rendimiento.
    Private Const Iterations As Integer = 300000
    Private Const SaltSize As Integer = 16  ' 128-bit
    Private Const KeySize As Integer = 32   ' 256-bit

    Public Function HashPassword(password As String) As String
        If String.IsNullOrEmpty(password) Then Throw New ArgumentException("Password vacío")

        Dim salt(SaltSize - 1) As Byte
        Using rng = RandomNumberGenerator.Create()
            rng.GetBytes(salt)
        End Using

        Dim key As Byte()
        Using pbkdf2 As New Rfc2898DeriveBytes(password, salt, Iterations)
            key = pbkdf2.GetBytes(KeySize)
        End Using

        Dim saltB64 = Convert.ToBase64String(salt)
        Dim keyB64 = Convert.ToBase64String(key)

        ' Formato versionado para poder migrar a SHA256/Argon2 cuando actualices framework
        Return $"v1|pbkdf2-sha1|{Iterations}|{saltB64}|{keyB64}"
    End Function

    Public Function VerifyPassword(password As String, stored As String) As Boolean
        If String.IsNullOrEmpty(password) OrElse String.IsNullOrEmpty(stored) Then Return False

        Dim parts = stored.Split("|"c)
        If parts.Length <> 5 Then Return False
        If parts(0) <> "v1" Then Return False

        Dim algo = parts(1)
        If algo <> "pbkdf2-sha1" Then Return False

        Dim iters As Integer
        If Not Integer.TryParse(parts(2), iters) Then Return False

        Dim salt As Byte()
        Dim expected As Byte()
        Try
            salt = Convert.FromBase64String(parts(3))
            expected = Convert.FromBase64String(parts(4))
        Catch
            Return False
        End Try

        Dim actual As Byte()
        Using pbkdf2 As New Rfc2898DeriveBytes(password, salt, iters)
            actual = pbkdf2.GetBytes(expected.Length)
        End Using

        Return FixedTimeEquals(actual, expected)
    End Function

    ' Comparación en tiempo constante (evita timing attacks)
    Private Function FixedTimeEquals(a As Byte(), b As Byte()) As Boolean
        If a Is Nothing OrElse b Is Nothing Then Return False
        If a.Length <> b.Length Then Return False

        Dim diff As Integer = 0
        For i As Integer = 0 To a.Length - 1
            diff = diff Or (a(i) Xor b(i))
        Next
        Return diff = 0
    End Function

End Module
