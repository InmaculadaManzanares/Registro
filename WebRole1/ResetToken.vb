Imports System
Imports System.Security.Cryptography

Public Module ResetToken

    Public Function GenerateTokenBase64Url(bytesLen As Integer) As String
        Dim data(bytesLen - 1) As Byte
        Using rng = RandomNumberGenerator.Create()
            rng.GetBytes(data)
        End Using
        Dim b64 = Convert.ToBase64String(data)
        ' Base64Url (sin + / =)
        b64 = b64.Replace("+"c, "-"c).Replace("/"c, "_"c).TrimEnd("="c)
        Return b64
    End Function

    Public Function Sha256Bytes(input As String) As Byte()
        Dim bytes = System.Text.Encoding.UTF8.GetBytes(input)
        Using sha As SHA256 = SHA256.Create()
            Return sha.ComputeHash(bytes)
        End Using
    End Function

    Public Function FixedTimeEquals(a As Byte(), b As Byte()) As Boolean
        If a Is Nothing OrElse b Is Nothing Then Return False
        If a.Length <> b.Length Then Return False
        Dim diff As Integer = 0
        For i As Integer = 0 To a.Length - 1
            diff = diff Or (a(i) Xor b(i))
        Next
        Return diff = 0
    End Function

End Module
