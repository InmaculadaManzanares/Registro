Imports System
Imports System.Data
Imports System.Text.RegularExpressions

Partial Public Class ResetPassword
    Inherits System.Web.UI.Page

    Private Function Q(name As String) As String
        Dim vq = Request.QueryString(name)
        If vq Is Nothing Then Return ""
        Return vq.Trim()
    End Function

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim loginUsuario As String = Q("u")
            Dim token As String = Q("t")

            If loginUsuario = "" OrElse token = "" Then
                ShowError("Enlace no válido.")
                panelReset.Visible = False
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub btnSet_Click(sender As Object, e As EventArgs) Handles btnSet.Click
        Dim loginUsuario As String = Q("u")
        Dim token As String = Q("t")

        If loginUsuario = "" OrElse token = "" Then
            ShowError("Enlace no válido.")
            panelReset.Visible = False
            Exit Sub
        End If

        If txtNewPass.Text = "" OrElse txtNewPass2.Text = "" Then
            ShowError("Debe indicar la nueva contraseña dos veces.")
            Exit Sub
        End If

        If txtNewPass.Text <> txtNewPass2.Text Then
            ShowError("Las contraseñas no coinciden.")
            Exit Sub
        End If

        ' Validación de contraseña (INCIBE: longitud + mayúscula + minúscula + número + especial)
        Dim msg As String = ""
        If Not PasswordCumplePolitica(txtNewPass.Text, msg) Then
            ShowError(msg)
            Exit Sub
        End If

        Dim con As New ConexionSQL(VariablesGlobales.CadenadeConexion(Application))
        Try
            ' 1) Buscar último token válido no usado para ese usuario(login)
            Dim sql As String =
                "SELECT TOP 1 Id, TokenHash " &
                "FROM PasswordResetTokens " &
                "WHERE Email=" & v(loginUsuario) & " AND Used=0 AND ExpiresAt >= GETDATE() " &
                "ORDER BY Id DESC"

            Dim dt As DataTable = con.SelectSQL(sql)
            If dt.Rows.Count = 0 Then
                ShowError("Token caducado o inválido.")
                Exit Sub
            End If

            Dim tokenId As Integer = CInt(dt.Rows(0)("Id"))
            Dim dbHash As Byte() = DirectCast(dt.Rows(0)("TokenHash"), Byte())

            ' 2) Comparar hash del token (tiempo constante)
            Dim tokenHash As Byte() = ResetToken.Sha256Bytes(token)
            If Not ResetToken.FixedTimeEquals(dbHash, tokenHash) Then
                ShowError("Token inválido.")
                Exit Sub
            End If

            ' 3) Guardar nueva contraseña (hash)
            Dim newStored As String = PasswordHasher.HashPassword(txtNewPass.Text)

            con.Ejecuta("UPDATE Usuarios SET password=" & v(newStored) &
            ", PasswordLastChangedUtc = GETUTCDATE() WHERE Email=" & v(loginUsuario))

            con.Ejecuta("UPDATE PasswordResetTokens SET Used=1 WHERE Id=" & tokenId.ToString())

            panelReset.Visible = False
            panelOk.Visible = True
            labelmensaje.Text = "Contraseña actualizada correctamente."

        Catch ex As Exception
            ShowError("Error al actualizar la contraseña.")
        Finally
            If con IsNot Nothing Then con.CerrarConexion()
        End Try
    End Sub

    ' Política de contraseña tipo INCIBE (mínimo 8, mayúscula, minúscula, número, especial)
    Private Function PasswordCumplePolitica(pass As String, ByRef errorMsg As String) As Boolean
        If pass Is Nothing Then
            errorMsg = "Contraseña no válida."
            Return False
        End If

        If pass.Length < 8 Then
            errorMsg = "La contraseña debe tener al menos 8 caracteres."
            Return False
        End If

        If Not Regex.IsMatch(pass, "[A-Z]") Then
            errorMsg = "La contraseña debe contener al menos una letra mayúscula."
            Return False
        End If

        If Not Regex.IsMatch(pass, "[a-z]") Then
            errorMsg = "La contraseña debe contener al menos una letra minúscula."
            Return False
        End If

        If Not Regex.IsMatch(pass, "[0-9]") Then
            errorMsg = "La contraseña debe contener al menos un número."
            Return False
        End If

        If Not Regex.IsMatch(pass, "[^a-zA-Z0-9]") Then
            errorMsg = "La contraseña debe contener al menos un carácter especial."
            Return False
        End If

        ' Opcional: evitar espacios
        If Regex.IsMatch(pass, "\s") Then
            errorMsg = "La contraseña no puede contener espacios."
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkVolver_Click(sender As Object, e As EventArgs) Handles lnkVolver.Click
        Response.Redirect("~/RegistroHorario.aspx")
    End Sub

    Protected Sub lnkIrLogin_Click(sender As Object, e As EventArgs) Handles lnkIrLogin.Click
        Response.Redirect("~/RegistroHorario.aspx")
    End Sub

    Private Sub ShowError(msg As String)
        mensaje.Text = msg
        Dim script As String = "MostrarAlerta();"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "alerta", script, True)
    End Sub

End Class
