Public Module Alea
    Private inicializado As Boolean = False
    Private objrandom As New System.Random

    Private result As String

    Public Sub Inicializa()
        'ahora no hay "randomize"
    End Sub

    Public Function CadenaAleatoria(ByVal numchar As Integer) As String
        Dim i As Integer
        Dim s As String
        s = ""
        For i = 1 To numchar
            s &= Chr(objrandom.Next(65, 90))
        Next
        Return s
    End Function

    Public Function aleatorio_en_rango(ByVal ini As Integer, ByVal fin As Integer) As Integer
        Return objrandom.Next(ini, fin)
    End Function

End Module