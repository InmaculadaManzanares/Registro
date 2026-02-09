Imports Microsoft.VisualBasic
'hijos
Public Class Hijos

    Public idboton As String
    Public Nombreboton As String
    Public url As String
    Public urlhist As String
    Public Codigodocumento As String

End Class

Public Class Coleccion_Hijos
    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal valor As Hijos)
        MyBase.Add(valor.idboton, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As Hijos
        Return DirectCast(MyBase.Item(index).valor, Hijos)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As Hijos
        Return DirectCast(MyBase.Item(nombre).valor, Hijos)
    End Function


End Class
