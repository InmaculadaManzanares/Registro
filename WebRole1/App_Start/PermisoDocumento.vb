Imports Microsoft.VisualBasic

Public Class PermisoDocumento

    Public idboton As String

    Public estadoorigen As Integer
    Public estadodestino As Integer
    Public Nombreboton As String
    Public Nombreaccion As String
    Public CampoFechaDoc As String
    Public campusuarioDoc As String
    Public CamposaLimpiar As String
    Public CamposObligatoriosEnTransicion As String
End Class


Public Class Coleccion_PermisoDocumento
    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal valor As PermisoDocumento)
        MyBase.Add(valor.idboton, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As PermisoDocumento
        Return DirectCast(MyBase.Item(index).valor, PermisoDocumento)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As PermisoDocumento
        Return DirectCast(MyBase.Item(nombre).valor, PermisoDocumento)
    End Function



End Class

Public Class EstadoDocumento
    Public Estado As Integer
    Public puedeBorrarse As Boolean
    Public puedeEditarse As Boolean
End Class

Public Class Coleccion_EstadoDocumento
    Inherits coleeccion_object


    Public Overloads Sub Add(ByVal valor As EstadoDocumento)
        MyBase.Add(valor.Estado.ToString, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As EstadoDocumento
        Return DirectCast(MyBase.Item(index).valor, EstadoDocumento)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As EstadoDocumento
        Return DirectCast(MyBase.Item(nombre).valor, EstadoDocumento)
    End Function



End Class