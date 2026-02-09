Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System
Imports System.Collections
Imports System.Linq
Imports System.Web

Imports System.Xml.Linq
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
<System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
 Public Class AutoComplete
    Inherits System.Web.Services.WebService

    Dim cn As New SqlClient.SqlConnection()
    Dim ds As New DataSet
    Dim dt As New DataTable
    <WebMethod()> _
     Public Function GetCustomers(ByVal prefix As String) As String()
        Dim customers As New List(Of String)()
        Using conn As New SqlConnection()
            conn.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Using cmd As New SqlCommand()
                'cmd.CommandText = "select Nombre,Email from Usuarios where " & "Nombre like @SearchText + '%'"
                cmd.CommandText = "SELECT top (10) (CODPOSTAL + ' ' + Poblacion) as nombre,  codpostal   FROM   CODPOSTAL where (CODPOSTAL  + ' ' + Poblacion)  like '%' + @SearchText + '%' order by codpostal"
                'cmd.CommandText = "select (Nombre + ' ' +Email) as Nombre, Email from Usuarios where " & "(Nombre + ' ' +Email) like @SearchText + '%'"
                cmd.Parameters.AddWithValue("@SearchText", prefix)
                cmd.Connection = conn
                conn.Open()
                Using sdr As SqlDataReader = cmd.ExecuteReader()
                    While sdr.Read()
                        customers.Add(String.Format("{0}-{1}", sdr("nombre"), sdr("codpostal")))
                    End While
                End Using
                conn.Close()
            End Using
            Return customers.ToArray()
        End Using
    End Function

    <WebMethod()> _
    Public Function GetCodPostal(ByVal prefix As String) As String()
        Dim customers As New List(Of String)()
        Using conn As New SqlConnection()
            conn.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Using cmd As New SqlCommand()
                'cmd.CommandText = "select Nombre,Email from Usuarios where " & "Nombre like @SearchText + '%'"
                cmd.CommandText = "SELECT  codpostal  FROM   CODPOSTAL group by CODPOSTAL having  (CODPOSTAL)  like  @SearchText + '%' order by codpostal"
                'cmd.CommandText = "select (Nombre + ' ' +Email) as Nombre, Email from Usuarios where " & "(Nombre + ' ' +Email) like @SearchText + '%'"
                cmd.Parameters.AddWithValue("@SearchText", prefix)
                cmd.Connection = conn
                conn.Open()
                Using sdr As SqlDataReader = cmd.ExecuteReader()
                    While sdr.Read()
                        customers.Add(String.Format("{0}-{1}", sdr("codpostal"), sdr("codpostal")))
                    End While
                End Using
                conn.Close()
            End Using
            Return customers.ToArray()
        End Using
    End Function
    <WebMethod()> _
    Public Function GetLocalidad(ByVal prefix As String) As String()
        Dim customers As New List(Of String)()
        Using conn As New SqlConnection()
            conn.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Using cmd As New SqlCommand()
                'cmd.CommandText = "select Nombre,Email from Usuarios where " & "Nombre like @SearchText + '%'"
                cmd.CommandText = "SELECT TOP (10) Poblacion AS nombre FROM CODPOSTAL GROUP BY CODPOSTAL.Poblacion having (CODPOSTAL.Poblacion LIKE   @SearchText + '%') ORDER BY CODPOSTAL.Poblacion"
                'cmd.CommandText = "select (Nombre + ' ' +Email) as Nombre, Email from Usuarios where " & "(Nombre + ' ' +Email) like @SearchText + '%'"
                cmd.Parameters.AddWithValue("@SearchText", prefix)
                cmd.Connection = conn
                conn.Open()
                Using sdr As SqlDataReader = cmd.ExecuteReader()
                    While sdr.Read()
                        customers.Add(String.Format("{0}-{1}", sdr("nombre"), sdr("nombre")))
                    End While
                End Using
                conn.Close()
            End Using
            Return customers.ToArray()
        End Using
    End Function
    <WebMethod()> _
    Public Function GetLocalidadcp(ByVal prefix As String, ByVal cp As String) As String()
        Dim customers As New List(Of String)()
        Using conn As New SqlConnection()
            conn.ConnectionString = VariablesGlobales.CadenadeConexion(Application)
            Using cmd As New SqlCommand()
                'cmd.CommandText = "select Nombre,Email from Usuarios where " & "Nombre like @SearchText + '%'"
                If cp <> "" Then
                    cmd.CommandText = "SELECT TOP (10) Poblacion AS nombre FROM CODPOSTAL GROUP BY CODPOSTAL.Poblacion,CODPOSTAL.CODPOSTAL having (CODPOSTAL.Poblacion LIKE '%' + @SearchText + '%') and CODPOSTAL.CODPOSTAL = '" & cp & "'  ORDER BY CODPOSTAL.Poblacion"

                Else
                    cmd.CommandText = "SELECT TOP (10) Poblacion AS nombre FROM CODPOSTAL GROUP BY CODPOSTAL.Poblacion having (CODPOSTAL.Poblacion LIKE '%' + @SearchText + '%') ORDER BY CODPOSTAL.Poblacion"

                End If
                'cmd.CommandText = "select (Nombre + ' ' +Email) as Nombre, Email from Usuarios where " & "(Nombre + ' ' +Email) like @SearchText + '%'"
                cmd.Parameters.AddWithValue("@SearchText", prefix)
                cmd.Connection = conn
                conn.Open()
                Using sdr As SqlDataReader = cmd.ExecuteReader()
                    While sdr.Read()
                        customers.Add(String.Format("{0}*{1}", sdr("nombre"), sdr("nombre")))
                    End While
                End Using
                conn.Close()
            End Using
            Return customers.ToArray()
        End Using
    End Function

End Class