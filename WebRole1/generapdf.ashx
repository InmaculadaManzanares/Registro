<%@ WebHandler Language="VB" Class="generapdf" %>
Imports WebRole1
Imports System
Imports System.Web

Public Class generapdf : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim path, ruta As String
        Dim iddelabd, codigoserie, idfichero As Long
        Dim temporal As Integer
        Dim borrador As Integer
        Dim path_temp As String
        If Not Long.TryParse(context.Request.QueryString("BD"), iddelabd) Then
            Exit Sub
        End If
        If Not Long.TryParse(context.Request.QueryString("Serie"), codigoserie) Then
            Exit Sub
        End If
        If Not Long.TryParse(context.Request.QueryString("ID"), idfichero) Then
            Exit Sub
        End If

        If Not context.Request.QueryString("Extension").ToUpper = "PDF" Then
            Exit Sub
        End If


        If Integer.TryParse(context.Request.QueryString("temp"), temporal) Then
            If temporal = 1 Then
                If Integer.TryParse(context.Request.QueryString("borr"), borrador) Then
                    If borrador = 1 Then
                        path_temp = System.Configuration.ConfigurationManager.AppSettings("ContenedorFicheros_temp").ToString & _
                        VariablesGlobales.Usuario(context.Session, context.Application) & "\" & CARPETA_BORRADOR
                        If Not System.IO.Directory.Exists(path_temp) Then
                            System.IO.Directory.CreateDirectory(path_temp)
                        End If
                        path = path_temp
                    Else
                        path_temp = System.Configuration.ConfigurationManager.AppSettings("ContenedorFicheros_temp").ToString & _
                        VariablesGlobales.Usuario(context.Session, context.Application)
                        If Not System.IO.Directory.Exists(path_temp) Then
                            System.IO.Directory.CreateDirectory(path_temp)
                        End If
                        path = path_temp
                    End If
                Else
                    path_temp = System.Configuration.ConfigurationManager.AppSettings("ContenedorFicheros_temp").ToString & _
                    VariablesGlobales.Usuario(context.Session, context.Application)
                    If Not System.IO.Directory.Exists(path_temp) Then
                        System.IO.Directory.CreateDirectory(path_temp)
                    End If
                    path = path_temp
                End If
            Else
                
                'path = System.Configuration.ConfigurationManager.AppSettings("ContenedorFicheros").ToString
                'path = VariablesGlobales.ContenedorFicheros(context.Session)
                Dim con As ConexionSQL
                Dim sen As SentenciaSQL
                Dim cadena As String
                con = New ConexionSQL(VariablesGlobales.CadenadeConexion(context.Application))

                sen = New SentenciaSQL
                With sen
                    .sql_from = TABLA_BD
                    .add_campo_select(CAMPO_ContenedorFicheros)
                    .add_condicion(CAMPO_ID, iddelabd)
                End With
           
                cadena = con.ejecuta1v_string(sen.texto_sql)
                If cadena <> "" Then
                    If Integer.TryParse(context.Request.QueryString("borr"), borrador) Then
                        If borrador = 1 Then
                            path = cadena & CARPETA_BORRADOR
                        Else
                            path = cadena
                        End If
                    Else
                        path = cadena
                    End If
                Else
                    If Integer.TryParse(context.Request.QueryString("borr"), borrador) Then
                        If borrador = 1 Then
                            path = System.Configuration.ConfigurationManager.AppSettings("ContenedorFicheros") & CARPETA_BORRADOR
                        Else
                            path = System.Configuration.ConfigurationManager.AppSettings("ContenedorFicheros")
                        End If
                    Else
                        path = System.Configuration.ConfigurationManager.AppSettings("ContenedorFicheros")
                    End If
                End If
            End If
        Else
            Exit Sub
        End If


            ruta = path & iddelabd.ToString() & "\" & codigoserie.ToString() & "\" & idfichero & "." & context.Request.QueryString("Extension")




            If System.IO.File.Exists(ruta) Then

                With context.Response
                    .ContentType = "octet/stream"
                    '.AppendHeader("Content-Disposition", "attachment; filename=" & idfichero & ".PDF")
                
                    .TransmitFile(ruta)
                    .Flush()
                    If temporal = 1 Then
                
                        System.IO.File.Delete(ruta)
                    End If
                    .End()
                End With
            End If
        
        
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class