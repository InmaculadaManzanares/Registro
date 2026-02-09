<%@ WebHandler Language="VB" Class="Descarga" %>
Imports WebRole1
Imports System
Imports System.Web

Public Class Descarga : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        'add some logic to validate the user if he has access to download file
       
        Dim file As String = context.Request.QueryString("file")
      
      
        context.Response.ContentType = "application/octet-stream"
        'I have set the ContentType to "application/octet-stream" which cover any type of file
        ' context.Response.AddHeader("content-disposition", "attachment;filename=" & IO.Path.GetFileName(file))
        'context.Response.WriteFile(context.Server.MapPath(file))
        'here you can do some statistic or tracking
        'you can also implement other business request such as delete the file after download
        context.Response.TransmitFile(context.Server.MapPath(file))
        context.Response.Flush()

                
        System.IO.File.Delete(context.Server.MapPath(file))
   
       
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class