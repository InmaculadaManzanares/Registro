Imports ICSharpCode.SharpZipLib.Core
Imports ICSharpCode.SharpZipLib.Zip
Imports System.io

Public Module util_icsharpziplib

    Public Sub Descomprimir(ByVal nombreficherozip As String, ByVal directorio As String)
        Dim strzipin As ZipInputStream
        Dim strmFile As System.IO.FileStream
        Dim abyBuffer() As Byte
        Dim objZipEntry As ZipEntry
        Dim ficheroatratar As String

        strzipin = New ZipInputStream(System.IO.File.Open(nombreficherozip, System.IO.FileMode.Open))
        objZipEntry = strzipin.GetNextEntry()

        While Not (objZipEntry Is Nothing)
            ReDim abyBuffer(CInt(objZipEntry.Size))

            strzipin.Read(abyBuffer, CInt(objZipEntry.Offset), CInt(objZipEntry.Size))

            ficheroatratar = directorio & objZipEntry.Name
            Util.BorraficheroSiExiste(ficheroatratar)
            strmFile = New System.IO.FileStream(directorio & objZipEntry.Name, System.IO.FileMode.Create)

            strmFile.Write(abyBuffer, 0, CInt(objZipEntry.Size))
            strmFile.Close()
            objZipEntry = strzipin.GetNextEntry()
        End While
        strzipin.Close()


    End Sub

    Public Sub AñadeFicheroaZip(ByVal nombrefichero As String, ByVal nombreamostrar As String, ByVal z As ZipOutputStream)
        Dim strmFile As FileStream
        Dim abyBuffer() As Byte
        Dim objZipEntry As ZipEntry

        strmFile = File.OpenRead(nombrefichero)
        ReDim abyBuffer(CInt(strmFile.Length - 1))

        strmFile.Read(abyBuffer, 0, abyBuffer.Length)

        objZipEntry = New ZipEntry(nombreamostrar)
        With objZipEntry

            .DateTime = DateTime.Now
            .Size = strmFile.Length
        End With

        strmFile.Close()

        z.PutNextEntry(objZipEntry)
        z.Write(abyBuffer, 0, abyBuffer.Length)
    End Sub


End Module