Imports System.Drawing
Imports System.Drawing.imaging
Imports System.io
Public Module Overlay
    'C:\desarrollo\proyectos\desarrollo\PruebasOverlay

    Private espaciohorizontal As Integer

    Public Function SobreponerPlantilla(ByVal rutaplantilla As String, ByVal rutafichero As String, ByVal offsetx As Integer, ByVal offsety As Integer, ByVal tamfuente As Integer, ByVal mult_espacio_vertical As Single, ByVal mult_espacio_horizontal As Single, ByVal negrita As Boolean) As Image

        Dim objgraphic As Graphics
        Dim texto As String
        Dim fuente As Font
        Dim brocha As Brush = New SolidBrush(Color.Black)
        Dim f As System.IO.StreamReader
        Dim bm_original, bm_noindizado As Bitmap
        Dim nombrefuente As String = "Courier New"
        Try
            fuente = ObtenerFuente(nombrefuente, tamfuente, negrita)

            If Not System.IO.File.Exists(rutaplantilla) Then
                Throw New Exception("No se encuentra el fichero: " & rutaplantilla)
            End If

            If Not System.IO.File.Exists(rutafichero) Then
                Throw New Exception("No se encuentra el fichero: " & rutafichero)
            End If

            bm_original = New Bitmap(rutaplantilla)
            bm_noindizado = New Bitmap(bm_original.Width, bm_original.Height, PixelFormat.Format24bppRgb)

            objgraphic = Graphics.FromImage(bm_noindizado)
            objgraphic.DrawImage(bm_original, New Point(0, 0))
            bm_original.Dispose()




            f = New StreamReader(rutafichero, System.Text.Encoding.Default)
            While Not f.EndOfStream
                texto = f.ReadLine()
                Imprimetexto(objgraphic, texto, fuente, brocha, offsetx, offsety, mult_espacio_vertical, mult_espacio_horizontal)
                'fuenteKerning.DrawString(objgraphic, texto, offsetx, offsety, Color.Black, Color.White)
                'offsety += fuente.Height * mult_espacio_vertical
                'objgraphic.DrawString(texto, fuente, brocha, offsetx, offsety)
                'offsety += fuente.Height * mult_espacio_vertical
            End While
            f.Close()

            objgraphic.Dispose()
            objgraphic = Nothing
        Catch ex As Exception
            Throw New Exception("Error sobreponiendo Plantilla: " & ex.Message)
        End Try
        Return bm_noindizado


    End Function

    Private Sub Imprimetexto(ByVal g As Graphics, ByVal texto As String, ByVal fuente As Font, ByVal brocha As Brush, ByVal offsetx As Integer, ByRef offsety As Integer, ByVal mult_ver As Single, ByVal mult_hor As Single)
        Dim x As Integer

        x = offsetx
        For Each c As String In texto
            g.DrawString(c, fuente, brocha, x, offsety)
            x = CInt(x + espaciohorizontal * mult_hor)
        Next


        offsety = CInt(offsety + fuente.Height * mult_ver)

        'para comparar con el espaciado normal
        'g.DrawString(texto, fuente, brocha, offsetx, offsety)
        'offsety += fuente.Height * mult_ver
    End Sub


    Private Function ObtenerFuente(ByVal nombrefuente As String, ByVal tamfuente As Integer, ByVal negrita As Boolean) As Font
        Dim fuente As Font

        Try
            If negrita Then
                fuente = New Font(nombrefuente, tamfuente, FontStyle.Bold)
            Else
                fuente = New Font(nombrefuente, tamfuente)
            End If


        Catch ex As Exception
            Throw New Exception("No es posible encontrar la fuente " & nombrefuente)
        End Try


        'Ahora estimamos el ancho de los caracteres y la separación entre caracteres


        espaciohorizontal = estimaespacio(fuente)

        '  espaciohorizontal = estimadoranchofijo(fuente)



        Return fuente
    End Function




    Private Function estimaespacio(ByVal fuente As Font) As Integer

        Dim e1, e2, e3, e4 As Single
        Const cad1 As String = "A"
        Const cad2 As String = "i"
        Const cad3 As String = "0"
        Const cad4 As String = "."

        e1 = CSng(estimacadena(fuente, cad1))
        e2 = CSng(estimacadena(fuente, cad2))
        e3 = CSng(estimacadena(fuente, cad3))
        e4 = CSng(estimacadena(fuente, cad4))
        Dim val As Integer
        val = CInt((e1 + e2 + e3 + e4))
        Return val \ 4

    End Function
    Private Function estimacadena(ByVal fuente As Font, ByVal cadena As String) As Long
        Dim normal, doble As SizeF
        Dim bmp As Bitmap
        Dim g As Graphics
        Dim tam4 As Integer = CInt(fuente.SizeInPoints * 4)

        bmp = New Bitmap(tam4, tam4, PixelFormat.Format24bppRgb)
        g = Graphics.FromImage(bmp)
        g.PageUnit = GraphicsUnit.Point

        normal = g.MeasureString(cadena, fuente)
        doble = g.MeasureString(cadena & cadena, fuente)

        'Normal, es el ancho del caracter A
        'doble, es el ancho del caracter A + espaciohorizontal + el ancho del caracter A

        Return (doble.ToSize.Width - normal.ToSize.Width)
    End Function
End Module
