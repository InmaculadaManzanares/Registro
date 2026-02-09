Imports Microsoft.VisualBasic

Imports System.Xml.Serialization
Public Module clavesconfiguracionPrivadaDocBox
    Const separador_claves As String = "X"
    Const cte_ofuscacion As String = "..,.,.,.,.,99092,kjlklk.kjkj....."
    Const cte_clave_ofuscada As String = "A4A0C879B896FDEA57A2F60EB76EA935F09C"
    Public Function ClavePrivada() As String
        Dim g As Crypto
        g = New Crypto(cte_ofuscacion)
        Return g.Desencriptar(cte_clave_ofuscada)
    End Function

#Const generarclave = False
#If generarclave Then
    Public Function generaclaveOfuscada() As String
        Const cteCLAVE_ESPACELAND As String = "MirlasKnives69.com"
        Dim g As Crypto
        g = New Crypto(cte_ofuscacion)
        Return g.Encriptar(cteCLAVE_ESPACELAND)
    End Function
#End If

    Public Function LeeObjeto(ByVal kinstalacion As String, ByVal kespaceland As String, ByVal texto As String, ByRef config As configuracionPrivadaDocBox) As Boolean
        Dim plano as string = string.empty
        Dim string_reader As System.IO.StringReader
        Dim xml_serializer As XmlSerializer


        If Not Desencripta(kinstalacion, kespaceland, texto, plano) Then
            Return False
        End If
        Try
            config = New configuracionPrivadaDocBox
            xml_serializer = New XmlSerializer(config.GetType)

            string_reader = New System.IO.StringReader(plano)

            config = _
                DirectCast(xml_serializer.Deserialize(string_reader), _
                    configuracionPrivadaDocBox)
            string_reader.Close()

        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function
    Public Function encripta(ByVal claveinstalacion As String, ByVal claveespaceland As String, ByVal texto As String) As String
        Dim gen As New Crypto(claveinstalacion & separador_claves & claveespaceland)

        Dim ta As String
        ta = Alea.CadenaAleatoria(Alea.aleatorio_en_rango(0, 128)) & texto
        Return gen.Encriptar(ta)
    End Function
    Public Function encripta_valor(ByVal claveinstalacion As String, ByVal claveespaceland As String, ByVal texto As String) As String
        Dim gen As New Crypto(claveinstalacion & separador_claves & claveespaceland)

        Dim ta As String
        ta = texto
        Return gen.Encriptar(ta)
    End Function
    Public Function Desencripta(ByVal claveinstalacion As String, ByVal claveespaceland As String, ByVal texto As String, ByRef resultado As String) As Boolean
        Dim gen As New Crypto(claveinstalacion & separador_claves & claveespaceland)
        Dim r As String
        Dim i As Integer

        r = gen.Desencriptar(texto)
        'buscamos el xml 
        i = r.IndexOf("<?xml")
        If i < 0 Then
            'no encontrado
            Return False
        Else
            resultado = r.Substring(i)
        End If
        Return True
    End Function
    Public Function Desencripta_valor(ByVal claveinstalacion As String, ByVal claveespaceland As String, ByVal texto As String) As String
        Dim gen As New Crypto(claveinstalacion & separador_claves & claveespaceland)
        Dim r As String
        Dim i As Integer

        r = gen.Desencriptar(texto)

        Return r.Substring(i)

    End Function
End Module

Public Class configuracionPrivadaDocBox


    Public numeromaximousuarios As Integer


End Class