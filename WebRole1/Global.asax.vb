Imports System.Web.Optimization

Public Class Global_asax
    Inherits HttpApplication
    Const separador_claves As String = "X"

    Const cte_ofuscacion As String = "..,.,.,.,.,99092,kjlklk.kjkj....."

    'Const cteCLAVE_ESPACELAND As String = "MirlasKnives69.com"
    'cte_clave_ofuscada es Encripta(cteCLAVE_ESPACELAND) con cte_ofuscacion como clave
    Const cte_clave_ofuscada As String = "A4A0C879B896FDEA57A2F60EB76EA935F09C"


    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        'codigo inicial de global
        ' Se desencadena al iniciar la aplicación
        RouteConfig.RegisterRoutes(RouteTable.Routes)
        BundleConfig.RegisterBundles(BundleTable.Bundles)


        ' Código que se ejecuta al iniciarse la aplicación

        leevalor(VariablesGlobales.cteCadenadeConexion)
        leevalor(VariablesGlobales.cteusu_correo)
        leevalor(VariablesGlobales.ctePas_correo)

        leevalor(VariablesGlobales.cteDirectorioBaseAdjuntos)
        leevalor(VariablesGlobales.cteServidorCorreo)
        leevalor(VariablesGlobales.cteAutenticacionAD)
        leevalor(VariablesGlobales.cteNumeroMaximoNodos)
        leevalor(VariablesGlobales.cteNumeroMaximoNodosObjetos)
        leevalor(VariablesGlobales.cteNumeroMaximoResultadosBusqueda)
        leevalor(VariablesGlobales.cteNumeroAvance)
        leevalor(VariablesGlobales.cteNorma)
        leevalor(VariablesGlobales.ctestrAccount)
        leevalor(VariablesGlobales.ctestrKey)
        'configuracion encriptada
        Dim idinstalacion, configcrypt As String

        leevalor(VariablesGlobales.cteIDInstalacion)
        idinstalacion = VariablesGlobales.IdInstalacion(Application)

        configcrypt = System.Configuration.ConfigurationManager.AppSettings("ConfigPRIVATE")


        Dim configxml As String = String.Empty
        Dim config As New configuracionPrivadaDocBox


        If clavesconfiguracionPrivadaDocBox.Desencripta(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), configcrypt, configxml) Then
            VariablesGlobales.NumMaxUsuarios(Application) = config.numeromaximousuarios
        Else
            'Valores 
            VariablesGlobales.NumMaxUsuarios(Application) = -1
        End If

        'para sacar los valores de encriptacion 
        'Dim valor As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "Passw0rd")
        Dim valor1 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "Server=192.168.0.18;Database=RegistroCorielEspaceland;User ID=sa;Password=es1;Pooling=False")
        Dim valor2 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "patronos@fundacionunicaja.com")
        Dim valor3 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "rg0*8pb-hJ")
        Dim valor4 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "patronos@fundacionunicaja.com")
        'Dim valor2 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "campus@obrasocialunicaja.com")
        'Dim valor3 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "Soft2121.")
        'Dim valor3 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "Server=DOCBOX-AZURE;Database=DocBoxMain;User ID=sa;Password=es1;Pooling=False")
        'Dim valor4 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "Server=SPA-DOCBOX;Database=DocBoxMain;User ID=sa;Password=es1;Pooling=False")

        'Dim valor4 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "Server=SPA-DOCBOX;Database=DocBoxMain;User ID=sa;Password=es1;Pooling=False")

        'Dim valor1 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "Server=srvazu01;Database=GesCampus;User ID=sa;Password=es1;Pooling=False")

        'Passw0rd=F6DCAEADB6B3A8E0
        'cadena de conexion = "F5DBA8AEA7E0B4EBBEF102C5EDBFD9F5F7F7E4B0E66BF88E99BF4F86B56BC8924AB2B1CAC98ECEFEA4ACC1E8A988934CA3B8D2F9DEAEDAA7129C8AE7DAE6DEF3BAFEBFE1FCD2E323AE"
        'usuario correo = "D5F1969289DDC68F8DC725F9C7C0FB8DA5A0E9B9FE"
        'password correo = "F5D1B6B2ECE8B3AEC9"
        'cadena de conexion azure = "F5DBA8AEA7E0B4DCB4CE1EECE9A6B0FDEEDDFFF1810C98EAA2A14D82CE3B8FFB00FEEBE6D18CC7BE91A3EF899B92F644B9B0929EB7C6A5922DA2E18A82A2C4BECE8BC2CDFFDEE46AC4C903F6C1"
        'Dim valor3 As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "Server=docbox-SaaS;Database=DocBoxMain;User ID=sa;Password=es1;Pooling=False")
        Dim prueba As String = clavesconfiguracionPrivadaDocBox.Desencripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "C5EF83839ADCF5BFEDBA4C90AEA59DE4D7D8D680D553D2EFA6AB59")
        Dim prueba2 As String = clavesconfiguracionPrivadaDocBox.Desencripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "D5F19C858AD6CB92D18D6CB28E81A6C4F1D0D191CB43C2B2F1FC0DD2CD15ADDB08F8BB98A1F8")
        '
        Dim valor3UsuarioCorreo As String = clavesconfiguracionPrivadaDocBox.encripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "campus@fundacionunicaja.com")

        '  <add key="UsuarioCorreo" value="D5F19C858AD6CB92D18D6CB28E81A6C4F1D0D88FC944D7ABE2E917D1D50AB8CC14EFE38DB8E3AB" />
        '   Dim prueba2UsuarioCorreo As String = clavesconfiguracionPrivadaDocBox.Desencripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "D5F19C858AD6CB92D18D6CB28E81A6C4F1D0D88FC944D7ABE2E917D1D50AB8CC14EFE38DB8E3AB")
        '  <add key="Password_Correo" value="C8BBE0A1BDADB7CD" />
        '   Dim prueba2Password_Correo As String = clavesconfiguracionPrivadaDocBox.Desencripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "C8BBE0A1BDADB7CD")
        '  <add key="Email_WebCampus" value="D5F19C858AD6CB92D18D6CB28E81A6C4F1D0D88FC944D7ABE2E917D1D50AB8CC14EFE38DB8E3AB" />
        '  Dim prueba2Email_WebCampus As String = clavesconfiguracionPrivadaDocBox.Desencripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "D5F19C858AD6CB92D18D6CB28E81A6C4F1D0D88FC944D7ABE2E917D1D50AB8CC14EFE38DB8E3AB")
        '  <add key="Email_WebCampusGestor" value="C5EF83839ADCF5B6F3B84B85B7B688FED4C0C995C24FC5B7B3B24285" />
        '  Dim prueba2Email_WebCampusGestor As String = clavesconfiguracionPrivadaDocBox.Desencripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), "C5EF83839ADCF5B6F3B84B85B7B688FED4C0C995C24FC5B7B3B24285")

        'valor configuracion gescampus = F5DBA8AEA7E0B4EBBEF102C5EDBFD9F5F7F7E4B0E66BF88E99BC469FAD7DCBAC6082D4C1E4B5E280B3B6A2C596ED9D738E95FBC8F294B99C3FF3EFE9EBD7ECC48AC7DCF9C3E0CE18

        leevalor_encry(VariablesGlobales.cteCadenadeConexion, idinstalacion)
        leevalor_encry(VariablesGlobales.cteusu_correo, idinstalacion)
        leevalor_encry(VariablesGlobales.ctePas_correo, idinstalacion)
        leevalor_encry(VariablesGlobales.cteEmail_WebCampus, idinstalacion)
        leevalor_encry(VariablesGlobales.cteEmail_WebCampusGestor, idinstalacion)


    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Código que se ejecuta durante el cierre de aplicaciones

        'Dim sessionCookieKey = Response.Cookies.AllKeys.SingleOrDefault(Function(c) c.ToLower() = "asp.net_sessionid")
        'Dim sessionCookie = Response.Cookies.[Get](sessionCookieKey)
        'If sessionCookie IsNot Nothing Then
        '    sessionCookie.Secure = True
        'End If

    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Código que se ejecuta al producirse un error no controlado



    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Código que se ejecuta cuando finaliza una sesión. 
        ' Nota: El evento Session_End se desencadena sólo con el modo sessionstate
        ' se establece como InProc en el archivo Web.config. Si el modo de sesión se establece como StateServer 
        ' o SQLServer, el evento no se genera.

    End Sub

    Sub leevalor(ByVal etiq As String)
        Dim s As String
        s = System.Configuration.ConfigurationManager.AppSettings(etiq)
        Application(etiq) = s
    End Sub
    Sub leevalor_encry(ByVal etiq As String, idinstalacion As String)
        Dim s As String
        s = System.Configuration.ConfigurationManager.AppSettings(etiq)
        Application(etiq) = clavesconfiguracionPrivadaDocBox.Desencripta_valor(idinstalacion, clavesconfiguracionPrivadaDocBox.ClavePrivada(), s)
    End Sub
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

    Private Function cteCadenadeConexion() As String
        Throw New NotImplementedException
    End Function

End Class