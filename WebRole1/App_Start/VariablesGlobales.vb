Imports Microsoft.VisualBasic
Imports System.Web.SessionState.HttpSessionState

Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls

Public Module CteTablas

    'Public Const TABLA_ROL As String = "G_ROL"
    'Public Const CAMPO_IDROL As String = "IDRol"
    'Public Const CAMPO_NOMBRE As String = "NOMBRE"

    Public Const COLONIA_BECA_USA As Integer = 5

    Public Const TABLA_USUARIO As String = "USUARIOS"
    Public Const CAMPO_CODUSUARIO As String = "CodUsuario"
    Public Const CAMPO_CODPERFIL As String = "CodPerfil"
    Public Const CAMPO_PASSWORD As String = "PassWord"
    Public Const CAMPO_PASSWORD_ENC As String = "CONVERT(VARCHAR(300), DECRYPTBYPASSPHRASE('" & Util.sSecretKey & "',Password))"
    Public Const CAMPO_EMAIL As String = "Email"

    Public Const TABLA_ROLUSRBD As String = "G_ROLUSRBD"
    Public Const CAMPO_ADMIN As String = "Admin"
    Public Const CAMPO_SISTEMA As String = "Sistema"
    Public Const CAMPO_IDBD As String = "IDBd"

    Public Const TABLA_GRUPOBD As String = "G_GRUPOBD"
    Public Const CAMPO_IDGRUPO As String = "IDGrupo"

    Public Const TABLA_USUARIOGRUPO As String = "G_USRGRUPO"

    Public Const CAMPO_ID As String = "ID"
    Public Const CAMPO_TODOS As String = "*"
    'Public Const CAMPO_CODPERFIL As String = "CodPerfil"
    Public Const TABLA_GRUPO As String = "G_GRUPO"

    Public Const TABLA_PERFIL As String = "PERFILES"
    Public Const CAMPO_NOMBRE As String = "Nombre"
    Public Const CAMPO_OBJETO As String = "Objeto"
    Public Const CAMPO_METADATOS As String = "Metadatos"

    Public Const CAMPO_CAMPOSTAMP As String = "CampoStamp"
    Public Const CAMPO_VALOR1STAMP As String = "ValorCampoStamp1"
    Public Const CAMPO_FICHERO1STAMP As String = "FicheroStamp1"
    Public Const CAMPO_VALOR2STAMP As String = "ValorCampoStamp2"
    Public Const CAMPO_FICHERO2STAMP As String = "FicheroStamp2"
    Public Const CAMPO_TITULO As String = "Titulo"

    Public Const TABLA_BD As String = "G_BD"
    Public Const CAMPO_DESCRIPCION As String = "Descripcion"
    Public Const CAMPO_UBICACION As String = "Ubicacion"
    Public Const CAMPO_ExtensionesAFullText As String = "ExtensionesAFullText"
    Public Const CAMPO_ExtensionesAOcr As String = "ExtensionesAOcr"
    Public Const CAMPO_ExtensionesANoComprimir As String = "ExtensionesANoComprimir"
    Public Const CAMPO_AplicacionEscaner As String = "AplicacionEscaner"
    Public Const CAMPO_ContenedorFicheros As String = "ContenedorFicheros"
    Public Const CAMPO_FiltrarComboCarpeta As String = "FiltrarComboCarpeta"
    Public Const CAMPO_UrlSubidaPlantilla As String = "UrlSubidaPlantilla"
    Public Const CAMPO_ContenedorPlantilla As String = "ContenedorPlantillas"
    Public Const CAMPO_ContenedorDescarga As String = "ContenedorDescargas"
    Public Const CAMPO_ConservacionHistorico As String = "ConservacionHistorico"
    Public Const CAMPO_BusquedaCampos As String = "BusquedaCampos"

    Public Const TABLA_COBJETO As String = "DB_COBJETO"
    Public Const TABLA_LOBJETO As String = "DB_LOBJETO"
    Public Const TABLA_DATOS_OBJETO As String = "DB_DATOS_OBJETO"
    Public Const TABLA_DATOS_COBJETO As String = "DB_DATOS_COBJETO"
    Public Const CAMPO_IDOBJETO As String = "IdObjeto"
    Public Const CAMPO_IDLOBJETO As String = "IdLObjeto"
    Public Const CAMPO_METAOBJETOS As String = "MetaObjetos"

    Public Const TABLA_CSERIE As String = "DB_CSERIE"
    Public Const TABLA_LSERIE As String = "DB_LSERIE"
    Public Const CAMPO_TIPOSQL As String = "TipoSql"
    Public Const CAMPO_VFORM As String = "VForm"
    Public Const CAMPO_TIPOCON As String = "TipoCon"
    Public Const CAMPO_TAMAÑO As String = "Tamano"
    Public Const CAMPO_ORDEN As String = "Orden"
    Public Const CAMPO_OBLIGATORIO As String = "Obligat"
    Public Const CAMPO_VALIDA As String = "Valida"
    Public Const CAMPO_SELECTOR As String = "Selector"

    Public Const TABLA_CUADRO As String = "DB_CUADRO"
    Public Const CAMPO_IDPADRE As String = "IdPadre"
    Public Const CAMPO_IDSERIE As String = "IdSerie"
    Public Const CAMPO_CARPETA As String = "Carpeta"
    Public Const CAMPO_NHIJOS As String = "NHijos"
    Public Const CAMPO_PLANTILLA As String = "Plantilla"
    Public Const CAMPO_IDNORMA_DEST As String = "IdNorma_dest"
    Public Const CAMPO_FECHA_DEST As String = "campofecha_dest"

    Public Const TABLA_BORRADOR_CUADRO As String = "DB_BORRADOR_CUADRO"
    Public Const TABLA_BORRADOR_DATOS As String = "DB_BORRADOR_DATOS"
    Public Const TABLA_BORRADOR_PERMISOS As String = "DB_BORRADOR_PERMISOS"
    Public Const TABLA_BORRADOR_FICHEROS As String = "DB_BORRADOR_FICHEROS"

    Public Const TABLA_CDATOS As String = "DB_CDATOS"
    Public Const CAMPO_IDCUADRO As String = "IDCuadro"

    Public Const TABLA_ACCION_REG As String = "DB_ACCION_REG"
    Public Const CAMPO_ACCION As String = "Accion"
    Public Const CAMPO_FECHAREG As String = "FechaRegistro"
    Public Const CAMPO_USUARIOREG As String = "UsuarioRegistro"
    Public Const TABLA_ACCION_OBJETO_REG As String = "DB_ACCION_OBJETO_REG"

    Public Const TABLA_DATOS_REG As String = "DB_DATOS_REG"
    Public Const TABLA_DATOS_OBJETO_REG As String = "DB_DATOS_OBJETO_REG"

    Public Const TABLA_DATOS As String = "DB_DATOS"
    Public Const CAMPO_IDLSERIE As String = "IdLSerie"
    Public Const CAMPO_VALOR As String = "valor"
    Public Const CAMPO_FILE As String = "FileContent"

    Public Const TABLA_PERMISOS As String = "DB_PERMISOS"
    Public Const TABLA_PERMISOS_CARPETAS As String = "DB_PERMISOS_CARPETAS"
    Public Const TABLA_PERMISOS_OBJETO As String = "DB_PERMISOS_OBJETO"
    Public Const TABLA_PERMISOS_IDOBJETO As String = "DB_PERMISOS_IDOBJETO"
    Public Const CAMPO_VER As String = "Ver"
    Public Const CAMPO_AÑADIR As String = "Añadir"
    Public Const CAMPO_MODICIAR As String = "Modificar"
    Public Const CAMPO_BORRAR As String = "Borrar"
    Public Const CAMPO_IMPRIMIR As String = "Imprimir"
    Public Const CAMPO_HISTORICO As String = "Historico"
    Public Const CAMPO_SEGURIDAD As String = "Seguridad"

    Public Const TABLA_FICHEROS As String = "DB_FICHEROS"
    Public Const CAMPO_NOMBREFICHERO As String = "NombreFichero"
    Public Const CAMPO_NOMBREORIGINAL As String = "NombreOriginal"
    Public Const CAMPO_FICHEROBINARY As String = "FileContent"
    Public Const CAMPO_TIPOFICHERO As String = "FileType"

    Public Const TABLA_BUSQUEDA As String = "DB_BUSQUEDA"
    Public Const TABLA_CONSULTA As String = "DB_CONSULTA"
    Public Const CAMPO_CONSULTA As String = "Consulta"
    Public Const CAMPO_serie As String = "Serie"
    Public Const CAMPO_referencia As String = "Referencia"
    Public Const CAMPO_IDENTIFICADOR As String = "Identificador"
    Public Const CAMPO_ComboCarpeta As String = "ComboCarpeta"
    Public Const CAMPO_SignoCarpeta As String = "SignoCarpeta"
    Public Const CAMPO_ValorCarpeta As String = "ValorCarpeta"
    Public Const CAMPO_CampoDato1 As String = "CampoDato1"
    Public Const CAMPO_SignoDato1 As String = "SignoDato1"
    Public Const CAMPO_ValorDato1 As String = "ValorDato1"
    Public Const CAMPO_AndOr1 As String = "AndOr1"
    Public Const CAMPO_CampoDato2 As String = "CampoDato2"
    Public Const CAMPO_SignoDato2 As String = "SignoDato2"
    Public Const CAMPO_ValorDato2 As String = "ValorDato2"
    Public Const CAMPO_AndOr2 As String = "AndOr2"
    Public Const CAMPO_CampoDato3 As String = "CampoDato3"
    Public Const CAMPO_SignoDato3 As String = "SignoDato3"
    Public Const CAMPO_ValorDato3 As String = "ValorDato3"
    Public Const CAMPO_AndOr3 As String = "AndOr3"
    Public Const CAMPO_CampoDato4 As String = "CampoDato4"
    Public Const CAMPO_SignoDato4 As String = "SignoDato4"
    Public Const CAMPO_ValorDato4 As String = "ValorDato4"
    Public Const CAMPO_TextoCompleto As String = "TextoCompleto"
    Public Const CAMPO_CamposSerie As String = "CamposSerie"
    Public Const TABLA_PLANTILLA As String = "DB_PLANTILLA"
    Public Const CAMPO_FICHERO As String = "Fichero"
    Public Const CAMPO_POSX As String = "PosX"
    Public Const CAMPO_POSY As String = "PosY"
    Public Const CAMPO_ESPACIADOH As String = "EspaciadoH"
    Public Const CAMPO_ESPACIADOV As String = "EspaciadoV"
    Public Const CAMPO_TAMFONT As String = "TamFont"
    Public Const CAMPO_NEGRITA As String = "Negrita"

    Public Const TABLA_NORMA_DEST As String = "DB_NORMA_DEST"
    Public Const CAMPO_MESES As String = "meses"
    Public Const CAMPO_TIPO As String = "tipo"

    Public Const TABLA_LOGDESTRUCCION As String = "DB_LOGDESTRUCCION"
    Public Const CAMPO_FECHA As String = "Fecha"
    Public Const CAMPO_USUARIO As String = "Usuario"
    Public Const CAMPO_IDEJECUCION_NORMA As String = "IDEJECUCIONNORMA"

    Public Const TABLA_EJECNORMA As String = "DB_EJEC_NORMA_DEST"
    Public Const CAMPO_IDNORMA As String = "Idnorma"
    Public Const CAMPO_PROCESADO As String = "Procesado"
    Public Const CAMPO_FECHAPROCESADO As String = "FechaProcesado"



    Public Const Permiso_Administrador As String = "1"
    Public Const Permiso_Sistema As String = "1"
    Public Const Permiso_Usuario As String = "0"

    Public Const ID_DocBoxMain As Long = 1

    Public Const prefijo_bd As String = "DocBox_"
    Public Const prefijo_hbd As String = "HDocBox_"
    Public Const prefijo_historico As String = "H"
    Public Const imagen_carpeta As String = "~/ImagenesDocbox/carpeta.png"
    Public Const imagen_archivo As String = "~/ImagenesDocbox/doc.png"
    Public Const imagen_carpeta_inicio As String = "~/ImagenesDocbox/carpeta.png"
    Public Const imagen_mas_nodos As String = "~/Imagenes/icono_mas_nodos.jpg"
    Public Const imagen_menos_nodos As String = "~/Imagenes/icono_mas_nodos.jpg"
    Public Const imagen_hijos_nodos As String = "~/Imagenes/icono_hijos_nodos.jpg"
    Public Const TEXTO_NODO_MAS_NODOS As String = "(...mas...)"
    Public Const LETRA_BASE As String = "Arial"
    Public Const SIZE_BASE As Integer = 8
    Public Const SIZE_BASE_objeto As Integer = 10
    Public Const ancho_maximo_txt As Integer = 700
    Public Const alto_maximo_txt As Integer = 25
    Public Const alto_default As Integer = 12 + 1
    Public Const alto_default_objeto As Integer = 15
    Public Const alto_default_combo As Integer = 26
    Public Const SELECT_VACIO As String = SQL_SELECT & "'-1'" & COMA & "''"
    Public Const SELECT_VACIO_COMBO As String = SQL_SELECT & "'' AS nombre "


    Public Const UDF_PERMISOVER_CARPETA As String = "dbo.udf_PermisoVer_Carpeta"
    Public Const UDF_PERMISOS_CARPETAS As String = "dbo.udf_Permiso_carpeta"
    Public Const UDF_PERMISOS As String = "dbo.udf_Permiso"
    Public Const UDF_PERMISOS_OBJETO As String = "dbo.udf_Permiso_objeto"
    Public Const UDF_PERMISOS_idOBJETO As String = "dbo.udf_Permiso_idobjeto"
    Public Const F_GENERAARBOL As String = "dbo.f_GeneraArbol" '"dbo.f_GeneraArbol"
    Public Const F_GENERAARBOL_OBJETO As String = "dbo.f_GeneraArbol_Objeto"
    Public Const F_GENERAARBOL_MAS As String = "dbo.f_GeneraArbol_mas"
    Public Const F_GENERAARBOL_MENOS As String = "dbo.f_GeneraArbol_menos"
    Public Const F_GENERAARBOL_HIJOS As String = "dbo.f_GeneraArbol_hijos"

    'Public Const F_GENERAARBOL_MAS_OBJETO As String = "dbo.f_GeneraArbol_mas_objeto"
    'Public Const F_GENERAARBOL_MENOS_OBJETO As String = "dbo.f_GeneraArbol_menos_objeto"
    'Public Const F_GENERAARBOL_HIJOS_OBJETO As String = "dbo.f_GeneraArbol_hijos_objeto"
    Public Const F_GENERAARBOL_CARPETAS As String = "dbo.f_GeneraArbol_carpetas" '"dbo.f_GeneraArbol"
    Public Const F_GENERAARBOL_MAS_CARPETAS As String = "dbo.f_GeneraArbol_mas_carpetas"
    Public Const F_GENERAARBOL_MENOS_CARPETAS As String = "dbo.f_GeneraArbol_menos_carpetas"
    Public Const F_GENERAARBOL_HIJOS_CARPETAS As String = "dbo.f_GeneraArbol_hijos_carpetas"
    Public Const BUSCADATO As String = "dbo.udf_BuscarDato"

    Public Const PERMISO_AÑADIR As String = "Añadir"
    Public Const PERMISO_BORRAR As String = "Borrar"
    Public Const PERMISO_SEGURIDAD As String = "Seguridad"
    Public Const PERMISO_IMPRIMIR As String = "Imprimir"
    Public Const PERMISO_HISTORICO As String = "Historico"
    Public Const PERMISO_MODIFICAR As String = "Modificar"
    Public Const PERMISO_VER As String = "Ver"
    Public Const TENGO_PERMISOS As String = "1"

    Public Const POSICION_SUBCARPETA As String = "Subcarpeta"
    Public Const POSICION_SUBDOCUMENTO As String = "SubDocumento"
    Public Const POSICION_ARRIBA As String = "Arriba"
    Public Const POSICION_ABAJO As String = "Abajo"

    Public Const EXTENSION_COMPRESION As String = ".ILLO"
    Public Const EXTENSION_OCR As String = ".TXT"

    Public Const ACCION_IMPRIMIR As String = "Imprimir"
    Public Const ACCION_AÑADIRCARPETA As String = "Añadir Carpeta"
    Public Const ACCION_MODIFICARCAPERTA As String = "Modificar Carpeta"
    Public Const ACCION_AÑADIRDOCUMENTO As String = "Añadir Documento"
    Public Const ACCION_MODIFICARDOCUMENTO As String = "Modificar Documento"
    Public Const ACCION_PERMISOS As String = "Permisos"
    Public Const ACCION_BORRARDOCUMENTO As String = "Borrar Documento"
    Public Const ACCION_VER As String = "Ver"
    Public Const ACCION_DESCARGAR As String = "Descargar"
    Public Const ACCION_HISTORICO As String = "Historial"
    Public Const ACCION_EMAIL As String = "Email"
    Public Const ACCION_MODIFICAROBJETO As String = "Modificar Objeto"
    Public Const ACCION_AÑADIROBJETO As String = "Añadir Objeto"
    Public Const ACCION_VER_BORRADOR As String = "Ver Borrador"
    Public Const ACCION_HISTORICO_BORRADOR As String = "Historial Borrador"
    Public Const ACCION_RECUPERAR As String = "Recuperar Documento"
    Public Const ACCION_CAMBIO_EXPEDIENTE As String = "Cambio Expediente del "

    Public Const ACCION_PETICION_CAMBIO_TURNO As String = "Petición cambio Turno"
    Public Const ACCION_CREAR_SOLICITUD As String = "Crear Solicitud"
    Public Const ACCION_VER_SOLICITUD As String = "Ver Solicitud"
    Public Const ACCION_MODIFICAR_SOLICITUD As String = "Modificar Solicitud"
    Public Const ACCION_CAMBIAR_CONTRASEÑA As String = "Cambiar Contraseña"
    Public Const ACCION_CANCELAR_SOLICITUD As String = "Cancelar Solicitud"
    Public Const ACCION_DESCARGAR_CARTA As String = "Descargar Carta"
    Public Const ACCION_RECUPERAR_CONTRASEÑA As String = "Recuperar Contraseña"
    Public Const ACCION_MODIFICAR_MISDATOS As String = "Modificar Mis Datos"
    Public Const ACCION_ASIGNACION_PLAZA As String = "Asignación Plaza"
    Public Const ACCION_LISTAESPERA As String = "Lista de Espera"
    Public Const ACCION_DESASIGNACION_PLAZA As String = "Desasignación Plaza"
    Public Const ACCION_CREAR_USUARIO As String = "Crear Usuario"

    Public Const ACCION_ENVIOCARTA_ASIGNACION As String = "Email Asignación Plaza"
    Public Const ACCION_ENVIOCARTA_LISTAESPERA As String = "Email Lista de Espera"
    Public Const ACCION_ENVIO_EMAIL_MANUAL As String = "Email Manual"
    Public Const ACCION_EMAIL_PETICION_CAMBIO_TURNO As String = "Email al Gestor Petición cambio Turno"
    Public Const ACCION_ENVIOEMAIL_CREAR_SOLICITUD As String = "Email Crear Solicitud"
    Public Const ACCION_EMAIL_CREAR_USUARIO As String = "Email Usuario creado"
    Public Const ACCION_EMAIL_CANCELAR_SOLICITUD As String = "Email al Gestor Cancelar Solicitud"
    Public Const ACCION_EMAIL_RECUPERAR_CONTRASEÑA As String = "Email Recuperar Contraseña"
    Public Const ACCION_EMAIL_REGISTRO As String = "Email Registro"

    Public Const ACCION_Modifica_ALIMENTACION As String = "Modifica Alimentación "
    Public Const ACCION_Añade_ALIMENTACION As String = "Añade Alimentación "
    Public Const ACCION_Borra_ALIMENTACION As String = "Borra Alimentación "
    Public Const ACCION_ver_ALIMENTACION As String = "Ver Alimentación "
    Public Const ACCION_Modifica_MEDICACION As String = "Modifica Mediación "
    Public Const ACCION_Añade_MEDICACION As String = "Añade Medicación "
    Public Const ACCION_Borra_MEDICACION As String = "Borra Medicación "
    Public Const ACCION_ver_MEDICACION As String = "Ver Medicación "
    Public Const ACCION_Modifica_FICHAMEDICA As String = "Modifica Ficha Médica "
    Public Const ACCION_Añade_FICHAMEDICA As String = "Añade Ficha Médica "
    Public Const ACCION_Borra_FICHAMEDICA As String = "Borra Ficha Médica "
    Public Const ACCION_Ver_FICHAMEDICA As String = "Ver Ficha Médica "




    Public Const PARAMETRO_IMPRESION_filtro As String = "Filtro"
    Public Const PARAMETRO_IMPRESION_id As String = "id"

    Public Const PARAMETRO_IMPRESION_operador As String = "operador"

    Public Const PARAMETRO_IMPRESION_filtro1 As String = "Filtro1"
    Public Const PARAMETRO_IMPRESION_id1 As String = "id1"

    Public Const PARAMETRO_IMPRESION_NombreInforme As String = "NombreInforme"
    Public Const NOMBRE_INFORME_HISTORICO As String = "Inf_HISTORICO"
    Public Const NOMBRE_INFORME_HISTORICO_DATOS As String = "Inf_HISTORICO_DATOS"
    Public Const NOMBRE_INFORME_HISTORICO_BORRADOR_DATOS As String = "Inf_HISTORICO_BORRADOR_DATOS"
    Public Const NOMBRE_INFORME_HISTORICO_DATOS_OBJETO As String = "Inf_HISTORICO_DATOS_OBJETO"
    Public Const NOMBRE_INFORME_RECUENTO As String = "Inf_Recuento"

    Public Const CARPETA_BORRADOR As String = "Borrador\"

    Public MensajeImagenResize As String
End Module

Public Class VariablesGlobales
    Public Const cteCadenadeConexion = "CadenaConexion"
    Public Shared ReadOnly Property CadenadeConexion(ByVal a As HttpApplicationState) As String
        Get
            ' Public Const CadenadeConexion As String = "Server=srvsoft2;Database=Calidad;User ID=sa;Password=es1;Pooling=False"
            Return a(cteCadenadeConexion)
        End Get
    End Property
    Public Const ctestrAccount = "strAccount"
    Public Shared ReadOnly Property strAccount(ByVal a As HttpApplicationState) As String
        Get
            Return a(ctestrAccount)
        End Get
    End Property
    Public Const ctestrKey = "strKey"
    Public Shared ReadOnly Property strKey(ByVal a As HttpApplicationState) As String
        Get
            Return a(ctestrKey)
        End Get
    End Property
    Private Const cteCadenaConexionDocBox As String = "CadenaConexionDocBox"
    Public Shared Property CadenaConexionDocBox(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteCadenaConexionDocBox))
        End Get
        Set(ByVal value As String)
            s(cteCadenaConexionDocBox) = value
        End Set
    End Property
    Private Const cteExtensionesAFullText As String = "ExtensionesAFullText"
    Public Shared Property ExtensionesAFullText(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteExtensionesAFullText))
        End Get
        Set(ByVal value As String)
            s(cteExtensionesAFullText) = value
        End Set
    End Property
    Private Const cteExtensionesAOcr As String = "ExtensionesAOcr"

    Public Shared Property ExtensionesAOcr(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteExtensionesAOcr))
        End Get
        Set(ByVal value As String)
            s(cteExtensionesAOcr) = value
        End Set
    End Property
    Private Const cteExtensionesANoComprimir As String = "ExtensionesANoComprimir"
    Public Shared Property ExtensionesANoComprimir(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteExtensionesANoComprimir))
        End Get
        Set(ByVal value As String)
            s(cteExtensionesANoComprimir) = value
        End Set
    End Property
    Private Const cteAplicacionEscaner As String = "AplicacionEscaner"
    Public Shared Property AplicacionEscaner(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteAplicacionEscaner))
        End Get
        Set(ByVal value As String)
            s(cteAplicacionEscaner) = value
        End Set
    End Property
    Private Const cteContenedorFicheros As String = "ContenedorFicheros"
    Public Shared Property ContenedorFicheros(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteContenedorFicheros))
        End Get
        Set(ByVal value As String)
            s(cteContenedorFicheros) = value
        End Set
    End Property
    Private Const cteContenedorPlantilla As String = "ContenedorPlantilla"
    Public Shared Property ContenedorPlantilla(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteContenedorPlantilla))
        End Get
        Set(ByVal value As String)
            s(cteContenedorPlantilla) = value
        End Set
    End Property
    Private Const cteContenedorDescarga As String = "ContenedorDescarga"
    Public Shared Property ContenedorDescarga(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteContenedorDescarga))
        End Get
        Set(ByVal value As String)
            s(cteContenedorDescarga) = value
        End Set
    End Property
    Private Const cteFiltrarComboCarpetas As String = "FiltrarComboCarpetas"
    Public Shared Property FiltrarComboCarpetas(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteFiltrarComboCarpetas))
        End Get
        Set(ByVal value As String)
            s(cteFiltrarComboCarpetas) = value
        End Set
    End Property

    Public Const cteDirectorioBaseAdjuntos = "DirectorioBaseAdjuntos"
    Public Shared ReadOnly Property DirectorioBaseAdjuntos(ByVal a As HttpApplicationState) As String
        Get
            Return a(cteDirectorioBaseAdjuntos)
        End Get
    End Property
    'datos de correo
    Public Const cteServidorCorreo = "ServidorCorreo"
    Public Shared ReadOnly Property Servidor_Correo(ByVal a As HttpApplicationState) As String
        Get

            Return a(cteServidorCorreo)
        End Get
    End Property

    Public Const cteusu_correo = "UsuarioCorreo"
    Public Shared ReadOnly Property usu_correo(ByVal a As HttpApplicationState) As String
        Get

            Return a(cteusu_correo)
        End Get
    End Property
    Public Const ctePas_correo = "Password_Correo"
    Public Shared ReadOnly Property Pas_Correo(ByVal a As HttpApplicationState) As String
        Get

            Return a(ctePas_correo)
        End Get
    End Property

    Public Const cteEmail_WebCampus = "Email_WebCampus"
    Public Shared ReadOnly Property Email_WebCampus(ByVal a As HttpApplicationState) As String
        Get

            Return a(cteEmail_WebCampus)
        End Get
    End Property
    Public Const cteEmail_WebCampusGestor = "Email_WebCampusGestor"
    Public Shared ReadOnly Property Email_WebCampusGestor(ByVal a As HttpApplicationState) As String
        Get

            Return a(cteEmail_WebCampusGestor)
        End Get
    End Property

    Public Const cteAutenticacionAD = "AutenticacionAD"
    Public Shared ReadOnly Property AutenticacionAD(ByVal a As HttpApplicationState) As String
        Get
            Return a(cteAutenticacionAD)
        End Get
    End Property

    Private Const cteCAMPOSActuales As String = "CalidadCamposActualesSelector"
    Public Shared Property CamposActuales(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return s(idcontrol & cteCAMPOSActuales)
        End Get
        Set(ByVal value As String)
            s(cteCAMPOSActuales) = value
        End Set
    End Property
    Private Const cteValorTabla As String = "CalidadTablaActualSelector"
    Public Shared Property ValorTabla(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return s(idcontrol & cteValorTabla)
        End Get
        Set(ByVal value As String)
            s(idcontrol & cteValorTabla) = value
        End Set
    End Property
    Private Const ctePathDocumento As String = "CalidadPathDocumentoSelector"
    Public Shared Property PathDocumento(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return s(idcontrol & ctePathDocumento)
        End Get
        Set(ByVal value As String)
            s(idcontrol & ctePathDocumento) = value
        End Set
    End Property
    Private Const cteRolActual As String = "CalidadRolusuarioActual"
    Public Shared Property RolActual(ByVal s As HttpSessionState, ByVal a As HttpApplicationState) As Long
        Get
            Return s(cteRolActual)
        End Get
        Set(ByVal value As Long)
            'Dim conexion As ConexionSQL
            Dim sen As New SentenciaSQL
            s(cteRolActual) = value
        End Set
    End Property
    Private Const cteOrdenCampoEstado As String = "CalidadOrdenCampoEstado"
    Public Shared Property OrdenCampoEstado(ByVal idcontrol As String, ByVal s As HttpSessionState) As Integer
        Get
            Return s(idcontrol & cteOrdenCampoEstado)
        End Get
        Set(ByVal value As Integer)
            s(idcontrol & cteOrdenCampoEstado) = value
        End Set
    End Property
    Private Const cteOrdenCampoID As String = "CalidadOrdenCampoID"
    Public Shared Property OrdenCampoID(ByVal idcontrol As String, ByVal s As HttpSessionState) As Integer
        Get
            Return s(idcontrol & cteOrdenCampoID)
        End Get
        Set(ByVal value As Integer)
            s(idcontrol & cteOrdenCampoID) = value
        End Set
    End Property
    Public Const cteNorma As String = "Norma"
    Public Shared ReadOnly Property Norma(ByVal s As HttpApplicationState) As String
        Get
            Return s(cteNorma)
        End Get
    End Property

    Public Const cteNumeroMaximoNodos As String = "NumeroMaximoNodos"
    Public Shared ReadOnly Property NumeroMaximoNodos(ByVal s As HttpApplicationState) As Integer
        Get
            Return s(cteNumeroMaximoNodos)
        End Get
    End Property

    Public Const cteNumeroMaximoResultadosBusqueda As String = "NumeroMaximoResultadosBusqueda"
    Public Shared ReadOnly Property NumeroMaximoResultadosBusqueda(ByVal s As HttpApplicationState) As Integer
        Get
            Return s(cteNumeroMaximoResultadosBusqueda)
        End Get
    End Property
    Public Const cteNumeroMaximoNodosObjetos As String = "NumeroMaximoNodosObjeto"
    Public Shared ReadOnly Property NumeroMaximoNodosObjetos(ByVal s As HttpApplicationState) As Integer
        Get
            Return s(cteNumeroMaximoNodosObjetos)
        End Get
    End Property
    Public Const cteNumeroAvance As String = "NumeroAvance"
    Public Shared ReadOnly Property NumeroAvance(ByVal s As HttpApplicationState) As Integer
        Get
            Return s(cteNumeroAvance)
        End Get
    End Property
    Private Const cteOrdenCampoCodigoDocumento As String = "CalidadOrdenCampoCodigoDocumento"
    Public Shared Property OrdenCampoCodigoDocumento(ByVal idcontrol As String, ByVal s As HttpSessionState) As Integer
        Get
            Return s(idcontrol & cteOrdenCampoCodigoDocumento)
        End Get
        Set(ByVal value As Integer)
            s(idcontrol & cteOrdenCampoCodigoDocumento) = value
        End Set
    End Property
    Private Const cteCampoCodigoDocumento As String = "CalidadCampoCodigoDocumento"
    Public Shared Property CampoCodigoDocumento(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return s(idcontrol & cteCampoCodigoDocumento)
        End Get
        Set(ByVal value As String)
            s(idcontrol & cteCampoCodigoDocumento) = value
        End Set
    End Property
    Private Const cteTieneRevision As String = "CalidadTieneRevisionDocActual"
    Public Shared Property TieneRevision(ByVal idcontrol As String, ByVal s As HttpSessionState) As Boolean
        Get
            Return s(idcontrol & cteTieneRevision)
        End Get
        Set(ByVal value As Boolean)
            s(idcontrol & cteTieneRevision) = value
        End Set
    End Property

    Private Const cteTieneVersion As String = "CalidadTieneRevisionDocActual"
    Public Shared Property TieneVersion(ByVal idcontrol As String, ByVal s As HttpSessionState) As Boolean
        Get
            Return s(idcontrol & cteTieneVersion)
        End Get
        Set(ByVal value As Boolean)
            s(idcontrol & cteTieneVersion) = value
        End Set


    End Property
    Private Const cteTieneImpresion As String = "CalidadTieneImpresionDocActual"
    Public Shared Property TieneImpresion(ByVal idcontrol As String, ByVal s As HttpSessionState) As Boolean
        Get
            Return s(idcontrol & cteTieneImpresion)
        End Get
        Set(ByVal value As Boolean)
            s(idcontrol & cteTieneImpresion) = value
        End Set


    End Property
    Private Const cteTablaHistorico As String = "CalidadTablaHistorico"
    Public Shared Property TablaHistorico(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return s(idcontrol & cteTablaHistorico)
        End Get
        Set(ByVal value As String)
            s(idcontrol & cteTablaHistorico) = value
        End Set


    End Property
    Private Const cteEstadoInicialDocumento As String = "CalidadEstadoInicialDocumento"
    Public Shared Property EstadoInicialDocumento(ByVal codigodocumento As String, ByVal s As HttpSessionState) As Integer
        Get
            Return s(codigodocumento & cteEstadoInicialDocumento)
        End Get
        Set(ByVal value As Integer)
            s(codigodocumento & cteEstadoInicialDocumento) = value
        End Set


    End Property

    Public Enum tiposdepermisoendocumento As Integer
        Ver = 1
        Nuevo = 2
        NuevaVersion = 3
        Editar = 4
        Borrar = 5
        Imprimir = 6
    End Enum
    Private Const ctePermisoDocumento As String = "CalidadPermisoDocumento"

    Public Shared Property TienePermiso(ByVal s As HttpSessionState, ByVal idpermiso As tiposdepermisoendocumento) As Boolean
        Get
            Return CBool(s(ctePermisoDocumento & idpermiso.ToString()))
        End Get
        Set(ByVal value As Boolean)
            s(ctePermisoDocumento & idpermiso.ToString()) = value
        End Set
    End Property
    Private Const ctePermisosTransaccionDocumento As String = "CalidadPermisoTransaccionDocumento"
    Public Shared Property PermisosTransaccionDocumento(ByVal idcontrol As String, ByVal s As HttpSessionState) As Coleccion_PermisoDocumento
        Get
            Return DirectCast(s(idcontrol & ctePermisosTransaccionDocumento), Coleccion_PermisoDocumento)
        End Get
        Set(ByVal value As Coleccion_PermisoDocumento)
            s(idcontrol & ctePermisosTransaccionDocumento) = value
        End Set
    End Property
    Private Const ctePermisosEstadoDocumento As String = "CalidadPermisoEstadoDocumento"
    Public Shared Property PermisosEstadoDocumento(ByVal idcontrol As String, ByVal s As HttpSessionState) As Coleccion_EstadoDocumento
        Get
            Return DirectCast(s(idcontrol & ctePermisosEstadoDocumento), Coleccion_EstadoDocumento)
        End Get
        Set(ByVal value As Coleccion_EstadoDocumento)
            s(idcontrol & ctePermisosEstadoDocumento) = value
        End Set
    End Property

    'hijos
    Private Const ctehijos As String = "Calidadhijos"
    Public Shared Property hijos(ByVal idcontrol As String, ByVal s As HttpSessionState) As Coleccion_Hijos
        Get
            Return DirectCast(s(idcontrol & ctehijos), Coleccion_Hijos)
        End Get
        Set(ByVal value As Coleccion_Hijos)
            s(idcontrol & ctehijos) = value
        End Set
    End Property
    Private Const cteCodigoDocumento As String = "Calidadcodigodocumento"
    Public Shared Property Codigodocumento(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(idcontrol & cteCodigoDocumento))
        End Get
        Set(ByVal value As String)
            s(idcontrol & cteCodigoDocumento) = value
        End Set
    End Property
    Private Const cteUltimoFiltro As String = "CalidadUltimoFiltroSelector"
    Public Shared Property UltimoFiltro(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(idcontrol & cteUltimoFiltro))
        End Get
        Set(ByVal value As String)
            s(idcontrol & cteUltimoFiltro) = value

        End Set
    End Property


    Private Const cteEsSistemaGeneral As String = "CalidadEsSistemaGeneral"
    Public Shared Property esSistemaGeneral(ByVal s As HttpSessionState) As Boolean
        Get
            Return CBool(s(cteEsSistemaGeneral))
        End Get
        Set(ByVal value As Boolean)
            s(cteEsSistemaGeneral) = value
        End Set
    End Property

    Private Const cteDireccion As String = "CalidadDireccion"
    Public Shared Property Direccion(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteDireccion))
        End Get
        Set(ByVal value As String)
            s(cteDireccion) = value
        End Set
    End Property
    Private Const cteUnidad As String = "CalidadUnidad"
    Public Shared Property Unidad(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteUnidad))
        End Get
        Set(ByVal value As String)
            s(cteUnidad) = value
        End Set
    End Property

    Private Const cteBdActual As String = "DocBoxBdActual"
    Public Shared Property BDActual(ByVal s As HttpSessionState) As Long
        Get
            If s(cteBdActual) Is Nothing Then
                Return -1
            Else
                Return CLng(s(cteBdActual))
            End If

        End Get
        Set(ByVal value As Long)
            s(cteBdActual) = value
        End Set
    End Property
    Private Const cteperfil As String = "perfil"
    Public Shared Property Perfil(ByVal s As HttpSessionState, ByVal a As HttpApplicationState) As String
        Get
            Return CStr(s(cteperfil))
        End Get
        Set(ByVal value As String)
            s(cteperfil) = value
        End Set
    End Property
    Private Const cteUsuario As String = "CalidadUsuario"
    Public Shared Property Usuario(ByVal s As HttpSessionState, ByVal a As HttpApplicationState) As String
        Get
            Return CStr(s(cteUsuario))
        End Get
        Set(ByVal value As String)
            Dim sen As SentenciaSQL
            Dim con As ConexionSQL

            s(cteUsuario) = value
            sen = New SentenciaSQL

            With sen
                .sql_select = CAMPO_CODUSUARIO
                .sql_from = TABLA_USUARIO
                .add_condicion(CAMPO_CODUSUARIO, value)
            End With

            con = New ConexionSQL(VariablesGlobales.CadenadeConexion(a))
            Try
                VariablesGlobales.NombreUsuario(s) = con.ejecuta1v_string(sen.texto_sql)
            Finally
                con.CerrarConexion()
            End Try

        End Set
    End Property
    Private Const cteNombreUsuario As String = "CalidadNombreUsuario"
    Public Shared Property NombreUsuario(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteNombreUsuario))
        End Get
        Set(ByVal value As String)
            s(cteNombreUsuario) = value
        End Set
    End Property
    Private Const ctePermisoMedico As String = "PermisoMedico"
    Public Shared Property PermisoMedico(ByVal s As HttpSessionState) As Boolean
        Get
            Return CBool(s(ctePermisoMedico))
        End Get
        Set(ByVal value As Boolean)
            s(ctePermisoMedico) = value
        End Set
    End Property
    Private Const cteEjerecicioSolicitud As String = "EjercicioSolicitud"
    Public Shared Property EjercicioSolicitud(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteEjerecicioSolicitud))
        End Get
        Set(ByVal value As String)
            s(cteEjerecicioSolicitud) = value
        End Set
    End Property

    Private Const cteIDDocumentoActualVisor As String = "CalidadTipoDocumentoActualVisor"
    Public Shared Property IDDocumentoActualVisor(ByVal coddocumento As String, ByVal s As HttpSessionState) As Long
        Get
            Return CLng(s(coddocumento & "-" & cteIDDocumentoActualVisor))
        End Get
        Set(ByVal value As Long)
            s(coddocumento & "-" & cteIDDocumentoActualVisor) = value
        End Set
    End Property
    Private Const cteTipodeAperturaVisor As String = "CalidadTipoAperturaVisor"
    Public Enum tav_valores
        Editar
        Nuevo
        Nueva_Version
        Ver
    End Enum
    Public Shared Property TipodeAperturaVisor(ByVal coddocumento As String, ByVal s As HttpSessionState) As tav_valores
        Get
            Return s(coddocumento & "-" & cteTipodeAperturaVisor)
        End Get
        Set(ByVal value As tav_valores)
            s(coddocumento & "-" & cteTipodeAperturaVisor) = value
        End Set
    End Property

    Private Const cteVisorAsociado As String = "CalidadVisorAsociadoSelector"
    Public Shared Property VisorAsociado(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(idcontrol & "-" & cteVisorAsociado))
        End Get
        Set(ByVal value As String)
            s(idcontrol & "-" & cteVisorAsociado) = value
        End Set
    End Property

    Private Const cteSelectorAsociado As String = "CalidadSelectorAsociadoSelector"
    Public Shared Property SelectorAsociado(ByVal idcontrol As String, ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(idcontrol & "-" & cteSelectorAsociado))
        End Get
        Set(ByVal value As String)
            s(idcontrol & "-" & cteSelectorAsociado) = value
        End Set
    End Property

    Private Const cteControlActualesVisor As String = "CalidadControlActualesVisor"
    Public Shared Property ControlesActualesVisor(ByVal idcontrol As String, ByVal s As HttpSessionState) As coleccion_control_visor
        Get
            Return CType(s(idcontrol & "-" & cteControlActualesVisor), coleccion_control_visor)
        End Get
        Set(ByVal value As coleccion_control_visor)
            s(idcontrol & "-" & cteControlActualesVisor) = value
        End Set
    End Property

    Private Const cteControlActualesobjetosVisor As String = "CalidadControlActualesobjetosVisor"
    Public Shared Property ControlesActualesobjetosVisor(ByVal idcontrol As String, ByVal s As HttpSessionState) As coleccion_control_visor
        Get
            Return CType(s(idcontrol & "-" & cteControlActualesobjetosVisor), coleccion_control_visor)
        End Get
        Set(ByVal value As coleccion_control_visor)
            s(idcontrol & "-" & cteControlActualesobjetosVisor) = value
        End Set
    End Property

    Private Const cteControlActualesaltaobjetosVisor As String = "CalidadControlActualesaltaobjetosVisor"
    Public Shared Property ControlesActualesaltaobjetosVisor(ByVal idcontrol As String, ByVal s As HttpSessionState) As coleccion_control_visor
        Get
            Return CType(s(idcontrol & "-" & cteControlActualesaltaobjetosVisor), coleccion_control_visor)
        End Get
        Set(ByVal value As coleccion_control_visor)
            s(idcontrol & "-" & cteControlActualesaltaobjetosVisor) = value
        End Set
    End Property
    Private Const cteDocumentoPadre As String = "CalidadDocumentoPadre"
    Public Shared Property DocumentoPadre(ByVal idcoddoc As String, ByVal s As HttpSessionState) As LibrodeFamilia

        Get
            Return s(idcoddoc & "-" & cteDocumentoPadre)
        End Get
        Set(ByVal value As LibrodeFamilia)
            s(idcoddoc & "-" & cteDocumentoPadre) = value
        End Set
    End Property

    Private Const cteDocumentoActualHistorico As String = "CalidadDocumentoActualHistorico"
    Public Shared Property DocumentoActualHistorico(ByVal idcontrol As String, ByVal s As HttpSessionState) As Boolean
        Get
            Return CBool(s(idcontrol & "-" & cteDocumentoActualHistorico))
        End Get
        Set(ByVal value As Boolean)
            s(idcontrol & "-" & cteDocumentoActualHistorico) = value
        End Set
    End Property
    Private Const cteIDNombre_SelectorActual As String = "CalidadIDNombre_SelectorActual"
    Public Shared Property IDNombre_SelectorActual(ByVal s As HttpSessionState) As String
        Get
            Return CStr(s(cteIDNombre_SelectorActual))
        End Get
        Set(ByVal value As String)
            s(cteIDNombre_SelectorActual) = value
        End Set
    End Property
    Private Const cteIDActual As String = "CalidadIDActual"
    Public Shared Property IDActual(ByVal s As HttpSessionState) As Long
        Get
            Return CLng(s(cteIDActual))
        End Get
        Set(ByVal value As Long)
            s(cteIDActual) = value
        End Set
    End Property
    Private Const cteHomologacionActual As String = "CalidadHomologacionActual"
    Public Shared Property HomologacionActual(ByVal s As HttpSessionState) As Long
        Get
            Return CLng(s(cteHomologacionActual))
        End Get
        Set(ByVal value As Long)
            s(cteHomologacionActual) = value
        End Set
    End Property
    Public Const cteIDInstalacion As String = "CalidadIDInstalacion"
    Public Shared Property IdInstalacion(ByVal s As HttpApplicationState) As String
        Get
            Return CLng(s(cteIDInstalacion))
        End Get
        Set(ByVal value As String)
            s(cteIDInstalacion) = value
        End Set
    End Property
    Public Const cteNumMaxUsuarios As String = "CalidadNumMaxUsuarios"
    Public Shared Property NumMaxUsuarios(ByVal s As HttpApplicationState) As String
        Get
            Return CLng(s(cteNumMaxUsuarios))
        End Get
        Set(ByVal value As String)
            s(cteNumMaxUsuarios) = value
        End Set
    End Property

    Public Const cteNumUsuariosActuales As String = "CalidadNumUsuariosActuales"
    Public Shared Property NumUsuariosActuales(ByVal s As HttpApplicationState) As String
        Get
            Return CLng(s(cteNumUsuariosActuales))
        End Get
        Set(ByVal value As String)
            s(cteNumUsuariosActuales) = value
        End Set
    End Property
    Public Shared ReadOnly Property NombreFicheroEnDisco(ByVal Session As HttpSessionState, ByVal a As HttpApplicationState, ByVal pathdocumento As String, ByVal iddocumento As Long, ByVal idcontrol As String) As String
        Get
            Dim s As String
            Dim pos As Integer
            If pathdocumento.StartsWith(".") Then
                pos = 1
                If pathdocumento.Substring(pos).StartsWith(System.IO.Path.DirectorySeparatorChar) Then
                    pos = 2
                End If
                s = Util.NormalizaRuta(VariablesGlobales.DirectorioBaseAdjuntos(a)) & pathdocumento.Substring(pos) & VariablesGlobales.Direccion(Session) & "-" & VariablesGlobales.Unidad(Session) & System.IO.Path.DirectorySeparatorChar & iddocumento
            Else
                s = pathdocumento & VariablesGlobales.Direccion(Session) & "-" & VariablesGlobales.Unidad(Session) & System.IO.Path.DirectorySeparatorChar & iddocumento
            End If
            Return System.IO.Path.GetFullPath(s & "-" & idcontrol)
            'Me.IDDocumentoActualVisor
        End Get
    End Property

End Class
Public Class EnviodeEmail

    Public Function GetStreamFile(ByVal filePath As String) As IO.Stream
        Using fileStream As IO.FileStream = System.IO.File.OpenRead(filePath)
            Dim memStream As New IO.MemoryStream()
            memStream.SetLength(fileStream.Length)
            fileStream.Read(memStream.GetBuffer(), 0, CInt(fileStream.Length))
            Return memStream
        End Using
    End Function
    Public Function EnviarEmail_html(ByVal De As String, ByVal Para As String, ByVal Asunto As String, ByVal Cuerpo As String, ByVal Adjunto As String, a As System.Web.HttpApplicationState) As Boolean
        Dim oSmtp As System.Net.Mail.SmtpClient, oMsg As New System.Net.Mail.MailMessage


        Dim Var_Smtp As String = System.Configuration.ConfigurationManager.AppSettings("ServidorCorreo").ToString

        'Dim Var_Smtpuser As String = System.Configuration.ConfigurationManager.AppSettings("UsuarioCorreo").ToString
        'Dim Var_SmtpPass As String = System.Configuration.ConfigurationManager.AppSettings("Password_Correo").ToString
        Dim Var_Smtpuser As String = VariablesGlobales.usu_correo(a)
        Dim Var_SmtpPass As String = VariablesGlobales.Pas_Correo(a)
        Dim tipo As String
        Try
            'Var_Smtpuser = "solicitudescampus.obrasocialunicaja.com"
            'De = "solicitudescampus@obrasocialunicaja.com"
            'Var_SmtpPass = "n8Z1p%sV"
            'Var_Smtp = "mail.obrasocialunicaja.com"
            oMsg.Subject = Asunto
            oMsg.To.Add(Para)
            oMsg.From = New System.Net.Mail.MailAddress(De)
            oMsg.IsBodyHtml = True
            oMsg.Body = Cuerpo


            'Dim oAttch As Net.Mail.Attachment = New Net.Mail.Attachment(GetStreamFile(Adjunto))
            'oMsg.Attachments.Add(oAttch)
            If Adjunto <> VACIO Then
                tipo = TipoMime(System.IO.Path.GetExtension(Adjunto).ToUpper)
                oMsg.Attachments.Add(New Net.Mail.Attachment(GetStreamFile(Adjunto), Adjunto, tipo))
            End If

            oSmtp = New System.Net.Mail.SmtpClient(Var_Smtp)
            oSmtp.Credentials = New System.Net.NetworkCredential(Var_Smtpuser, Var_SmtpPass)

            oSmtp.Send(oMsg)


            'Dim file As System.IO.File
            'file.Delete(Adjunto)
            If Adjunto <> VACIO Then
                'IO.File.Delete(Adjunto)
            End If

        Catch ex As Exception
            'MsgBox("Error:" & ex.Message)
            oSmtp = Nothing
            oMsg = Nothing
            Return False
            Exit Function
        Finally
        End Try
        oSmtp = Nothing
        oMsg = Nothing
        Return True
    End Function
    Public Function EnviarEmail(ByVal De As String, ByVal Para As String, ByVal Asunto As String, ByVal Cuerpo As String, ByVal Adjunto As String, a As System.Web.HttpApplicationState) As Boolean
        Dim oSmtp As System.Net.Mail.SmtpClient, oMsg As New System.Net.Mail.MailMessage


        Dim Var_Smtp As String = System.Configuration.ConfigurationManager.AppSettings("ServidorCorreo").ToString

        'Dim Var_Smtpuser As String = System.Configuration.ConfigurationManager.AppSettings("UsuarioCorreo").ToString
        'Dim Var_SmtpPass As String = System.Configuration.ConfigurationManager.AppSettings("Password_Correo").ToString
        Dim Var_Smtpuser As String = VariablesGlobales.usu_correo(a)
        Dim Var_SmtpPass As String = VariablesGlobales.Pas_Correo(a)
        Dim tipo As String
        Try

            oMsg.Subject = Asunto
            oMsg.To.Add(Para)
            oMsg.From = New System.Net.Mail.MailAddress(De)
            oMsg.IsBodyHtml = True
            oMsg.Body = Cuerpo


            'Dim oAttch As Net.Mail.Attachment = New Net.Mail.Attachment(GetStreamFile(Adjunto))
            'oMsg.Attachments.Add(oAttch)
            If Adjunto <> VACIO Then
                tipo = TipoMime(System.IO.Path.GetExtension(Adjunto).ToUpper)
                oMsg.Attachments.Add(New Net.Mail.Attachment(GetStreamFile(Adjunto), Adjunto, tipo))
            End If

            oSmtp = New System.Net.Mail.SmtpClient(Var_Smtp)
            oSmtp.Credentials = New System.Net.NetworkCredential(Var_Smtpuser, Var_SmtpPass)

            oSmtp.Send(oMsg)


            'Dim file As System.IO.File
            'file.Delete(Adjunto)
            If Adjunto <> VACIO Then
                IO.File.Delete(Adjunto)
            End If

        Catch ex As Exception
            'MsgBox("Error:" & ex.Message)
            oSmtp = Nothing
            oMsg = Nothing
            Return False
            Exit Function
        Finally
        End Try
        oSmtp = Nothing
        oMsg = Nothing
        Return True
    End Function
    Public Function TipoMime(ByVal extension As String) As String
        'Habria que unificar esta funcion en el UTIL (pero ya existe en unicaja, por ejemplo)
        Select Case extension.Substring(1) ' lequitamos el punto
            Case "SCAN"
                Return "image/jpeg"
            Case "PDF"
                Return "application/pdf"
            Case "JPG", "JPEG"
                Return "image/jpeg"
            Case "PNG"
                Return "image/png"
            Case "TIFF", "TIF"
                Return "image/tiff"
            Case "BMP"
                Return "image/bmp"
            Case "GIF"
                Return "imagen/gif"
            Case "TXT"
                Return "text/plain"
            Case "DOC", "DOT", "W6W", "WIZ", "WORD"
                Return "application/msword"
            Case "DOCX"
                Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            Case "ODT"
                Return "application/vnd.oasis.opendocument.text"
            Case "XPS"
                Return "application/vnd.ms-xpsdocument"
            Case "RTF"
                Return "text/rtf"
            Case "DWG"
                Return "application/acad"
            Case "ARJ"
                Return "application/arj"
            Case "BOO", "BOOK"
                Return "application/book"
            Case "CDF"
                Return "application/cdf"
            Case "DRW"
                Return "application/drafting"
            Case "XL", "XLA", "XLB", "XLC", "XLS"
                Return "application/excel"
            Case "TGZ"
                Return "application/gnutar"
            Case "POT", "PPS", "PPT", "PPZ"
                Return "application/mspowerpoint"

            Case "AI", "EPS", "PS"
                Return "application/postscript"

            Case "WP", "WP5", "WP6", "WPD", "W60", "WP5", "W61"
                Return "application/wordperfect"
            Case "BZ"
                Return "Application/x-bzip"
            Case "BOZ", "BZ2"
                Return "application/x-bzip2"
            Case "GZ", "TGZ", "Z", "ZIP"
                Return "application/x-compressed"
            Case "GTAR"
                Return "application/x-gtar"
            Case "GZ", "GZIP"
                Return "application/x-gzip"
            Case "MPC", "MPT", "MPV", "MPX"
                Return "application/x-project"
            Case "VSD", "VST", "VSW"
                Return "application/x-visio"
            Case "ZIP"
                Return "application/x-zip-compressed"
            Case "XML"
                Return "application/xml"
            Case "AIF", "AIFC", "AIFF"
                Return "audio/aiff"
            Case "AU", "SND"
                Return "audio/basic"
            Case "M2A", "MPA", "MPG", "MPGA"
                Return "audio/mpeg"
            Case "MP3"
                Return "audio/mpeg"

            Case Else
                Return "application/octet-stream"
        End Select
    End Function

End Class
Public Class control_visor
    Public Const TIPO_CONTROL_TEXTBOX As String = "T"
    Public Const TIPO_CONTROL_DROPDOWNLIST As String = "C"
    Public Const TIPO_CONTROL_VINCULO As String = "V"
    Public Const TIPO_CONTROL_VINCULO_EXPRESION As String = "E"
    Public Const TIPO_CONTROL_CALENDARIO As String = "A"
    'el control vinculo sera validado automaticamente con los campos definidos en la validacion
    Public Const TIPO_CONTROL_LABEL As String = "L"
    Public tipo As String

    Public etiqueta As String
    Public id As String
    Public tamaño As Integer
    Public ancho As Integer
    Public Fuente_en_negrita As Boolean
    Public fuente_tamaño As Integer
    Public fuente_color As System.Drawing.Color
    Public fuente_color_fondo As System.Drawing.Color
    Public texto As String
    Public selector As String
    Public camposelector_valor As String
    Public camposelector_texto As String

    Public obligatorio As String
    Public validacion As String

    Public nombretemporal As String

    Public valor_fecha As Date

    Public Const TIPO_SQL_FILENAME As String = "FILE"
    Public Const TIPO_SQL_NVARCHAR As String = "T"
    Public Const TIPO_SQL_DECIMAL As String = "N"
    Public Const TIPO_SQL_DATETIME As String = "F"
    Public tiposql As String
End Class

Public Class coleccion_control_visor
    Inherits coleeccion_object

    Public Const PREFIJO_BOTON_IMAGEN As String = "BOTONIMAGEN"
    Public Const PREFIJO_UPLOAD As String = "UPLOAD"
    Public Const SIN_FICHERO_ADJUNTO As String = "Sin Fichero Adjunto"


    Public Const PREFIJO_CALENDAR As String = "CALENDARIO_"
    Public Const PREFIJO_CALENDAR_TEXTBOX As String = "CALENDARIO_TEXTBOX_"
    Public Const PREFIJO_CALENDAR_BUTTON As String = "CALENDARIO_BUTTON_"
    Public Const PREFIJO_VINCULO_BUTTON As String = "VINCULO_BUTTON_"

    Public Overloads Sub Add(ByVal nombre As String, ByVal valor As control_visor)
        MyBase.Add(nombre, valor)
    End Sub

    Public Shadows Function Item(ByVal index As Integer) As control_visor
        Return CType(MyBase.Item(index).valor, control_visor)
    End Function

    Public Shadows Function ITem(ByVal nombre As String) As control_visor
        Return CType(MyBase.Item(nombre).valor, control_visor)
    End Function

    Public Sub LeeDefinicionCampos(ByVal cadenasql As String, ByVal nombretabla As String, ByVal idcuadro As Long, ByVal tipoapertura As VariablesGlobales.tav_valores, ByVal idserie As Long)
        Dim sen As New SentenciaSQL
        Dim dt, dt_combo As DataTable
        Dim dr As DataRow
        Dim con As New ConexionSQL(cadenasql)

        Dim valor_campos as string = string.empty
        Dim cv As control_visor
        Dim cad As String
        'FICHAS

        With sen
            .add_campo_select(CAMPO_NOMBRE, CAMPO_ID, CAMPO_TIPOSQL, CAMPO_TIPOCON, CAMPO_TAMAÑO, CAMPO_SELECTOR, CAMPO_OBLIGATORIO, CAMPO_VALIDA)
            .sql_from = nombretabla
            .add_condicion(CAMPO_IDSERIE, idserie)
            .sql_orderby = CAMPO_ORDEN

        End With
        dt = con.SelectSQL(sen.texto_sql)

        For Each dr In dt.Rows
            cv = New control_visor
            cv.id = dr(CAMPO_ID).ToString
            cv.etiqueta = dr(CAMPO_NOMBRE).ToString
            cv.tipo = dr(CAMPO_TIPOCON).ToString.ToUpper
            cv.tiposql = dr(CAMPO_TIPOSQL).ToString.ToUpper
            cv.tamaño = CInt(dr(CAMPO_TAMAÑO))
            cv.ancho = cv.tamaño * 8
            cv.obligatorio = dr(CAMPO_OBLIGATORIO).ToString
            cv.validacion = dr(CAMPO_VALIDA).ToString


            'If tipoapertura = VariablesGlobales.tav_valores.Editar Or tipoapertura = VariablesGlobales.tav_valores.Ver Then
            With sen
                .Limpia()
                .sql_select = CAMPO_VALOR
                .sql_from = TABLA_DATOS
                .add_condicion(CAMPO_IDCUADRO, idcuadro)
                .add_condicion(CAMPO_IDLSERIE, CLng(cv.id))
            End With

            cad = con.ejecuta1v_string(sen.texto_sql)


            Select Case cv.tiposql
                Case control_visor.TIPO_SQL_DATETIME
                    If cad <> "" Then
                        cv.valor_fecha = DateSerial(CInt(cad.Substring(6, 4)), CInt(cad.Substring(3, 2)), CInt(cad.Substring(0, 2)))
                        cv.texto = cv.valor_fecha.ToString(Util.FORMATO_FECHA_DMA)
                    End If
                Case Else
                    cv.texto = cad
            End Select

            If Not dr(CAMPO_SELECTOR).ToString Is System.DBNull.Value Then
                cv.selector = dr(CAMPO_SELECTOR).ToString
                If cv.selector <> VACIO And cv.tipo = control_visor.TIPO_CONTROL_DROPDOWNLIST Then
                    dt_combo = con.SelectSQL(cv.selector)

                    Dim obj As Object
                    If dt_combo.Rows.Count > 0 Then
                        obj = dt_combo.Rows(0)(0)
                        If Not obj Is System.DBNull.Value Then
                            If cv.texto Is Nothing Then
                                cv.texto = CStr(obj)
                            End If
                            cv.camposelector_valor = dt_combo.Columns(0).ColumnName
                            If dt_combo.Columns.Count > 1 Then
                                cv.camposelector_texto = dt_combo.Columns(1).ColumnName
                            Else
                                cv.camposelector_texto = cv.camposelector_valor
                            End If
                        End If
                        con.Libera(dt_combo)
                    End If
                End If
            End If



            MyBase.Add(cv.id, cv)
        Next
        con.LiberaTodo()
        con.CerrarConexion()

    End Sub
    Public Sub LeeDefinicionCampos_objeto(ByVal cadenasql As String, ByVal nombretabla As String, ByVal objeto As String, ByVal tipoapertura As VariablesGlobales.tav_valores, ByVal idobjeto As Long)
        Dim sen As New SentenciaSQL
        Dim dt, dt_combo As DataTable
        Dim dr As DataRow
        Dim con As New ConexionSQL(cadenasql)
        Dim tipo As String = "Objeto"
        Dim valor_campos as string = string.empty
        Dim cv As control_visor
        Dim cad As String
        'FICHAS

        With sen
            .add_campo_select(CAMPO_NOMBRE, CAMPO_ID, CAMPO_TIPOSQL, CAMPO_TIPOCON, CAMPO_TAMAÑO, CAMPO_SELECTOR, CAMPO_OBLIGATORIO, CAMPO_VALIDA)
            .sql_from = nombretabla
            .add_condicion(CAMPO_IDOBJETO, idobjeto)
            .sql_orderby = CAMPO_ORDEN

        End With
        dt = con.SelectSQL(sen.texto_sql)

        For Each dr In dt.Rows
            cv = New control_visor
            cv.id = tipo & dr(CAMPO_ID).ToString
            cv.etiqueta = dr(CAMPO_NOMBRE).ToString
            cv.tipo = dr(CAMPO_TIPOCON).ToString.ToUpper
            cv.tiposql = dr(CAMPO_TIPOSQL).ToString.ToUpper
            cv.tamaño = CInt(dr(CAMPO_TAMAÑO))
            cv.ancho = cv.tamaño * 8
            cv.obligatorio = dr(CAMPO_OBLIGATORIO).ToString
            cv.validacion = dr(CAMPO_VALIDA).ToString


            'If tipoapertura = VariablesGlobales.tav_valores.Editar Or tipoapertura = VariablesGlobales.tav_valores.Ver Then
            With sen
                .Limpia()
                .sql_select = CAMPO_VALOR
                .sql_from = TABLA_DATOS_OBJETO
                .add_condicion(CAMPO_IDOBJETO, idobjeto)
                .add_condicion(CAMPO_IDLOBJETO, CLng(dr(CAMPO_ID).ToString))
                .add_condicion(CAMPO_OBJETO, objeto)
            End With

            cad = con.ejecuta1v_string(sen.texto_sql)


            Select Case cv.tiposql
                Case control_visor.TIPO_SQL_DATETIME
                    If cad <> "" Then
                        cv.valor_fecha = DateSerial(CInt(cad.Substring(6, 4)), CInt(cad.Substring(3, 2)), CInt(cad.Substring(0, 2)))
                        cv.texto = cv.valor_fecha.ToString(Util.FORMATO_FECHA_DMA)
                    End If
                Case Else
                    cv.texto = cad
            End Select

            'End If


            If Not dr(CAMPO_SELECTOR).ToString Is System.DBNull.Value Then
                cv.selector = dr(CAMPO_SELECTOR).ToString
                If cv.selector <> VACIO And cv.tipo = control_visor.TIPO_CONTROL_DROPDOWNLIST Then
                    dt_combo = con.SelectSQL(cv.selector)

                    Dim obj As Object
                    If dt_combo.Rows.Count > 0 Then
                        obj = dt_combo.Rows(0)(0)
                        If Not obj Is System.DBNull.Value Then
                            If cv.texto Is Nothing Then
                                cv.texto = CStr(obj)
                            End If
                            cv.camposelector_valor = dt_combo.Columns(0).ColumnName
                            If dt_combo.Columns.Count > 1 Then
                                cv.camposelector_texto = dt_combo.Columns(1).ColumnName
                            Else
                                cv.camposelector_texto = cv.camposelector_valor
                            End If
                        End If
                        con.Libera(dt_combo)
                    End If
                End If
            End If



            MyBase.Add(cv.id, cv)
        Next
        con.LiberaTodo()
        con.CerrarConexion()

    End Sub

    Function GetLiteral(ByVal text As String)
        Dim rv As Literal
        rv = New Literal
        rv.Text = text
        GetLiteral = rv
    End Function

    ''' <summary>
    ''' Función para asociar Datepicker a un id de objeto.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub activarcalendario(ByVal a As System.Web.UI.Control, ByVal txt As String)

        Dim script As String = "activar_calendario('" & "#" & "ContentPlaceHolder1_" & txt & "');"
        ScriptManager.RegisterStartupScript(a, GetType(Page), script, script, True)

    End Sub

    Public Sub CargaFicha(ByVal a As HttpApplicationState, ByVal wf As System.Web.UI.Control, ByVal ph_datos As PlaceHolder, ByVal tipodeapertura As VariablesGlobales.tav_valores, ByVal cadenasql As String, Optional ByVal mensaje As String = "")
        Dim i As Integer
        Dim cv As control_visor
        Dim lbl As Label
        Dim etilbl As Label
        Dim txt As TextBox
        Dim combo As DropDownList
        Dim vincu1 As Button
        Dim vincu As HyperLink
        Dim sqlds As SqlDataSource
        'Dim VR As RequiredFieldValidator
        'Dim VE As RegularExpressionValidator
        Dim VS As ValidationSummary
        Dim infopadre As String = String.Empty
        Dim columna As Integer = 3
        Dim ancho As Integer = 0
        Dim totalancho As Integer = 0
        Dim salto As Boolean = False

        Dim classlabel As String = "col-md-4 control-label pq"
        Dim classtxt As String = "form-control pq"
        Dim classlabeloblig As String = "control-label pq"


        For i = 0 To Me.Count - 1

            cv = Me.Item(i)

            lbl = New Label
            lbl.Text = cv.etiqueta
            lbl.CssClass = classlabel

            Select Case cv.tipo

                Case control_visor.TIPO_CONTROL_LABEL
                    etilbl = New Label
                    etilbl.ID = cv.id
                    etilbl.CssClass = classtxt
                    etilbl.Text = cv.texto

                    ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                    ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                    ph_datos.Controls.Add(lbl)
                    ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                    ph_datos.Controls.Add(etilbl)
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))

                Case control_visor.TIPO_CONTROL_VINCULO

                    vincu1 = New Button
                    With vincu1
                        .ID = PREFIJO_VINCULO_BUTTON & cv.id
                        .ToolTip = cv.texto
                        .Text = "..."
                        .Height = 19
                        .Width = 24
                        .CausesValidation = False
                        .UseSubmitBehavior = False
                    End With
                    vincu1.OnClientClick = "javascript:AbrirVentana('" & cv.texto & "');"

                    ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                    ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                    ph_datos.Controls.Add(lbl)
                    ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                    ph_datos.Controls.Add(vincu1)
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))

                Case control_visor.TIPO_CONTROL_VINCULO_EXPRESION
                    vincu = New HyperLink
                    vincu.ID = cv.id
                    vincu.Text = cv.texto
                    vincu.NavigateUrl = cv.texto
                    vincu.Target = "_blank"

                    ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                    ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                    ph_datos.Controls.Add(lbl)
                    ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                    ph_datos.Controls.Add(vincu)
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))

                Case control_visor.TIPO_CONTROL_TEXTBOX

                    Select Case cv.tiposql
                        Case control_visor.TIPO_SQL_DATETIME

                            'FECHA

                            Dim codcontrol As String = PREFIJO_CALENDAR_TEXTBOX & cv.id
                            txt = New TextBox
                            With txt
                                .ID = PREFIJO_CALENDAR_TEXTBOX & cv.id
                                .CssClass = classtxt

                            End With
                            If tipodeapertura = VariablesGlobales.tav_valores.Ver Then
                                txt.ReadOnly = True
                            End If

                            If cv.texto <> "" Then
                                txt.Text = cv.valor_fecha.ToString(Util.FORMATO_FECHA_DMA)
                            End If
                            If tipodeapertura <> VariablesGlobales.tav_valores.Ver Then

                                txt.Attributes.Add("type", "text")

                                If cv.obligatorio = "S" Then


                                    txt.Attributes.Add("Required", "True")

                                    ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                                    ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                                    ph_datos.Controls.Add(lbl)
                                    ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                                    ph_datos.Controls.Add(txt)
                                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                    activarcalendario(wf, codcontrol)

                                Else

                                    ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                                    ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                                    ph_datos.Controls.Add(lbl)
                                    ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                                    ph_datos.Controls.Add(txt)
                                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                    ph_datos.Controls.Add(GetLiteral("                               </div>"))

                                    activarcalendario(wf, codcontrol)

                                End If
                                'validacion expresion
                                'If cv.validacion <> "" Then
                                '    VE = New RegularExpressionValidator
                                '    If cv.tiposql = control_visor.TIPO_SQL_DATETIME Then
                                '        VE.ControlToValidate = PREFIJO_CALENDAR_TEXTBOX & cv.id
                                '    Else
                                '        VE.ControlToValidate = cv.id
                                '    End If
                                '    VE.Text = "*"
                                '    VE.ErrorMessage = "El valor del campo " & cv.etiqueta & " no es correcto"
                                '    VE.ValidationExpression = cv.validacion
                                '    ph_datos.Controls.Add(VE)
                                'End If

                            End If

                        Case Else
                            txt = New TextBox
                            txt.ID = cv.id
                            txt.CssClass = classtxt

                            If tipodeapertura = VariablesGlobales.tav_valores.Ver Then
                                txt.ReadOnly = True
                            End If
                            'TODO añadir formatos Numericos, texto, fecha....
                            Select Case cv.tiposql
                                Case control_visor.TIPO_SQL_NVARCHAR
                                    txt.Attributes.Add("type", "text")
                                Case control_visor.TIPO_SQL_DECIMAL
                                    txt.Attributes.Add("type", "number")
                            End Select
                            txt.Text = cv.texto

                            If cv.obligatorio = "S" Then

                                txt.Attributes.Add("Required", "True")

                                ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                                ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                                ph_datos.Controls.Add(lbl)
                                ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                                ph_datos.Controls.Add(txt)
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))

                            Else

                                ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                                ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                                ph_datos.Controls.Add(lbl)
                                ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                                ph_datos.Controls.Add(txt)
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))

                            End If


                    End Select

                Case control_visor.TIPO_CONTROL_DROPDOWNLIST
                    combo = New DropDownList
                    combo.ID = cv.id
                    combo.CssClass = classtxt
                    combo.AutoPostBack = True
                    If tipodeapertura = VariablesGlobales.tav_valores.Ver Then
                        combo.Enabled = False
                    End If
                    If cv.selector <> "" Then
                        sqlds = New SqlDataSource
                        sqlds.ConnectionString = cadenasql
                        'si es obligatorio se controla abajo
                        'If cv.obligatorio = "N" Then
                        'sqlds.SelectCommand = SELECT_VACIO_COMBO & sql_UNION_ALL & cv.selector
                        'Else
                        sqlds.SelectCommand = cv.selector
                        'End If
                        sqlds.ID = "DS_" & cv.id
                        combo.DataSourceID = "DS_" & cv.id
                        combo.DataTextField = cv.camposelector_texto
                        combo.DataValueField = cv.camposelector_valor
                        ph_datos.Controls.Add(sqlds)
                    End If
                    If cv.texto <> "" Then combo.Text = cv.texto

                    If cv.obligatorio = "S" Then

                        ' Asignamos el atributo Obligatorio
                        combo.Attributes.Add("Required", "True")

                        ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                        ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                        ph_datos.Controls.Add(lbl)
                        ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                        ph_datos.Controls.Add(combo)
                        ph_datos.Controls.Add(GetLiteral("                               </div>"))
                        ph_datos.Controls.Add(GetLiteral("                               </div>"))
                        ph_datos.Controls.Add(GetLiteral("                               </div>"))

                    Else

                        ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                        ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                        ph_datos.Controls.Add(lbl)
                        ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                        ph_datos.Controls.Add(combo)
                        ph_datos.Controls.Add(GetLiteral("                               </div>"))
                        ph_datos.Controls.Add(GetLiteral("                               </div>"))
                        ph_datos.Controls.Add(GetLiteral("                               </div>"))

                    End If

            End Select

        Next
        'validacion MENSAJE ERRORES 
        If tipodeapertura <> VariablesGlobales.tav_valores.Ver Then
            VS = New ValidationSummary
            ph_datos.Controls.Add(VS)
        End If

    End Sub

    Public Sub CargaFicha_objeto(ByVal a As HttpApplicationState, ByVal wf As System.Web.UI.Control, ByVal ph_datos As PlaceHolder, ByVal tipodeapertura As VariablesGlobales.tav_valores, ByVal cadenasql As String, Optional ByVal encolumna As Boolean = False)
        Dim i As Integer
        Dim cv As control_visor
        Dim lbl As Label
        Dim txt As TextBox
        Dim combo As DropDownList
        Dim vincu As HyperLink
        Dim vincu1 As Button
        Dim sqlds As SqlDataSource
        'Dim VE As RegularExpressionValidator
        Dim VS As ValidationSummary
        Dim infopadre as string = string.empty
        Dim columna As Integer = 3
        Dim ancho As Integer = 0
        Dim totalancho As Integer = 0
        Dim salto As Boolean = False
        Dim tipo As String = "Objeto"
        Dim abrirdiv As New LiteralControl("<div class='col-md-8'>")
        Dim cerrardiv As New LiteralControl("</div>")
        Dim classlabel As String = "col-md-4 control-label pq"
        Dim classtxt As String = "form-control pq"

        For i = 0 To Me.Count - 1
            cv = Me.Item(i)

            lbl = New Label
            lbl.Text = cv.etiqueta
            lbl.CssClass = classlabel

            Select Case cv.tipo

                Case control_visor.TIPO_CONTROL_LABEL
                    lbl = New Label
                    lbl.ID = cv.id
                    lbl.Text = cv.texto
                    lbl.CssClass = classlabel
                    ph_datos.Controls.Add(lbl)

                Case control_visor.TIPO_CONTROL_VINCULO
                    'vincu = New HyperLink
                    'vincu.ID = cv.id
                    'vincu.Width = cv.ancho
                    'ancho = cv.ancho
                    'vincu.Text = cv.texto
                    'vincu.NavigateUrl = cv.texto
                    'vincu.Height = alto_default_objeto
                    'vincu.Target = "_blank"
                    'vincu.Font.Bold = True
                    'vincu.Font.Name = LETRA_BASE
                    'vincu.Font.Size = SIZE_BASE_objeto
                    'vincu.ForeColor = Drawing.Color.Green
                    'ph_datos.Controls.Add(vincu)
                    vincu1 = New Button
                    With vincu1
                        .ID = PREFIJO_VINCULO_BUTTON & cv.id
                        .ToolTip = cv.texto
                        .Text = "..."
                        .Height = 19
                        .Width = 24
                        .CausesValidation = False
                        .UseSubmitBehavior = False
                    End With
                    'vincu1.OnClientClick = "javascript:AbrirVentana('" & cv.texto & "');"

                    ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                    ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                    ph_datos.Controls.Add(lbl)
                    ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                    ph_datos.Controls.Add(vincu1)
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))

                Case control_visor.TIPO_CONTROL_VINCULO_EXPRESION
                    vincu = New HyperLink
                    vincu.ID = cv.id
                    vincu.Width = cv.ancho
                    vincu.Text = cv.texto
                    vincu.NavigateUrl = cv.texto
                    vincu.Target = "_blank"
                    ph_datos.Controls.Add(vincu)

                Case control_visor.TIPO_CONTROL_TEXTBOX

                    Select Case cv.tiposql
                        Case control_visor.TIPO_SQL_DATETIME

                            'FECHA

                            Dim codcontrol As String = PREFIJO_CALENDAR_TEXTBOX & cv.id
                            txt = New TextBox
                            With txt
                                .ID = PREFIJO_CALENDAR_TEXTBOX & cv.id
                                .CssClass = classtxt
                            End With
                            If tipodeapertura = VariablesGlobales.tav_valores.Ver Then
                                txt.ReadOnly = True
                            Else
                                txt.Attributes.Add("type", "text")
                            End If

                            If cv.texto <> "" Then
                                txt.Text = cv.valor_fecha.ToString(Util.FORMATO_FECHA_DMA)
                            End If
                            If tipodeapertura <> VariablesGlobales.tav_valores.Ver Then

                                ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                                ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                                ph_datos.Controls.Add(lbl)
                                ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                                ph_datos.Controls.Add(txt)
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))

                                activarcalendario(wf, codcontrol)

                            End If

                        Case Else

                            txt = New TextBox
                            txt.ID = cv.id
                            txt.CssClass = classtxt


                            If tipodeapertura = VariablesGlobales.tav_valores.Ver Then
                                txt.ReadOnly = True
                            End If
                            'TODO añadir formatos Numericos, texto, fecha....
                            Select Case cv.tiposql
                                Case control_visor.TIPO_SQL_NVARCHAR
                                    txt.Attributes.Add("type", "text")
                                Case control_visor.TIPO_SQL_DECIMAL
                                    txt.Attributes.Add("type", "number")
                            End Select
                            txt.Text = cv.texto

                            ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                            ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                            ph_datos.Controls.Add(lbl)
                            ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                            ph_datos.Controls.Add(txt)
                            ph_datos.Controls.Add(GetLiteral("                               </div>"))
                            ph_datos.Controls.Add(GetLiteral("                               </div>"))
                            ph_datos.Controls.Add(GetLiteral("                               </div>"))



                    End Select

                Case control_visor.TIPO_CONTROL_DROPDOWNLIST
                    combo = New DropDownList
                    combo.ID = cv.id
                    combo.CssClass = classtxt
                    ancho = cv.ancho + 24
                    If tipodeapertura = VariablesGlobales.tav_valores.Ver Then
                        combo.Enabled = False
                    End If
                    If cv.selector <> "" Then
                        sqlds = New SqlDataSource
                        sqlds.ConnectionString = cadenasql
                        'If cv.obligatorio = "N" Then
                        'sqlds.SelectCommand = SELECT_VACIO_COMBO & sql_UNION_ALL & cv.selector
                        'Else
                        sqlds.SelectCommand = cv.selector
                        'End If
                        sqlds.ID = "DS_" & cv.id
                        combo.DataSourceID = "DS_" & cv.id
                        combo.DataTextField = cv.camposelector_texto
                        combo.DataValueField = cv.camposelector_valor

                        ph_datos.Controls.Add(sqlds)
                        'AddHandler combo.DataBound, AddressOf Me.EstableceValoresCombo
                    End If
                    If cv.texto <> "" Then combo.Text = cv.texto

                    ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                    ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                    ph_datos.Controls.Add(lbl)
                    ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                    ph_datos.Controls.Add(combo)
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))
                    ph_datos.Controls.Add(GetLiteral("                               </div>"))

            End Select




        Next
        'validacion MENSAJE ERRORES 
        If tipodeapertura <> VariablesGlobales.tav_valores.Ver Then
            VS = New ValidationSummary
            ph_datos.Controls.Add(VS)
        End If

    End Sub
    Public Sub LeeDefinicionCampos_tablaUusuario(ByVal cadenasql As String, ByVal nombretabla As String, ByVal id As String, ByVal tipoapertura As VariablesGlobales.tav_valores)
        Dim sen As New SentenciaSQL
        Dim dt As DataTable
        Dim dr As DataRow
        Dim con As New ConexionSQL(cadenasql)
        Dim valor_campos as string = string.empty
        Dim cv As control_visor
        Dim cad As String
        Dim identifica As String
        Dim cadsql As String
        cadsql = " select c.name  FROM sys.identity_columns c  join sys.columns t ON " & _
            " c.object_id = t.object_id join sys.tables r on c.object_id =r.object_id WHERE r.name =  " & v(nombretabla) & _
            " and t.is_identity  = 1 "
        identifica = con.ejecuta1v_string(cadsql)
        'FICHAS
        cad = "SELECT c.column_id as id, c.name as nombre ,c.user_type_id as tiposql, c.max_length as tamano FROM sys.columns c JOIN sys.tables t ON c.object_id = t.object_id WHERE t.name ='" & nombretabla & "' and C.is_identity = 0"
        dt = con.SelectSQL(cad)

        For Each dr In dt.Rows
            cv = New control_visor
            cv.id = dr(CAMPO_ID).ToString
            cv.etiqueta = dr(CAMPO_NOMBRE).ToString
            cv.tipo = "T" 'dr(CAMPO_TIPOCON).ToString.ToUpper
            If dr(CAMPO_TIPOSQL).ToString = "231" Then cv.tiposql = "T"
            If dr(CAMPO_TIPOSQL).ToString = "61" Then cv.tiposql = "F"
            If dr(CAMPO_TIPOSQL).ToString = "108" Then cv.tiposql = "N"
            cv.tamaño = CInt(dr(CAMPO_TAMAÑO))
            cv.ancho = CInt(dr(CAMPO_TAMAÑO)) * 5
            If cv.ancho > 500 Then cv.ancho = 500
            If tipoapertura = VariablesGlobales.tav_valores.Editar Or tipoapertura = VariablesGlobales.tav_valores.Ver Then
                With sen
                    .Limpia()
                    .sql_select = cv.etiqueta
                    .sql_from = nombretabla
                    .add_condicion(identifica, id)
                End With

                cad = con.ejecuta1v_string(sen.texto_sql)

                Select Case cv.tiposql
                    Case control_visor.TIPO_SQL_DATETIME
                        If cad <> "" Then
                            cv.valor_fecha = DateSerial(CInt(cad.Substring(6, 4)), CInt(cad.Substring(3, 2)), CInt(cad.Substring(0, 2)))
                            cv.texto = cv.valor_fecha.ToString(Util.FORMATO_FECHA_DMA)
                        End If
                    Case Else
                        cv.texto = cad
                End Select
            End If

            MyBase.Add(cv.id, cv)
        Next
        con.LiberaTodo()
        con.CerrarConexion()

    End Sub
    Public Sub CargaFicha_tablaUusuario(ByVal a As HttpApplicationState, ByVal wf As System.Web.UI.Control, ByVal ph_datos As PlaceHolder, ByVal tipodeapertura As VariablesGlobales.tav_valores, ByVal cadenasql As String, Optional ByVal encolumna As Boolean = False)
        Dim i As Integer
        Dim cv As control_visor
        Dim lbl As Label
        Dim txt As TextBox
        Dim VS As ValidationSummary
        Dim infopadre as string = string.empty
        Dim columna As Integer = 3
        Dim ancho As Integer = 0
        Dim totalancho As Integer = 0
        Dim salto As Boolean = False
        Dim abrirdiv As New LiteralControl("<div class='col-md-8'>")
        Dim cerrardiv As New LiteralControl("</div>")
        Dim classlabel As String = "col-md-4 control-label pq"
        Dim classtxt As String = "form-control pq"

        For i = 0 To Me.Count - 1
            cv = Me.Item(i)

            lbl = New Label
            lbl.Text = cv.etiqueta
            lbl.CssClass = classlabel

            Select Case cv.tipo

                Case control_visor.TIPO_CONTROL_TEXTBOX
                    Select Case cv.tiposql
                        Case control_visor.TIPO_SQL_DATETIME

                            Dim codcontrol As String = PREFIJO_CALENDAR_TEXTBOX & cv.id
                            txt = New TextBox
                            With txt
                                .ID = PREFIJO_CALENDAR_TEXTBOX & cv.id
                                .CssClass = classtxt
                            End With
                            If tipodeapertura = VariablesGlobales.tav_valores.Ver Then
                                txt.ReadOnly = True
                            End If
                            ph_datos.Controls.Add(txt)
                            If cv.texto <> "" Then
                                txt.Text = cv.valor_fecha.ToString(Util.FORMATO_FECHA_DMA)
                            End If
                            If tipodeapertura <> VariablesGlobales.tav_valores.Ver Then

                                ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                                ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                                ph_datos.Controls.Add(lbl)
                                ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))
                                ph_datos.Controls.Add(GetLiteral("                               </div>"))

                                activarcalendario(wf, codcontrol)

                            End If

                        Case Else


                            txt = New TextBox
                            txt.ID = cv.id
                            txt.Text = cv.texto
                            txt.CssClass = classtxt
                            Select Case cv.tiposql
                                Case control_visor.TIPO_SQL_NVARCHAR
                                    txt.Attributes.Add("type", "text")
                                Case control_visor.TIPO_SQL_DECIMAL
                                    txt.Attributes.Add("type", "number")
                            End Select
                            ph_datos.Controls.Add(GetLiteral("<div class='col-sm-12 col-lg-8'>"))
                            ph_datos.Controls.Add(GetLiteral("<div class='form-group'>"))
                            ph_datos.Controls.Add(lbl)
                            ph_datos.Controls.Add(GetLiteral("<div class='col-md-8'>"))
                            ph_datos.Controls.Add(txt)
                            ph_datos.Controls.Add(GetLiteral("                               </div>"))
                            ph_datos.Controls.Add(GetLiteral("                               </div>"))
                            ph_datos.Controls.Add(GetLiteral("                               </div>"))



                    End Select

            End Select


        Next

        If tipodeapertura <> VariablesGlobales.tav_valores.Ver Then
            VS = New ValidationSummary
            ph_datos.Controls.Add(VS)
        End If

    End Sub
End Class
Public Class LibrodeFamilia
    Public idpadre As Long
    Public tipocoddocumento As String
    Public tipocoddocumentohijo As String
    Public tablasql As String
    Public brujula As String
    Public padredirecto As String
    Public brujulapadredirecto As String
    Public urlpadre As String

    Public perm_padre As New permisospadre
End Class

Public Class permisospadre
    Public nuevaversion As Boolean = True
    Public imprimir As Boolean = True
    Public editar As Boolean = True
    Public borrar As Boolean = True
    Public ver As Boolean = True
End Class






