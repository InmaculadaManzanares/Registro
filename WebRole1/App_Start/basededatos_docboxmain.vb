Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Public Module basededatos_docboxmain


    '----- funciones que deben ser globales MOVER

    Public Function Scrub(ByVal text As String) As String
        Return text.Replace("&nbsp;", "")
    End Function
    Public Sub rellenadatostablaDB_DOCEXP(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = "insert into DB_DOCEXP (nombre) values ('Oficio del ISSSTE, notificando la incapacidad temporal o permanente')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Formato de reconocimiento de antigüedad (7) del ISSSTE')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Acuerdos por designación especial o servicio de carrera')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Cédula de inscripción individual FONAC (8)')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Propuestas de personal (oficio o formato)')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Actas de reinstalación')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Cédulas de antigüedad')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Renuncias')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Avisos de cambio de situación del personal')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Nombramientos')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Solicitud de empleo')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Comprobante de Domicilio')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Actas de nacimiento, de matrimonio y de defunción')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Credencial de elector')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Constancias de estudios (certificado, primaria, secundaria, preparatoria, cédula profesional, etc.)')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Curriculum vitae')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Cartilla del servicio militar nacional')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Clave única del registro de población')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Constancias de nombramiento')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Formatos únicos de personal')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Formato del seguro de vida institucional')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Formato de de designación de beneficiarios del seguro de vida institucional')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Formato de designación de beneficiarios de sueldos y demás prestaciones no cobradas')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Formato del seguro de separación individualizado y de beneficiarios')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Hoja de filiación')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Licencias médicas y licencias con y sin goce de sueldo')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Procedimientos Administrativos (Visitaduría, OIC (1), Consejo de profesionalización, Delegación esta')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Evaluaciones del CECC (2) y Evaluaciones del CEDH (3)')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Alta, modificación y baja del ISSSTE(4)')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Formato RT-09 (5) del ISSSTE y Formato RT-01 (6) del ISSSTE')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Volante de Movimientos')" & vbCrLf
        sql += "insert into DB_DOCEXP (nombre) values ('Varios')" & vbCrLf
        conexion.Ejecuta(sql)

    End Sub
    Public Sub crearfuncionPermisoVer_serie(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = "CREATE FUNCTION [dbo].[udf_PermisoVer_Serie] " & vbCrLf
        sql += "( " & vbCrLf
        sql += " @serie nvarchar(max), " & vbCrLf
        sql += " @usuario nvarchar(max) " & vbCrLf
        'sql += " @grupo nvarchar(max) " & vbCrLf
        sql += ") " & vbCrLf
        sql += "RETURNS nvarchar(max) " & vbCrLf
        sql += "AS " & vbCrLf
        sql += "BEGIN " & vbCrLf
        sql += "declare @ver nvarchar(1) " & vbCrLf
        sql += "declare @resultado nvarchar(1) " & vbCrLf

        sql += "select @ver = P.ver FROM [DB_PERMISOS] AS P INNER JOIN [DB_CUADRO] " & vbCrLf
        sql += "AS C ON P.idcuadro = C.id WHERE (P.CodUsuario =@usuario ) AND (C.IdSerie =@serie) " & vbCrLf
        sql += "GROUP BY P.ver HAVING (P.ver = N'1')" & vbCrLf

        sql += "if @ver is null " & vbCrLf

        sql += "    begin " & vbCrLf

        sql += "Select @ver = P.ver "
        sql += "FROM         DB_PERMISOS AS P INNER JOIN" & vbCrLf
        sql += "             DB_CUADRO AS C ON P.IdCuadro = C.id INNER JOIN" & vbCrLf
        sql += "              DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE     (C.IdSerie = @serie) AND (P.Ver = N'1') AND (Grupo.CodUsuario = @usuario) AND" & vbCrLf
        sql += "                 ((SELECT     Ver" & vbCrLf
        sql += "                     FROM         DB_PERMISOS AS P1" & vbCrLf
        sql += "                    WHERE     (P.IdCuadro = IdCuadro) AND (CodUsuario = @usuario)" & vbCrLf
        sql += "                    GROUP BY Ver) IS NULL) OR" & vbCrLf
        sql += "             (C.IdSerie = @serie) AND (P.Ver = N'1') AND (Grupo.CodUsuario = @usuario) AND" & vbCrLf
        sql += "                  ((SELECT     Ver" & vbCrLf
        sql += "                     FROM         DB_PERMISOS AS P1" & vbCrLf
        sql += "                    WHERE     (P.IdCuadro = IdCuadro) AND (CodUsuario = @usuario)" & vbCrLf
        sql += "                    GROUP BY Ver) = '1')" & vbCrLf


        'sql += "	select @ver = P.ver " & vbCrLf
        'sql += "       FROM [db_PERMISOS] AS P INNER JOIN [db_CUADRO] " & vbCrLf
        'sql += "		AS C ON P.IdCuadro = C.ID " & vbCrLf
        'sql += "		WHERE   (p.IdGrupo = @GRUPO  AND C.IDSerie =@serie AND P.ver = N'1' AND " & vbCrLf
        'sql += "				(select  P1.ver FROM [db_PERMISOS] AS P1  " & vbCrLf
        'sql += "				WHERE  P.IdCuadro = P1.IdCuadro and P1.CodUsuario =@usuario " & vbCrLf
        'sql += "				GROUP BY P1.ver ) IS NULL) " & vbCrLf
        'sql += "					OR " & vbCrLf
        'sql += "				(p.IdGrupo = @GRUPO  AND C.IDSerie =@serie AND P.ver = N'1' AND " & vbCrLf
        'sql += "				(select  P1.ver FROM [db_PERMISOS] AS P1  " & vbCrLf
        'sql += "				WHERE  P.IdCuadro = P1.IdCuadro and P1.CODUsuario =@usuario " & vbCrLf
        'sql += "				GROUP BY P1.ver ) = '1')  " & vbCrLf

        sql += "    if @ver is null " & vbCrLf
        sql += "	    set @resultado = '0'" & vbCrLf
        sql += "    Else" & vbCrLf
        sql += "	    set @resultado = @ver	" & vbCrLf
        sql += "    End" & vbCrLf

        sql += "ELSE" & vbCrLf
        sql += "	set @resultado = @ver" & vbCrLf

        sql += "return @resultado " & vbCrLf
        sql += "End "
        crearfuncionoactualizarfuncion(conexion, "dbo.udf_PermisoVer_Serie", sql)
    End Sub
    Public Sub crearfuncionPermiso_ver_nodorapido(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = "CREATE FUNCTION [dbo].[udf_permiso_ver_nodorapido] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "		 @carpeta nvarchar(max), " & vbCrLf
        sql += "         @usuario nvarchar(max)" & vbCrLf
        sql += ")" & vbCrLf
        sql += " RETURNS Int" & vbCrLf
        sql += "AS" & vbCrLf
        sql += " BEGIN" & vbCrLf
        sql += "declare @resultado int" & vbCrLf
        sql += "set @resultado =2" & vbCrLf
        sql += "select @resultado=ver" & vbCrLf
        sql += "from [db_Permisos]" & vbCrLf
        sql += "where idcuadro = @carpeta and" & vbCrLf
        sql += " codusuario=@usuario 	" & vbCrLf
        sql += "if @resultado=2 -- es decir no ha devuelto ninguna fila" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "set @resultado = 0" & vbCrLf
        sql += "select @resultado = p.ver" & vbCrLf
        sql += "from [db_Permisos]  P left outer join   DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo " & vbCrLf
        sql += "where P.idcuadro = @carpeta and p.Ver=1 and		" & vbCrLf
        sql += "Grupo.codusuario = @usuario		" & vbCrLf
        sql += "    End" & vbCrLf
        sql += "return @resultado" & vbCrLf
        sql += "  End" & vbCrLf
        crearfuncionoactualizarfuncion(conexion, "dbo.udf_permiso_ver_nodorapido", sql)
    End Sub
    Public Sub crearvista_permisos(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = "CREATE VIEW [dbo].[vista_permiso_pre] AS SELECT CodUsuario AS codu, IdCuadro AS id, MAX(CAST(Ver AS int)) + 10 AS ver " & vbCrLf
        sql += "FROM dbo.DB_PERMISOS AS P GROUP BY CodUsuario, IdCuadro UNION ALL SELECT Grupo.CodUsuario AS codu, P.IdCuadro AS id, " & vbCrLf
        sql += "MAX(CAST(P.Ver AS int)) AS ver FROM dbo.DB_PERMISOS AS P INNER JOIN DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo " & vbCrLf
        sql += "GROUP BY Grupo.CodUsuario, P.IdCuadro "
        crearfuncionoactualizarvista(conexion, "dbo.vista_permiso_pre", sql)

        sql = "CREATE VIEW [dbo].[vista_permiso] AS SELECT codu, id, MAX(ver) AS ver FROM dbo.vista_permiso_pre GROUP BY codu, id HAVING (MAX(ver) = 11 OR MAX(ver) = 1)"
        crearfuncionoactualizarvista(conexion, "dbo.vista_permiso", sql)

        sql = "create VIEW [dbo].[vista_sin_permiso] AS SELECT codu, id, MAX(ver) AS ver FROM dbo.vista_permiso_pre GROUP BY codu, id HAVING (MAX(ver) IN (0, 10))"
        crearfuncionoactualizarvista(conexion, "dbo.vista_sin_permiso", sql)

        sql = "create VIEW [dbo].[vista_permisos_global] AS SELECT CodUsuario AS codu, IdCuadro AS id, MAX(CAST(Ver AS int)) + 10 AS ver, " & vbCrLf
        sql += " MAX(CAST(Añadir AS int)) + 10 AS Añadir, MAX(CAST(Modificar AS int)) + 10 AS Modificar, MAX(CAST(Borrar AS int)) + 10 AS Borrar,  " & vbCrLf
        sql += " MAX(CAST(Imprimir AS int)) + 10 AS Imprimir, MAX(CAST(Historico AS int)) + 10 AS Historico, MAX(CAST(Seguridad AS int)) + 10 AS Seguridad " & vbCrLf
        sql += " FROM dbo.DB_PERMISOS AS P GROUP BY CodUsuario, IdCuadro UNION ALL SELECT Grupo.CodUsuario AS codu, P.IdCuadro AS id, MAX(CAST(P.Ver AS int)) AS ver, " & vbCrLf
        sql += " MAX(CAST(P.Añadir AS int)) AS Añadir, MAX(CAST(P.Modificar AS int)) AS Modificar, MAX(CAST(P.Borrar AS int)) AS Borrar, " & vbCrLf
        sql += " MAX(CAST(P.Imprimir AS int)) AS Imprimir, MAX(CAST(P.Historico AS int)) AS Historico, MAX(CAST(P.Seguridad AS int)) AS Seguridad " & vbCrLf
        sql += " FROM dbo.DB_PERMISOS AS P INNER JOIN DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo GROUP BY Grupo.CodUsuario, P.IdCuadro"
        crearfuncionoactualizarvista(conexion, "dbo.vista_permisos_global", sql)

        sql = "create VIEW [dbo].[vista_permisos_global_carpetas] AS SELECT CodUsuario AS codu, IdCuadro AS id, MAX(CAST(Ver AS int)) + 10 AS ver, " & vbCrLf
        sql += " MAX(CAST(Añadir AS int)) + 10 AS Añadir, MAX(CAST(Modificar AS int)) + 10 AS Modificar, MAX(CAST(Borrar AS int)) + 10 AS Borrar,  " & vbCrLf
        sql += " MAX(CAST(Imprimir AS int)) + 10 AS Imprimir, MAX(CAST(Historico AS int)) + 10 AS Historico, MAX(CAST(Seguridad AS int)) + 10 AS Seguridad " & vbCrLf
        sql += " FROM dbo.db_permisos_carpetas AS P GROUP BY CodUsuario, IdCuadro UNION ALL SELECT Grupo.CodUsuario AS codu, P.IdCuadro AS id, MAX(CAST(P.Ver AS int)) AS ver, " & vbCrLf
        sql += " MAX(CAST(P.Añadir AS int)) AS Añadir, MAX(CAST(P.Modificar AS int)) AS Modificar, MAX(CAST(P.Borrar AS int)) AS Borrar, " & vbCrLf
        sql += " MAX(CAST(P.Imprimir AS int)) AS Imprimir, MAX(CAST(P.Historico AS int)) AS Historico, MAX(CAST(P.Seguridad AS int)) AS Seguridad " & vbCrLf
        sql += " FROM dbo.db_permisos_carpetas AS P INNER JOIN DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo GROUP BY Grupo.CodUsuario, P.IdCuadro"
        crearfuncionoactualizarvista(conexion, "dbo.vista_permisos_global", sql)


        sql = "CREATE VIEW [dbo].[vista_permisos_global_objeto] " & vbCrLf
        sql += "AS" & vbCrLf
        sql += "SELECT     CodUsuario AS codu, IdObjeto, Objeto, MAX(CAST(Ver AS int)) + 10 AS ver, MAX(CAST(Añadir AS int)) + 10 AS Añadir, MAX(CAST(Modificar AS int)) " & vbCrLf
        sql += " + 10 AS Modificar, MAX(CAST(Borrar AS int)) + 10 AS Borrar, MAX(CAST(Imprimir AS int)) + 10 AS Imprimir, MAX(CAST(Historico AS int)) + 10 AS Historico, " & vbCrLf
        sql += "  MAX(CAST(Seguridad AS int)) + 10 AS Seguridad" & vbCrLf
        sql += "FROM dbo.DB_PERMISOS_OBJETO AS P" & vbCrLf
        sql += "GROUP BY CodUsuario, IdObjeto, Objeto" & vbCrLf
        sql += "UNION ALL" & vbCrLf
        sql += "SELECT     Grupo.CodUsuario AS codu, P.IdObjeto, P.Objeto, MAX(CAST(P.Ver AS int)) AS ver, MAX(CAST(P.Añadir AS int)) AS Añadir, MAX(CAST(P.Modificar AS int)) " & vbCrLf
        sql += " AS Modificar, MAX(CAST(P.Borrar AS int)) AS Borrar, MAX(CAST(P.Imprimir AS int)) AS Imprimir, MAX(CAST(P.Historico AS int)) AS Historico, " & vbCrLf
        sql += "MAX(CAST(P.Seguridad AS int)) AS Seguridad" & vbCrLf
        sql += "FROM  dbo.DB_PERMISOS_OBJETO AS P INNER JOIN" & vbCrLf
        sql += "DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "GROUP BY Grupo.CodUsuario, P.IdObjeto, P.Objeto" & vbCrLf

        crearfuncionoactualizarvista(conexion, "dbo.vista_permisos_global_objeto", sql)
    End Sub
    Public Sub crearvista_cuenta_carpeta(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = "CREATE VIEW [dbo].[vista_cuenta_carpetas] AS SELECT  ISNULL(COUNT(id), 0) AS hijos, IdPadre" & vbCrLf
        sql += "FROM  dbo.DB_CUADRO AS ct WHERE (Carpeta = N'S') GROUP BY IdPadre"
        crearfuncionoactualizarvista(conexion, "dbo.vista_cuenta_carpetas", sql)



        sql = " create VIEW [dbo].[vista_permiso_pre_carpetas] AS SELECT CodUsuario AS codu, IdCuadro AS id, MAX(CAST(Ver AS int)) + 10 AS ver "
        sql += " FROM dbo.DB_PERMISOS_carpetas AS P GROUP BY CodUsuario, IdCuadro UNION ALL SELECT Grupo.CodUsuario AS codu, P.IdCuadro AS id, "
        sql += " MAX(CAST(P.Ver AS int)) AS ver FROM dbo.DB_PERMISOS_carpetas AS P INNER JOIN DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo "
        sql += " GROUP BY Grupo.CodUsuario, P.IdCuadro "
        crearfuncionoactualizarvista(conexion, "dbo.vista_permiso_pre_carpetas", sql)


        sql = " CREATE VIEW [dbo].[vista_permiso_carpeta] AS SELECT codu, id, MAX(ver) AS ver FROM dbo.vista_permiso_pre_carpetas GROUP BY codu, id HAVING (MAX(ver) = 11 OR MAX(ver) = 1)"

        crearfuncionoactualizarvista(conexion, "dbo.vista_permiso_carpeta", sql)



    End Sub

    Public Sub crearfincionBuscarDato(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = "CREATE FUNCTION [dbo].[udf_BuscarDato] " & vbCrLf
        sql += "( " & vbCrLf
        sql += "@idcuadro numeric(18,0)," & vbCrLf
        sql += " @operador nvarchar(max), " & vbCrLf
        sql += "  @valor nvarchar(max) ," & vbCrLf
        sql += " @campo nvarchar(max) " & vbCrLf

        sql += " ) " & vbCrLf
        sql += "RETURNS nvarchar(max)" & vbCrLf
        sql += "AS " & vbCrLf
        sql += " BEGIN" & vbCrLf
        sql += " declare @id nvarchar(max) " & vbCrLf
        sql += " declare @resultado nvarchar(1) " & vbCrLf

        sql += "	IF @campo = ''" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "		if @operador = '=' " & vbCrLf
        sql += "			select @id = DB_DATOS.IdCuadro from DB_DATOS WHERE DB_DATOS.IdCuadro=@IdCuadro and DB_DATOS.valor = @valor  " & vbCrLf

        sql += "		if @operador = 'LIKE' " & vbCrLf
        sql += "			select @id = DB_DATOS.IdCuadro from DB_DATOS WHERE DB_DATOS.IdCuadro=@IdCuadro and DB_DATOS.valor like @valor  " & vbCrLf

        sql += "		if @operador = 'NOT LIKE' " & vbCrLf
        sql += "			select @id = DB_DATOS.IdCuadro from DB_DATOS WHERE DB_DATOS.IdCuadro=@IdCuadro and DB_DATOS.valor NOT like @valor  " & vbCrLf


        sql += "		if @operador = 'distinto' " & vbCrLf
        sql += "			 select @id = DB_DATOS.IdCuadro from DB_DATOS WHERE DB_DATOS.IdCuadro=@IdCuadro and DB_DATOS.valor <> @valor  " & vbCrLf

        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "		if @operador = '=' " & vbCrLf
        sql += "			select @id = DB_DATOS.IdCuadro from DB_DATOS WHERE DB_DATOS.IdCuadro=@IdCuadro and DB_DATOS.IdLSerie = @campo and DB_DATOS.valor = @valor " & vbCrLf

        sql += "		if @operador = 'LIKE' " & vbCrLf
        sql += "			select @id = DB_DATOS.IdCuadro from DB_DATOS WHERE DB_DATOS.IdCuadro=@IdCuadro and DB_DATOS.IdLSerie = @campo and DB_DATOS.valor like @valor " & vbCrLf

        sql += "		if @operador = 'NOT LIKE' " & vbCrLf
        sql += "			select @id = DB_DATOS.IdCuadro from DB_DATOS WHERE DB_DATOS.IdCuadro=@IdCuadro and DB_DATOS.IdLSerie = @campo and DB_DATOS.valor NOT like @valor " & vbCrLf


        sql += "		if @operador = 'distinto' " & vbCrLf
        sql += "			select @id = DB_DATOS.IdCuadro from DB_DATOS WHERE DB_DATOS.IdCuadro=@IdCuadro and DB_DATOS.IdLSerie = @campo and DB_DATOS.valor <> @valor " & vbCrLf

        sql += "End" & vbCrLf

        sql += " if @id is null " & vbCrLf
        sql += "	    set @resultado = '0'" & vbCrLf
        sql += " Else" & vbCrLf
        sql += " 	    set @resultado = '1'" & vbCrLf

        sql += "return @resultado " & vbCrLf

        sql += "   End"
        crearfuncionoactualizarfuncion(conexion, "dbo.udf_BuscarDato", sql)
    End Sub
    Public Sub crearfuncionf_generaarbol_hijos(ByVal conexion As ConexionSQL)
        Dim sql As String

        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol_hijos] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)," & vbCrLf
        sql += "	@limiteorden_min numeric(18,0)," & vbCrLf
        sql += "    @limiteorden_max numeric(18,0)" & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += "    Carpeta nvarchar(1)," & vbCrLf
        sql += "    Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "Carpeta nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "if @limiteorden_min = -1  and @limiteorden_max>-1" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "insert into @temp_q1" & vbCrLf
        '* 12/02/2012 INICIO
        'sql += "Select id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "Select c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        'sql += "from db_cuadro c" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        'sql += "where   c.idpadre =@nodoraiz and orden<@limiteorden_max and (dbo.udf_permiso_ver_nodorapido( C.ID, @usuario)=1 or" & vbCrLf
        'sql += "exists (select * from  db_cuadro H where  C.id = h.idpadre and" & vbCrLf
        'sql += "dbo.udf_permiso_ver_nodorapido( H.ID, @usuario)=1))" & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and orden<@limiteorden_max " & vbCrLf

        sql += "order by C.idpadre,C.orden desc" & vbCrLf
        sql += "     End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @limiteorden_min > -1 and @limiteorden_max= -1" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "	insert into @temp_q1" & vbCrLf
        'sql += "	Select  id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "	Select  c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        'sql += "	from db_cuadro c" & vbCrLf
        sql += "	from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        'sql += "	where   c.idpadre =@nodoraiz and orden>@limiteorden_min and (dbo.udf_permiso_ver_nodorapido( C.ID, @usuario)=1 or" & vbCrLf
        'sql += "	exists (select * from  db_cuadro H where  C.id = h.idpadre and" & vbCrLf
        'sql += "	dbo.udf_permiso_ver_nodorapido( H.ID, @usuario)=1))" & vbCrLf
        sql += "	where   c.idpadre =@nodoraiz and orden>@limiteorden_min " & vbCrLf
        sql += "	order by C.idpadre,C.orden desc" & vbCrLf
        sql += "End" & vbCrLf
        sql += "	else" & vbCrLf
        sql += "     begin" & vbCrLf
        sql += "	insert into @temp_q1" & vbCrLf
        'sql += "	Select  id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "	Select  c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        'sql += "	from db_cuadro c" & vbCrLf
        sql += "	from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        'sql += "	where   c.idpadre =@nodoraiz and orden>@limiteorden_min and orden<@limiteorden_max and (dbo.udf_permiso_ver_nodorapido( C.ID, @usuario)=1 or" & vbCrLf
        'sql += "	exists (select * from  db_cuadro H where  C.id = h.idpadre and" & vbCrLf
        'sql += "	dbo.udf_permiso_ver_nodorapido( H.ID, @usuario)=1))" & vbCrLf
        sql += "	where   c.idpadre =@nodoraiz and orden>@limiteorden_min and orden<@limiteorden_max " & vbCrLf
        '* 12/02/2012 FIN
        sql += "	order by C.idpadre,C.orden desc" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "insert @tablaresultado" & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos" & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre" & vbCrLf
        sql += "group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta" & vbCrLf
        sql += "        Return" & vbCrLf
        sql += "        End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_hijos", sql)
    End Sub
    'Public Sub crearfuncionf_generaarbol_hijos_objeto(ByVal conexion As ConexionSQL)
    '    Dim sql As String

    '    sql += "  CREATE FUNCTION  [dbo].[f_GeneraArbol_hijos_objeto] " & vbCrLf
    '    sql += " (" & vbCrLf
    '    sql += " -- Add the parameters for the function here" & vbCrLf
    '    sql += " @usuario nvarchar(20)," & vbCrLf
    '    sql += "  @nodoraiz numeric(18,0)," & vbCrLf
    '    sql += " @limiteorden_min numeric(18,0)," & vbCrLf
    '    sql += "   @limiteorden_max numeric(18,0)," & vbCrLf
    '    sql += "   @serie numeric(18,0), " & vbCrLf
    '    sql += "   @objeto nvarchar(100)" & vbCrLf
    '    sql += "   )" & vbCrLf
    '    sql += "      RETURNS" & vbCrLf
    '    sql += " @tablaresultado TABLE" & vbCrLf
    '    sql += " (" & vbCrLf
    '    sql += " 	ID numeric(18,0)," & vbCrLf
    '    sql += " 	Nombre nvarchar(300)," & vbCrLf
    '    sql += " 	IdPadre numeric(18,0)," & vbCrLf
    '    sql += " Orden numeric(18,0)," & vbCrLf
    '    sql += " IdSerie numeric(18,0)," & vbCrLf
    '    sql += "   Carpeta nvarchar(1)," & vbCrLf
    '    sql += "       Nhijos numeric(18, 0)" & vbCrLf
    '    sql += " )" & vbCrLf
    '    sql += " AS" & vbCrLf
    '    sql += "    BEGIN" & vbCrLf
    '    sql += " -- Fill the table variable with the rows for your result set" & vbCrLf
    '    sql += " DECLARE 	@temp_q1 TABLE" & vbCrLf
    '    sql += " (" & vbCrLf
    '    sql += " ID numeric(18,0)," & vbCrLf
    '    sql += " 	Nombre nvarchar(300)," & vbCrLf
    '    sql += " IdPadre numeric(18,0)," & vbCrLf
    '    sql += " Orden numeric(18,0)," & vbCrLf
    '    sql += " IdSerie numeric(18,0)," & vbCrLf
    '    sql += " Carpeta nvarchar(1)," & vbCrLf
    '    sql += "        objeto nvarchar(100)" & vbCrLf
    '    sql += " )" & vbCrLf
    '    sql += " if @limiteorden_min = -1  and @limiteorden_max>-1" & vbCrLf
    '    sql += " begin " & vbCrLf
    '    sql += " insert into @temp_q1" & vbCrLf
    '    sql += " Select c.id, Nombre,idpadre, orden, idserie, carpeta,objeto" & vbCrLf
    '    sql += " from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
    '    sql += " where   c.idpadre =@nodoraiz and orden<@limiteorden_max and idSerie = @serie and objeto = @objeto" & vbCrLf
    '    sql += " order by C.idpadre,C.orden desc" & vbCrLf
    '    sql += "      End" & vbCrLf
    '    sql += " else" & vbCrLf
    '    sql += " begin" & vbCrLf
    '    sql += " if @limiteorden_min > -1 and @limiteorden_max= -1" & vbCrLf
    '    sql += " begin" & vbCrLf
    '    sql += " 	insert into @temp_q1" & vbCrLf
    '    sql += " 	Select  c.id, Nombre,idpadre, orden, idserie, carpeta,objeto" & vbCrLf
    '    sql += " 	from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
    '    sql += " 	where   c.idpadre =@nodoraiz and orden>@limiteorden_min and idSerie = @serie and objeto = @objeto" & vbCrLf
    '    sql += " order by C.idpadre,C.orden desc" & vbCrLf
    '    sql += "       End" & vbCrLf
    '    sql += " else" & vbCrLf
    '    sql += "  begin" & vbCrLf
    '    sql += " insert into @temp_q1" & vbCrLf
    '    sql += " Select  c.id, Nombre,idpadre, orden, idserie, carpeta,objeto" & vbCrLf
    '    sql += " from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
    '    sql += " where   c.idpadre =@nodoraiz and orden>@limiteorden_min and orden<@limiteorden_max and idSerie = @serie and objeto = @objeto" & vbCrLf
    '    sql += " order by C.idpadre,C.orden desc" & vbCrLf
    '    sql += " End" & vbCrLf
    '    sql += " End" & vbCrLf
    '    sql += " insert @tablaresultado" & vbCrLf
    '    sql += " select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos" & vbCrLf
    '    sql += " from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre" & vbCrLf
    '    sql += " group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta" & vbCrLf
    '    sql += " Return" & vbCrLf
    '    sql += " End" & vbCrLf
    '    crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_hijos_objeto", sql)
    'End Sub
    Public Sub crearfuncionf_generaarbol_hijos_carpetas(ByVal conexion As ConexionSQL)
        Dim sql As String

        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol_hijos_carpetas] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)," & vbCrLf
        sql += "	@limiteorden_min numeric(18,0)," & vbCrLf
        sql += "    @limiteorden_max numeric(18,0)" & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += "    Carpeta nvarchar(1)," & vbCrLf
        sql += "    Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "Carpeta nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "if @limiteorden_min = -1  and @limiteorden_max>-1" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "insert into @temp_q1" & vbCrLf
        sql += "Select c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and orden<@limiteorden_max and c.carpeta = 'S'" & vbCrLf
        sql += "order by C.idpadre,C.orden desc" & vbCrLf
        sql += "     End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @limiteorden_min > -1 and @limiteorden_max= -1" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "	insert into @temp_q1" & vbCrLf
        sql += "	Select  c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "	from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        sql += "	where   c.idpadre =@nodoraiz and orden>@limiteorden_min and c.carpeta = 'S'" & vbCrLf
        sql += "	order by C.idpadre,C.orden desc" & vbCrLf
        sql += "End" & vbCrLf
        sql += "	else" & vbCrLf
        sql += "     begin" & vbCrLf
        sql += "	insert into @temp_q1" & vbCrLf
        sql += "	Select  c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "	from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        sql += "	where   c.idpadre =@nodoraiz and orden>@limiteorden_min and orden<@limiteorden_max and c.carpeta = 'S'" & vbCrLf
        sql += "	order by C.idpadre,C.orden desc" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "insert @tablaresultado" & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos" & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre and c.carpeta = 'S'" & vbCrLf
        sql += "group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta" & vbCrLf
        sql += "        Return" & vbCrLf
        sql += "        End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_hijos_carpetas", sql)
    End Sub
    Public Sub crearfuncionf_generaarbol_menos(ByVal conexion As ConexionSQL)
        Dim sql As String
        Dim NroAvance As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroAvance").ToString
        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol_menos] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)," & vbCrLf
        sql += "    	@limiteorden numeric(18,0)" & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += " Carpeta nvarchar(1)," & vbCrLf
        sql += " Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "Carpeta nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "insert into @temp_q1" & vbCrLf
        'sql += "Select top " & (NroAvance + 1).ToString & " id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "Select top " & (NroAvance + 1).ToString & " c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        'sql += "from db_cuadro c" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        'sql += "where   c.idpadre =@nodoraiz and orden<=@limiteorden and (dbo.udf_permiso_ver_nodorapido( C.ID, @usuario)=1 or" & vbCrLf
        'sql += "exists (select * from  db_cuadro H where  C.id = h.idpadre and" & vbCrLf
        'sql += "dbo.udf_permiso_ver_nodorapido( H.ID, @usuario)=1))" & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and orden<=@limiteorden " & vbCrLf
        sql += "order by C.idpadre,C.orden desc" & vbCrLf
        sql += "insert @tablaresultado" & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos" & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre" & vbCrLf
        sql += "group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta" & vbCrLf
        sql += "        Return" & vbCrLf
        sql += "        End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_menos", sql)
    End Sub
    Public Sub crearfuncionf_generaarbol_menos_carpetas(ByVal conexion As ConexionSQL)
        Dim sql As String
        Dim NroAvance As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroAvance").ToString
        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol_menos_carpetas] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)," & vbCrLf
        sql += "    	@limiteorden numeric(18,0)" & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += " Carpeta nvarchar(1)," & vbCrLf
        sql += " Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "Carpeta nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "insert into @temp_q1" & vbCrLf
        sql += "Select top " & (NroAvance + 1).ToString & " c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and orden<=@limiteorden and c.carpeta = 'S' " & vbCrLf
        sql += "order by C.idpadre,C.orden desc" & vbCrLf
        sql += "insert @tablaresultado" & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos" & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre and c.carpeta = 'S' " & vbCrLf
        sql += "group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta" & vbCrLf
        sql += "        Return" & vbCrLf
        sql += "        End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_menos_carpetas", sql)
    End Sub
    Public Sub crearfuncionf_generaarbol_mas(ByVal conexion As ConexionSQL)
        Dim sql As String
        Dim NroAvance As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroAvance").ToString
        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol_mas] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)," & vbCrLf
        sql += "    	@limiteorden numeric(18,0)" & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += " Carpeta nvarchar(1)," & vbCrLf
        sql += " Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "        Carpeta nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "insert into @temp_q1" & vbCrLf
        sql += "Select top " & (NroAvance + 1).ToString & " c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        'sql += "from db_cuadro c" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario" & vbCrLf
        'sql += "where   c.idpadre =@nodoraiz and orden>=@limiteorden and (dbo.udf_permiso_ver_nodorapido( C.ID, @usuario)=1 or" & vbCrLf
        'sql += "exists (select * from  db_cuadro H where  C.id = h.idpadre and" & vbCrLf
        'sql += "dbo.udf_permiso_ver_nodorapido( H.ID, @usuario)=1))" & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and orden>=@limiteorden " & vbCrLf
        sql += "order by C.idpadre,C.orden " & vbCrLf
        sql += "insert @tablaresultado" & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos" & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre" & vbCrLf
        sql += "group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta" & vbCrLf
        sql += "        Return" & vbCrLf
        sql += "        End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_mas", sql)
    End Sub
    Public Sub crearfuncionf_generaarbol_mas_carpetas(ByVal conexion As ConexionSQL)
        Dim sql As String
        Dim NroAvance As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroAvance").ToString
        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol_mas_carpetas] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)," & vbCrLf
        sql += "    	@limiteorden numeric(18,0)" & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += " Carpeta nvarchar(1)," & vbCrLf
        sql += " Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "        Carpeta nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "insert into @temp_q1" & vbCrLf
        sql += "Select top " & (NroAvance + 1).ToString & " c.id, Nombre,idpadre, orden, idserie, carpeta" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario " & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and orden>=@limiteorden and c.carpeta = 'S' " & vbCrLf
        sql += "order by C.idpadre,C.orden " & vbCrLf
        sql += "insert @tablaresultado" & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos" & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre and c.carpeta = 'S'" & vbCrLf
        sql += "group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta" & vbCrLf
        sql += "        Return" & vbCrLf
        sql += "        End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_mas_carpetas", sql)
    End Sub
   
   
    Public Sub crearfuncionf_generaarbol_carpetas(ByVal conexion As ConexionSQL)
        Dim sql As String
        Dim NroMaximoNodos As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroMaximoNodos").ToString
        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol_carpetas] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)" & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += " Carpeta nvarchar(1)," & vbCrLf
        sql += " Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "Carpeta nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "INSERT @temp_q1" & vbCrLf
        sql += "select ID, Nombre, IdPadre, Orden, IdSerie, Carpeta	" & vbCrLf
        sql += "from db_cuadro c" & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and  c.carpeta = 'S'" & vbCrLf
        sql += "insert @tablaresultado " & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta," & vbCrLf
        sql += "ISNULL((select hijos from vista_cuenta_carpetas v1 where v1.idpadre = t.id),0) as hijos" & vbCrLf
        sql += "from @temp_q1 as t  inner join dbo.vista_permiso_CARPETA v on t.id = v.id" & vbCrLf
        sql += "where v.codu=@usuario " & vbCrLf
        sql += "group by  t.idpadre, t.orden, t.id, t.Nombre, t.idserie, t.carpeta " & vbCrLf
        sql += "        Return" & vbCrLf
        sql += "        End"
        'sql += "INSERT @temp_q1" & vbCrLf
        'sql += "select top " & (NroMaximoNodos + 1).ToString & " c.ID, Nombre, IdPadre, Orden, IdSerie, Carpeta	" & vbCrLf
        'sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id" & vbCrLf
        'sql += "where   c.idpadre =@nodoraiz and  v.codu= @usuario and c.carpeta = 'S'" & vbCrLf
        'sql += "order by C.idpadre,C.orden" & vbCrLf
        'sql += "insert @tablaresultado " & vbCrLf
        'sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos " & vbCrLf
        'sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre and c.carpeta = 'S'" & vbCrLf
        'sql += "group by  t.idpadre, t.orden, t.id, t.Nombre, t.idserie, t.carpeta " & vbCrLf
        'sql += "        Return" & vbCrLf
        'sql += "        End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_carpetas", sql)
    End Sub
    'Public Sub crearfuncionf_generaarbol_mas_objeto(ByVal conexion As ConexionSQL)
    '    Dim sql As String
    '    Dim NroAvance As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroAvance").ToString
    '    sql = "      CREATE FUNCTION  [dbo].[f_GeneraArbol_mas_objeto]  " & vbCrLf
    '    sql += "("
    '    sql += "	-- Add the parameters for the function here  " & vbCrLf

    '    sql += "	@usuario nvarchar(20), " & vbCrLf
    '    sql += "   @nodoraiz numeric(18,0), " & vbCrLf
    '    sql += "	@limiteorden numeric(18,0), " & vbCrLf
    '    sql += "@serie numeric(18,0), " & vbCrLf
    '    sql += "@objeto nvarchar(100) " & vbCrLf
    '    sql += "   )"
    '    sql += "       RETURNS" & vbCrLf
    '    sql += "@tablaresultado TABLE " & vbCrLf
    '    sql += "( " & vbCrLf
    '    sql += "ID numeric(18, 0), " & vbCrLf
    '    sql += "Nombre nvarchar(300), " & vbCrLf
    '    sql += "IdPadre numeric(18, 0), " & vbCrLf
    '    sql += "Orden numeric(18, 0), " & vbCrLf
    '    sql += "IdSerie numeric(18, 0), " & vbCrLf
    '    sql += "Carpeta nvarchar(1), " & vbCrLf
    '    sql += " Nhijos numeric(18, 0) " & vbCrLf
    '    sql += ") " & vbCrLf
    '    sql += "AS " & vbCrLf
    '    sql += "BEGIN" & vbCrLf
    '    sql += "-- Fill the table variable with the rows for your result set " & vbCrLf
    '    sql += "DECLARE 	@temp_q1 TABLE " & vbCrLf
    '    sql += "( " & vbCrLf
    '    sql += "ID numeric(18, 0), " & vbCrLf
    '    sql += "Nombre nvarchar(300), " & vbCrLf
    '    sql += "IdPadre numeric(18, 0), " & vbCrLf
    '    sql += "Orden numeric(18, 0), " & vbCrLf
    '    sql += "IdSerie numeric(18, 0), " & vbCrLf
    '    sql += "Carpeta nvarchar(1), " & vbCrLf
    '    sql += "objeto nvarchar(100) " & vbCrLf
    '    sql += ") " & vbCrLf
    '    sql += "insert into @temp_q1 " & vbCrLf
    '    sql += "Select  top " & (NroAvance + 1).ToString & " c.id, Nombre,idpadre, orden, idserie, carpeta, objeto " & vbCrLf
    '    sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario " & vbCrLf
    '    sql += "where   c.idpadre =@nodoraiz and orden>=@limiteorden  and idSerie = @serie and objeto = @objeto " & vbCrLf
    '    sql += "order by C.idpadre,C.orden desc " & vbCrLf
    '    sql += "insert @tablaresultado " & vbCrLf
    '    sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos " & vbCrLf
    '    sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre " & vbCrLf
    '    sql += "group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta " & vbCrLf
    '    sql += "Return " & vbCrLf
    '    sql += "End " & vbCrLf
    '    crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_mas_objeto", sql)
    'End Sub
    'Public Sub crearfuncionf_generaarbol_menos_objeto(ByVal conexion As ConexionSQL)
    '    Dim sql As String
    '    Dim NroAvance As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroAvance").ToString
    '    sql = "      CREATE FUNCTION  [dbo].[f_GeneraArbol_menos_objeto]  " & vbCrLf
    '    sql += "("
    '    sql += "	-- Add the parameters for the function here  " & vbCrLf

    '    sql += "	@usuario nvarchar(20), " & vbCrLf
    '    sql += "   @nodoraiz numeric(18,0), " & vbCrLf
    '    sql += "	@limiteorden numeric(18,0), " & vbCrLf
    '    sql += "@serie numeric(18,0), " & vbCrLf
    '    sql += "@objeto nvarchar(100) " & vbCrLf
    '    sql += "   )"
    '    sql += "       RETURNS" & vbCrLf
    '    sql += "@tablaresultado TABLE " & vbCrLf
    '    sql += "( " & vbCrLf
    '    sql += "ID numeric(18, 0), " & vbCrLf
    '    sql += "Nombre nvarchar(300), " & vbCrLf
    '    sql += "IdPadre numeric(18, 0), " & vbCrLf
    '    sql += "Orden numeric(18, 0), " & vbCrLf
    '    sql += "IdSerie numeric(18, 0), " & vbCrLf
    '    sql += "Carpeta nvarchar(1), " & vbCrLf
    '    sql += " Nhijos numeric(18, 0) " & vbCrLf
    '    sql += ") " & vbCrLf
    '    sql += "AS " & vbCrLf
    '    sql += "BEGIN" & vbCrLf
    '    sql += "-- Fill the table variable with the rows for your result set " & vbCrLf
    '    sql += "DECLARE 	@temp_q1 TABLE " & vbCrLf
    '    sql += "( " & vbCrLf
    '    sql += "ID numeric(18, 0), " & vbCrLf
    '    sql += "Nombre nvarchar(300), " & vbCrLf
    '    sql += "IdPadre numeric(18, 0), " & vbCrLf
    '    sql += "Orden numeric(18, 0), " & vbCrLf
    '    sql += "IdSerie numeric(18, 0), " & vbCrLf
    '    sql += "Carpeta nvarchar(1), " & vbCrLf
    '    sql += "objeto nvarchar(100) " & vbCrLf
    '    sql += ") " & vbCrLf
    '    sql += "insert into @temp_q1 " & vbCrLf
    '    sql += "Select  top " & (NroAvance + 1).ToString & " c.id, Nombre,idpadre, orden, idserie, carpeta, objeto " & vbCrLf
    '    sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id and  v.codu= @usuario " & vbCrLf
    '    sql += "where   c.idpadre =@nodoraiz and orden<=@limiteorden  and idSerie = @serie and objeto = @objeto " & vbCrLf
    '    sql += "order by C.idpadre,C.orden desc " & vbCrLf
    '    sql += "insert @tablaresultado " & vbCrLf
    '    sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos " & vbCrLf
    '    sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre " & vbCrLf
    '    sql += "group by t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta " & vbCrLf
    '    sql += "Return " & vbCrLf
    '    sql += "End " & vbCrLf
    '    crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_menos_objeto", sql)
    'End Sub
    Public Sub crearfuncionf_generaarbol_objeto(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = " CREATE FUNCTION  [dbo].[f_GeneraArbol_objeto] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "-- Add the parameters for the function here" & vbCrLf
        sql += "@usuario nvarchar(20)," & vbCrLf
        sql += "@nodoraiz numeric(18,0)," & vbCrLf
        sql += "@serie numeric(18,0)," & vbCrLf
        sql += "@objeto nvarchar(100)," & vbCrLf
        sql += "@idlserie numeric(18,0)," & vbCrLf
        sql += "@tipocampo nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18, 0), " & vbCrLf
        sql += "Nombre nvarchar(300), " & vbCrLf
        sql += "IdPadre numeric(18, 0), " & vbCrLf
        sql += "Orden numeric(18, 0), " & vbCrLf
        sql += "IdSerie numeric(18, 0), " & vbCrLf
        sql += "Carpeta nvarchar(1), " & vbCrLf
        sql += "Nhijos numeric(18, 0)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18, 0), " & vbCrLf
        sql += "Nombre nvarchar(300), " & vbCrLf
        sql += "IdPadre numeric(18, 0), " & vbCrLf
        sql += "Orden numeric(18, 0), " & vbCrLf
        sql += "IdSerie numeric(18, 0), " & vbCrLf
        sql += "Carpeta nvarchar(1), " & vbCrLf
        sql += "objeto nvarchar(100)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "INSERT @temp_q1" & vbCrLf
        sql += "select  c.ID, Nombre, IdPadre, Orden, IdSerie, Carpeta	, objeto" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id" & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and  v.codu= @usuario and idSerie = @serie and objeto = @objeto" & vbCrLf
        sql += "order by C.idpadre,C.orden" & vbCrLf
        sql += "DECLARE 	@temp_q2 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "idcuadro numeric(18, 0), " & vbCrLf
        sql += "valorcampo nvarchar(max)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "if @tipocampo = 'F'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "INSERT @temp_q2" & vbCrLf
        sql += "select idcuadro, CONVERT (datetime, case when valor is not null then valor else '01/01/1990' end, 103) as valorcampo" & vbCrLf
        sql += "FROM DB_DATOS WHERE (IdLSerie = @idlserie) " & vbCrLf
        sql += "End" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "INSERT @temp_q2" & vbCrLf
        sql += "select idcuadro, valor as valorcampo" & vbCrLf
        sql += "FROM DB_DATOS WHERE (IdLSerie = @idlserie) " & vbCrLf
        sql += "End" & vbCrLf
        sql += "if @tipocampo = 'F'" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "insert @tablaresultado " & vbCrLf
        sql += "SELECT  t.id, t.Nombre, t.idpadre, t.orden, t.idserie, t.carpeta, ISNULL(COUNT(c.id), 0) AS hijos" & vbCrLf
        sql += "FROM  @temp_q1 AS t LEFT OUTER JOIN" & vbCrLf
        sql += "@temp_q2 AS d ON d.IdCuadro = t.id LEFT OUTER JOIN" & vbCrLf
        sql += "DB_CUADRO AS c ON t.id = c.IdPadre" & vbCrLf
        sql += "GROUP BY t.idpadre, t.orden, t.id, t.Nombre, t.idserie, t.carpeta, d.valorcampo" & vbCrLf
        sql += "ORDER BY convert(datetime,d.valorcampo,103) desc" & vbCrLf
        sql += "End" & vbCrLf
        sql += "ELSE" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "insert @tablaresultado " & vbCrLf
        sql += "SELECT  t.id, t.Nombre, t.idpadre, t.orden, t.idserie, t.carpeta, ISNULL(COUNT(c.id), 0) AS hijos" & vbCrLf
        sql += "FROM  @temp_q1 AS t LEFT OUTER JOIN" & vbCrLf
        sql += "@temp_q2 AS d ON d.IdCuadro = t.id LEFT OUTER JOIN" & vbCrLf
        sql += "DB_CUADRO AS c ON t.id = c.IdPadre" & vbCrLf
        sql += "GROUP BY t.idpadre, t.orden, t.id, t.Nombre, t.idserie, t.carpeta, d.valorcampo" & vbCrLf
        sql += "ORDER BY d.valorcampo " & vbCrLf
        sql += "End" & vbCrLf
        sql += "Return " & vbCrLf
        sql += "End" & vbCrLf
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_objeto", sql)
    End Sub
    Public Sub crearfuncionf_generaarbol_objeto_old(ByVal conexion As ConexionSQL)
        Dim sql As String
        Dim NroMaximoNodos As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroMaximoNodos").ToString
        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol_objeto_old] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)," & vbCrLf
        sql += "	@serie numeric(18,0)," & vbCrLf
        sql += "    @objeto nvarchar(100), " & vbCrLf
        sql += "    @idlserie numeric(18,0)," & vbCrLf
        sql += "    @tipocampo nvarchar(1) " & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += " Carpeta nvarchar(1)," & vbCrLf
        sql += " Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "Carpeta nvarchar(1)," & vbCrLf
        sql += "objeto nvarchar(100)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "INSERT @temp_q1" & vbCrLf
        sql += "select top " & (NroMaximoNodos + 1).ToString & " c.ID, Nombre, IdPadre, Orden, IdSerie, Carpeta, objeto	" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id" & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and  v.codu= @usuario and idSerie = @serie and objeto = @objeto " & vbCrLf
        sql += "order by C.idpadre,C.orden" & vbCrLf

        sql += "if @tipocampo = 'F'" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "insert @tablaresultado " & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta,   isnull(count(c.id),0) as hijos " & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre " & vbCrLf
        sql += "group by  t.idpadre, t.orden, t.id, t.Nombre, t.idserie, t.carpeta " & vbCrLf
        sql += "ORDER BY (Select LEFT(CONVERT(datetime, valor, 103), 11)  FROM DB_DATOS WHERE (IdLSerie = @idlserie) AND (IdCuadro = t.id)) desc" & vbCrLf
        sql += "End" & vbCrLf
        sql += "ELSE" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "insert @tablaresultado " & vbCrLf
        sql += "select  t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta,   isnull(count(c.id),0) as hijos " & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre " & vbCrLf
        sql += "group by  t.idpadre, t.orden, t.id, t.Nombre, t.idserie, t.carpeta " & vbCrLf
        sql += "order by (Select valor  FROM DB_DATOS WHERE (IdLSerie = @idlserie) AND (IdCuadro = t.id)) " & vbCrLf
        sql += "End" & vbCrLf

        sql += "Return" & vbCrLf
        sql += "End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol_objeto_old", sql)
    End Sub
    Public Sub crearfuncionf_generaarbol(ByVal conexion As ConexionSQL)
        Dim sql As String
        Dim NroMaximoNodos As Integer = System.Configuration.ConfigurationManager.AppSettings("NumeroMaximoNodos").ToString
        sql = "       CREATE FUNCTION  [dbo].[f_GeneraArbol] " & vbCrLf
        sql += "(" & vbCrLf
        sql += "	-- Add the parameters for the function here" & vbCrLf
        sql += "" & vbCrLf
        sql += "	@usuario nvarchar(20)," & vbCrLf
        sql += "    @nodoraiz numeric(18,0)" & vbCrLf
        sql += "    )" & vbCrLf
        sql += "RETURNS" & vbCrLf
        sql += "@tablaresultado TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "	ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "	IdPadre numeric(18,0)," & vbCrLf
        sql += "	Orden numeric(18,0)," & vbCrLf
        sql += "	IdSerie numeric(18,0)," & vbCrLf
        sql += " Carpeta nvarchar(1)," & vbCrLf
        sql += " Nhijos numeric(18,0)" & vbCrLf
        sql += " )" & vbCrLf
        sql += "AS" & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "-- Fill the table variable with the rows for your result set" & vbCrLf
        sql += "DECLARE 	@temp_q1 TABLE" & vbCrLf
        sql += "(" & vbCrLf
        sql += "ID numeric(18,0)," & vbCrLf
        sql += "	Nombre nvarchar(300)," & vbCrLf
        sql += "IdPadre numeric(18,0)," & vbCrLf
        sql += "Orden numeric(18,0)," & vbCrLf
        sql += "IdSerie numeric(18,0)," & vbCrLf
        sql += "Carpeta nvarchar(1)" & vbCrLf
        sql += ")" & vbCrLf
        sql += "INSERT @temp_q1" & vbCrLf
        'comprobar "Select top 201 c.id, Nombre,idpadre, orden, idserie, carpeta"
        sql += "select top " & (NroMaximoNodos + 1).ToString & " c.ID, Nombre, IdPadre, Orden, IdSerie, Carpeta	" & vbCrLf

        '12/02/2012 *** INICIO
        'sql += "from db_Cuadro c" & vbCrLf
        sql += "from db_cuadro c inner join dbo.vista_permiso v on c.id = v.id" & vbCrLf
        '12/02/2012 *** FIN
        '12/02/2012 *** INICIO
        'sql += "where   c.idpadre =@nodoraiz and (dbo.udf_permiso_ver_nodorapido( C.ID, @usuario)=1 or " & vbCrLf
        'sql += "exists (select * from  db_cuadro H where  C.id = h.idpadre and " & vbCrLf
        'sql += "dbo.udf_permiso_ver_nodorapido( H.ID, @usuario)=1)) " & vbCrLf
        sql += "where   c.idpadre =@nodoraiz and  v.codu= @usuario" & vbCrLf
        '12/02/2012 *** FIN

        sql += "order by C.idpadre,C.orden" & vbCrLf
        sql += "insert @tablaresultado " & vbCrLf
        sql += "select t.id, t.Nombre,t.idpadre,t.orden,t.idserie,t.carpeta, isnull(count(c.id),0) as hijos " & vbCrLf
        sql += "from @temp_q1 as t left outer join db_cuadro c on t.id = c.idpadre " & vbCrLf
        sql += "group by  t.idpadre, t.orden, t.id, t.Nombre, t.idserie, t.carpeta " & vbCrLf
        sql += "        Return" & vbCrLf
        sql += "        End"
        crearfuncionoactualizarfuncion(conexion, "dbo.f_GeneraArbol", sql)
    End Sub
    Public Sub crearfuncionPermiso_objeto(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = " create FUNCTION [dbo].[udf_Permiso_Objeto] " & vbCrLf
        sql += "( " & vbCrLf
        sql += " @idobjeto nvarchar(100), " & vbCrLf
        sql += " @objeto nvarchar(100)," & vbCrLf
        sql += "  @usuario nvarchar(max)," & vbCrLf
        sql += " @TipoPermiso nvarchar(max) " & vbCrLf
        sql += " ) " & vbCrLf
        sql += " RETURNS nvarchar(max)" & vbCrLf
        sql += " AS " & vbCrLf
        sql += "  BEGIN" & vbCrLf
        sql += "declare @cuenta int" & vbCrLf
        sql += "declare @permiso nvarchar(1)" & vbCrLf
        sql += "declare @resultado nvarchar(1) " & vbCrLf
        sql += "		select @cuenta = count(*) " & vbCrLf
        sql += " FROM DB_PERMISOS_OBJETO " & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) and (Objeto =@objeto)" & vbCrLf
        sql += "if @TipoPermiso = 'Ver'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "		if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "		select @resultado = isnull(Ver,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) and (Objeto =@objeto)" & vbCrLf
        sql += "End" & vbCrLf
        sql += "		else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO AS P  INNER JOIN" & vbCrLf
        sql += "  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and (Objeto =@objeto) and ( P.[Ver] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += "    End" & vbCrLf
        sql += "    End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "    begin" & vbCrLf
        sql += "if @TipoPermiso = 'Añadir'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @resultado = isnull(Añadir,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO " & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) and (Objeto =@objeto)" & vbCrLf
        sql += "End" & vbCrLf
        sql += "				else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO AS P  INNER JOIN" & vbCrLf
        sql += " DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and (Objeto =@objeto) and ( P.[Añadir] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += " End" & vbCrLf
        sql += " End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @TipoPermiso='Modificar'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @resultado = isnull(Modificar,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) and (Objeto =@objeto)" & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO AS P  INNER JOIN" & vbCrLf
        sql += "DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and (Objeto =@objeto) and ( P.[Modificar] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += " Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += "   End" & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @TipoPermiso='Borrar'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @resultado = isnull(Borrar,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) and (Objeto =@objeto)" & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO AS P  INNER JOIN" & vbCrLf
        sql += "DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and (Objeto =@objeto) and ( P.[Borrar] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += " End" & vbCrLf
        sql += " End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "	if @TipoPermiso='Imprimir'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "    begin " & vbCrLf
        sql += "select @resultado = isnull(Imprimir,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) and (Objeto =@objeto)" & vbCrLf
        sql += " End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO AS P  INNER JOIN" & vbCrLf
        sql += "  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and (Objeto =@objeto) and ( P.[Imprimir] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += " End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "  begin" & vbCrLf
        sql += "	if @TipoPermiso='Historico'" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "  begin " & vbCrLf
        sql += "select @resultado = isnull(Historico,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) and (Objeto =@objeto)" & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin " & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO AS P  INNER JOIN" & vbCrLf
        sql += " DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and (Objeto =@objeto) and ( P.[Historico] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += "   End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @TipoPermiso='Seguridad'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @resultado = isnull(Seguridad,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) and (Objeto =@objeto)" & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_OBJETO AS P  INNER JOIN" & vbCrLf
        sql += "	  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and (Objeto =@objeto) and ( P.[Seguridad] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "	set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += "End" & vbCrLf
        sql += " End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "set @resultado='0'" & vbCrLf
        sql += " End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "   End" & vbCrLf
        sql += " End" & vbCrLf
        sql += "   End" & vbCrLf
        sql += "   End" & vbCrLf
        sql += " return @resultado " & vbCrLf
        sql += "   End"
        crearfuncionoactualizarfuncion(conexion, " dbo.udf_Permiso_Objeto", sql)
    End Sub
    Public Sub crearfuncionPermiso_IDobjeto(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = " create FUNCTION [dbo].[udf_Permiso_IDObjeto] " & vbCrLf
        sql += "( " & vbCrLf
        sql += " @idobjeto nvarchar(100), " & vbCrLf
        sql += "  @usuario nvarchar(max)," & vbCrLf
        sql += " @TipoPermiso nvarchar(max) " & vbCrLf
        sql += " ) " & vbCrLf
        sql += " RETURNS nvarchar(max)" & vbCrLf
        sql += " AS " & vbCrLf
        sql += "  BEGIN" & vbCrLf
        sql += "declare @cuenta int" & vbCrLf
        sql += "declare @permiso nvarchar(1)" & vbCrLf
        sql += "declare @resultado nvarchar(1) " & vbCrLf
        sql += "		select @cuenta = count(*) " & vbCrLf
        sql += " FROM DB_PERMISOS_IDOBJETO " & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) " & vbCrLf
        sql += "if @TipoPermiso = 'Ver'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "		if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "		select @resultado = isnull(Ver,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) " & vbCrLf
        sql += "End" & vbCrLf
        sql += "		else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO AS P  INNER JOIN" & vbCrLf
        sql += "  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto)  and ( P.[Ver] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += "    End" & vbCrLf
        sql += "    End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "    begin" & vbCrLf
        sql += "if @TipoPermiso = 'Añadir'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @resultado = isnull(Añadir,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO " & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) " & vbCrLf
        sql += "End" & vbCrLf
        sql += "				else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO AS P  INNER JOIN" & vbCrLf
        sql += " DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto)  and ( P.[Añadir] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += " End" & vbCrLf
        sql += " End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @TipoPermiso='Modificar'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @resultado = isnull(Modificar,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) " & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO AS P  INNER JOIN" & vbCrLf
        sql += "DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto)  and ( P.[Modificar] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += " Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += "   End" & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @TipoPermiso='Borrar'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @resultado = isnull(Borrar,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) " & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO AS P  INNER JOIN" & vbCrLf
        sql += "DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and ( P.[Borrar] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += " End" & vbCrLf
        sql += " End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "	if @TipoPermiso='Imprimir'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "    begin " & vbCrLf
        sql += "select @resultado = isnull(Imprimir,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) " & vbCrLf
        sql += " End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO AS P  INNER JOIN" & vbCrLf
        sql += "  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto)  and ( P.[Imprimir] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += " End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "  begin" & vbCrLf
        sql += "	if @TipoPermiso='Historico'" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "  begin " & vbCrLf
        sql += "select @resultado = isnull(Historico,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) " & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin " & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO AS P  INNER JOIN" & vbCrLf
        sql += " DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto)  and ( P.[Historico] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += "   End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "if @TipoPermiso='Seguridad'" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "if @cuenta>0" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @resultado = isnull(Seguridad,'0')" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO" & vbCrLf
        sql += "WHERE (CodUsuario =@usuario ) AND (IdObjeto =@idobjeto) " & vbCrLf
        sql += "End" & vbCrLf
        sql += "else" & vbCrLf
        sql += "begin" & vbCrLf
        sql += "select @cuenta= Count(*)" & vbCrLf
        sql += "FROM DB_PERMISOS_IDOBJETO AS P  INNER JOIN" & vbCrLf
        sql += "	  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "WHERE (IdObjeto =@idobjeto) and ( P.[Seguridad] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "if @cuenta = 0        	" & vbCrLf
        sql += "	set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "set @resultado = '1'" & vbCrLf
        sql += "End" & vbCrLf
        sql += " End" & vbCrLf
        sql += "else" & vbCrLf
        sql += " begin" & vbCrLf
        sql += "set @resultado='0'" & vbCrLf
        sql += " End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "  End" & vbCrLf
        sql += "   End" & vbCrLf
        sql += " End" & vbCrLf
        sql += "   End" & vbCrLf
        sql += "   End" & vbCrLf
        sql += " return @resultado " & vbCrLf
        sql += "   End"
        crearfuncionoactualizarfuncion(conexion, " dbo.udf_Permiso_IDObjeto", sql)
    End Sub
    Public Sub crearfuncionPermiso(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = " CREATE FUNCTION [dbo].[udf_Permiso] " & vbCrLf
        sql += "       ( " & vbCrLf
        sql += "         @carpeta nvarchar(max), " & vbCrLf
        sql += "         @usuario nvarchar(max)," & vbCrLf
        sql += "		 @TipoPermiso nvarchar(max) " & vbCrLf
        sql += "        ) " & vbCrLf
        sql += "        RETURNS nvarchar(max) " & vbCrLf
        sql += "        AS " & vbCrLf
        sql += "        BEGIN " & vbCrLf
        sql += "        declare @cuenta int" & vbCrLf
        sql += "		declare @permiso nvarchar(1)" & vbCrLf
        sql += "        declare @resultado nvarchar(1) " & vbCrLf
        sql += "		--Vemos si existe alguna linea de permiso para ese usuario" & vbCrLf
        sql += "		select @cuenta = count(*) " & vbCrLf
        sql += "		FROM [DB_PERMISOS]" & vbCrLf
        sql += "		WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "		if @TipoPermiso = 'Ver'" & vbCrLf
        sql += "		begin" & vbCrLf
        sql += "				if @cuenta>0" & vbCrLf
        sql += "					begin " & vbCrLf
        sql += "								select @resultado = isnull(Ver,'0')" & vbCrLf
        sql += "								FROM [DB_PERMISOS]" & vbCrLf
        sql += "								WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "					end " & vbCrLf
        sql += "				else" & vbCrLf
        sql += "					begin" & vbCrLf
        sql += "								select @cuenta= Count(*)" & vbCrLf
        sql += "								FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "										  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "								WHERE (P.IdCuadro = @carpeta) and ( P.[Ver] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "									if @cuenta = 0        	" & vbCrLf
        sql += "										set @resultado = '0'" & vbCrLf
        sql += "									Else" & vbCrLf
        sql += "										set @resultado = '1'" & vbCrLf
        sql += "					end" & vbCrLf
        sql += "		end	" & vbCrLf
        sql += "        else" & vbCrLf
        sql += "		begin" & vbCrLf
        sql += "					if @TipoPermiso = 'Añadir'" & vbCrLf
        sql += "					begin			" & vbCrLf
        sql += "						if @cuenta>0" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "										select @resultado = isnull(Añadir,'0')" & vbCrLf
        sql += "										FROM [DB_PERMISOS]" & vbCrLf
        sql += "										WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "							end " & vbCrLf
        sql += "						else" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "										select @cuenta= Count(*)" & vbCrLf
        sql += "										FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "												  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "										WHERE (P.IdCuadro = @carpeta) and ( P.[Añadir] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "											if @cuenta = 0        	" & vbCrLf
        sql += "												set @resultado = '0'" & vbCrLf
        sql += "											Else" & vbCrLf
        sql += "												set @resultado = '1'" & vbCrLf
        sql += "							end" & vbCrLf
        sql += "					end" & vbCrLf
        sql += "					else" & vbCrLf
        sql += "					begin" & vbCrLf
        sql += "						if @TipoPermiso='Modificar'" & vbCrLf
        sql += "						begin" & vbCrLf
        sql += "						if @cuenta>0" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "										select @resultado = isnull(Modificar,'0')" & vbCrLf
        sql += "										FROM [DB_PERMISOS]" & vbCrLf
        sql += "										WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "							end " & vbCrLf
        sql += "						else" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "										select @cuenta= Count(*)" & vbCrLf
        sql += "										FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "												  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "										WHERE (P.IdCuadro = @carpeta) and ( P.[Modificar] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "											if @cuenta = 0        	" & vbCrLf
        sql += "												set @resultado = '0'" & vbCrLf
        sql += "											Else" & vbCrLf
        sql += "												set @resultado = '1'" & vbCrLf
        sql += "							end" & vbCrLf
        sql += "						end" & vbCrLf
        sql += "						else" & vbCrLf
        sql += "						begin" & vbCrLf
        sql += "							if @TipoPermiso='Borrar'" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "								if @cuenta>0" & vbCrLf
        sql += "									begin" & vbCrLf
        sql += "												select @resultado = isnull(Borrar,'0')" & vbCrLf
        sql += "												FROM [DB_PERMISOS]" & vbCrLf
        sql += "												WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "									end " & vbCrLf
        sql += "								else" & vbCrLf
        sql += "									begin" & vbCrLf
        sql += "												select @cuenta= Count(*)" & vbCrLf
        sql += "												FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "														  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "												WHERE (P.IdCuadro = @carpeta) and ( P.[Borrar] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "													if @cuenta = 0        	" & vbCrLf
        sql += "														set @resultado = '0'" & vbCrLf
        sql += "													Else" & vbCrLf
        sql += "														set @resultado = '1'" & vbCrLf
        sql += "									end" & vbCrLf
        sql += "							end" & vbCrLf
        sql += "							else" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "								if @TipoPermiso='Imprimir'" & vbCrLf
        sql += "								begin" & vbCrLf
        sql += "								if @cuenta>0" & vbCrLf
        sql += "										begin" & vbCrLf
        sql += "													select @resultado = isnull(Imprimir,'0')" & vbCrLf
        sql += "													FROM [DB_PERMISOS]" & vbCrLf
        sql += "													WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta) " & vbCrLf
        sql += "										end " & vbCrLf
        sql += "									else" & vbCrLf
        sql += "										begin" & vbCrLf
        sql += "													select @cuenta= Count(*)" & vbCrLf
        sql += "													FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "															  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "													WHERE (P.IdCuadro = @carpeta) and ( P.[Imprimir] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "														if @cuenta = 0        	" & vbCrLf
        sql += "															set @resultado = '0'" & vbCrLf
        sql += "														Else" & vbCrLf
        sql += "															set @resultado = '1'" & vbCrLf
        sql += "										end" & vbCrLf
        sql += "								end" & vbCrLf
        sql += "								else" & vbCrLf
        sql += "								begin" & vbCrLf
        sql += "									if @TipoPermiso='Historico'" & vbCrLf
        sql += "									begin" & vbCrLf
        sql += "										if @cuenta>0" & vbCrLf
        sql += "											begin" & vbCrLf
        sql += "														select @resultado = isnull(Historico,'0')" & vbCrLf
        sql += "														FROM [DB_PERMISOS]" & vbCrLf
        sql += "														WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta) " & vbCrLf
        sql += "											end " & vbCrLf
        sql += "										else" & vbCrLf
        sql += "											begin" & vbCrLf
        sql += "														select @cuenta= Count(*)" & vbCrLf
        sql += "														FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "																  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "														WHERE (P.IdCuadro = @carpeta) and ( P.[Historico] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "															if @cuenta = 0        	" & vbCrLf
        sql += "																set @resultado = '0'" & vbCrLf
        sql += "															Else" & vbCrLf
        sql += "																set @resultado = '1'" & vbCrLf
        sql += "											end" & vbCrLf
        sql += "									end" & vbCrLf
        sql += "									else" & vbCrLf
        sql += "									begin	" & vbCrLf
        sql += "											if @TipoPermiso='Seguridad'" & vbCrLf
        sql += "											begin" & vbCrLf
        sql += "												if @cuenta>0" & vbCrLf
        sql += "													begin" & vbCrLf
        sql += "																select @resultado = isnull(Seguridad,'0')" & vbCrLf
        sql += "																FROM [DB_PERMISOS]" & vbCrLf
        sql += "																WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta) " & vbCrLf
        sql += "													end " & vbCrLf
        sql += "												else" & vbCrLf
        sql += "													begin" & vbCrLf
        sql += "																select @cuenta= Count(*)" & vbCrLf
        sql += "																FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "																		  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "																WHERE (P.IdCuadro = @carpeta) and ( P.[Seguridad] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "																	if @cuenta = 0        	" & vbCrLf
        sql += "																		set @resultado = '0'" & vbCrLf
        sql += "																	Else" & vbCrLf
        sql += "																		set @resultado = '1'" & vbCrLf
        sql += "													end" & vbCrLf
        sql += "											end" & vbCrLf
        sql += "											else" & vbCrLf
        sql += "											begin" & vbCrLf
        sql += "													set @resultado='0'" & vbCrLf
        sql += "											end" & vbCrLf
        sql += "									end" & vbCrLf
        sql += "								end" & vbCrLf
        sql += "							end" & vbCrLf
        sql += "						end" & vbCrLf
        sql += "					end" & vbCrLf
        sql += "		end " & vbCrLf
        sql += "        return @resultado " & vbCrLf
        sql += "        End "

        crearfuncionoactualizarfuncion(conexion, " dbo.udf_Permiso_Carpeta", sql)
    End Sub
    Public Sub crearfuncionPermiso_Carpeta(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = " CREATE FUNCTION [dbo].[udf_Permiso_Carpeta] " & vbCrLf
        sql += "       ( " & vbCrLf
        sql += "         @carpeta nvarchar(max), " & vbCrLf
        sql += "         @usuario nvarchar(max)," & vbCrLf
        sql += "		 @TipoPermiso nvarchar(max) " & vbCrLf
        sql += "        ) " & vbCrLf
        sql += "        RETURNS nvarchar(max) " & vbCrLf
        sql += "        AS " & vbCrLf
        sql += "        BEGIN " & vbCrLf
        sql += "        declare @cuenta int" & vbCrLf
        sql += "		declare @permiso nvarchar(1)" & vbCrLf
        sql += "        declare @resultado nvarchar(1) " & vbCrLf
        sql += "		--Vemos si existe alguna linea de permiso para ese usuario" & vbCrLf
        sql += "		select @cuenta = count(*) " & vbCrLf
        sql += "		FROM [DB_PERMISOS_CARPETAS]" & vbCrLf
        sql += "		WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "		if @TipoPermiso = 'Ver'" & vbCrLf
        sql += "		begin" & vbCrLf
        sql += "				if @cuenta>0" & vbCrLf
        sql += "					begin " & vbCrLf
        sql += "								select @resultado = isnull(Ver,'0')" & vbCrLf
        sql += "								FROM [DB_PERMISOS_CARPETAS]" & vbCrLf
        sql += "								WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "					end " & vbCrLf
        sql += "				else" & vbCrLf
        sql += "					begin" & vbCrLf
        sql += "								select @cuenta= Count(*)" & vbCrLf
        sql += "								FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "										  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "								WHERE (P.IdCuadro = @carpeta) and ( P.[Ver] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "									if @cuenta = 0        	" & vbCrLf
        sql += "										set @resultado = '0'" & vbCrLf
        sql += "									Else" & vbCrLf
        sql += "										set @resultado = '1'" & vbCrLf
        sql += "					end" & vbCrLf
        sql += "		end	" & vbCrLf
        sql += "        else" & vbCrLf
        sql += "		begin" & vbCrLf
        sql += "					if @TipoPermiso = 'Añadir'" & vbCrLf
        sql += "					begin			" & vbCrLf
        sql += "						if @cuenta>0" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "										select @resultado = isnull(Añadir,'0')" & vbCrLf
        sql += "										FROM [DB_PERMISOS_CARPETAS]" & vbCrLf
        sql += "										WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "							end " & vbCrLf
        sql += "						else" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "										select @cuenta= Count(*)" & vbCrLf
        sql += "										FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "												  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "										WHERE (P.IdCuadro = @carpeta) and ( P.[Añadir] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "											if @cuenta = 0        	" & vbCrLf
        sql += "												set @resultado = '0'" & vbCrLf
        sql += "											Else" & vbCrLf
        sql += "												set @resultado = '1'" & vbCrLf
        sql += "							end" & vbCrLf
        sql += "					end" & vbCrLf
        sql += "					else" & vbCrLf
        sql += "					begin" & vbCrLf
        sql += "						if @TipoPermiso='Modificar'" & vbCrLf
        sql += "						begin" & vbCrLf
        sql += "						if @cuenta>0" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "										select @resultado = isnull(Modificar,'0')" & vbCrLf
        sql += "										FROM [DB_PERMISOS_CARPETAS]" & vbCrLf
        sql += "										WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "							end " & vbCrLf
        sql += "						else" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "										select @cuenta= Count(*)" & vbCrLf
        sql += "										FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "												  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "										WHERE (P.IdCuadro = @carpeta) and ( P.[Modificar] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "											if @cuenta = 0        	" & vbCrLf
        sql += "												set @resultado = '0'" & vbCrLf
        sql += "											Else" & vbCrLf
        sql += "												set @resultado = '1'" & vbCrLf
        sql += "							end" & vbCrLf
        sql += "						end" & vbCrLf
        sql += "						else" & vbCrLf
        sql += "						begin" & vbCrLf
        sql += "							if @TipoPermiso='Borrar'" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "								if @cuenta>0" & vbCrLf
        sql += "									begin" & vbCrLf
        sql += "												select @resultado = isnull(Borrar,'0')" & vbCrLf
        sql += "												FROM [DB_PERMISOS_CARPETAS]" & vbCrLf
        sql += "												WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta)" & vbCrLf
        sql += "									end " & vbCrLf
        sql += "								else" & vbCrLf
        sql += "									begin" & vbCrLf
        sql += "												select @cuenta= Count(*)" & vbCrLf
        sql += "												FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "														  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "												WHERE (P.IdCuadro = @carpeta) and ( P.[Borrar] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "													if @cuenta = 0        	" & vbCrLf
        sql += "														set @resultado = '0'" & vbCrLf
        sql += "													Else" & vbCrLf
        sql += "														set @resultado = '1'" & vbCrLf
        sql += "									end" & vbCrLf
        sql += "							end" & vbCrLf
        sql += "							else" & vbCrLf
        sql += "							begin" & vbCrLf
        sql += "								if @TipoPermiso='Imprimir'" & vbCrLf
        sql += "								begin" & vbCrLf
        sql += "								if @cuenta>0" & vbCrLf
        sql += "										begin" & vbCrLf
        sql += "													select @resultado = isnull(Imprimir,'0')" & vbCrLf
        sql += "													FROM [DB_PERMISOS_CARPETAS]" & vbCrLf
        sql += "													WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta) " & vbCrLf
        sql += "										end " & vbCrLf
        sql += "									else" & vbCrLf
        sql += "										begin" & vbCrLf
        sql += "													select @cuenta= Count(*)" & vbCrLf
        sql += "													FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "															  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "													WHERE (P.IdCuadro = @carpeta) and ( P.[Imprimir] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "														if @cuenta = 0        	" & vbCrLf
        sql += "															set @resultado = '0'" & vbCrLf
        sql += "														Else" & vbCrLf
        sql += "															set @resultado = '1'" & vbCrLf
        sql += "										end" & vbCrLf
        sql += "								end" & vbCrLf
        sql += "								else" & vbCrLf
        sql += "								begin" & vbCrLf
        sql += "									if @TipoPermiso='Historico'" & vbCrLf
        sql += "									begin" & vbCrLf
        sql += "										if @cuenta>0" & vbCrLf
        sql += "											begin" & vbCrLf
        sql += "														select @resultado = isnull(Historico,'0')" & vbCrLf
        sql += "														FROM [DB_PERMISOS_CARPETAS]" & vbCrLf
        sql += "														WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta) " & vbCrLf
        sql += "											end " & vbCrLf
        sql += "										else" & vbCrLf
        sql += "											begin" & vbCrLf
        sql += "														select @cuenta= Count(*)" & vbCrLf
        sql += "														FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "																  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "														WHERE (P.IdCuadro = @carpeta) and ( P.[Historico] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "															if @cuenta = 0        	" & vbCrLf
        sql += "																set @resultado = '0'" & vbCrLf
        sql += "															Else" & vbCrLf
        sql += "																set @resultado = '1'" & vbCrLf
        sql += "											end" & vbCrLf
        sql += "									end" & vbCrLf
        sql += "									else" & vbCrLf
        sql += "									begin	" & vbCrLf
        sql += "											if @TipoPermiso='Seguridad'" & vbCrLf
        sql += "											begin" & vbCrLf
        sql += "												if @cuenta>0" & vbCrLf
        sql += "													begin" & vbCrLf
        sql += "																select @resultado = isnull(Seguridad,'0')" & vbCrLf
        sql += "																FROM [DB_PERMISOS_CARPETAS]" & vbCrLf
        sql += "																WHERE (CodUsuario =@usuario ) AND (IdCuadro =@carpeta) " & vbCrLf
        sql += "													end " & vbCrLf
        sql += "												else" & vbCrLf
        sql += "													begin" & vbCrLf
        sql += "																select @cuenta= Count(*)" & vbCrLf
        sql += "																FROM DB_PERMISOS AS P  INNER JOIN" & vbCrLf
        sql += "																		  DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "																WHERE (P.IdCuadro = @carpeta) and ( P.[Seguridad] = 1) AND (Grupo.CodUsuario = @usuario)" & vbCrLf
        sql += "																	if @cuenta = 0        	" & vbCrLf
        sql += "																		set @resultado = '0'" & vbCrLf
        sql += "																	Else" & vbCrLf
        sql += "																		set @resultado = '1'" & vbCrLf
        sql += "													end" & vbCrLf
        sql += "											end" & vbCrLf
        sql += "											else" & vbCrLf
        sql += "											begin" & vbCrLf
        sql += "													set @resultado='0'" & vbCrLf
        sql += "											end" & vbCrLf
        sql += "									end" & vbCrLf
        sql += "								end" & vbCrLf
        sql += "							end" & vbCrLf
        sql += "						end" & vbCrLf
        sql += "					end" & vbCrLf
        sql += "		end " & vbCrLf
        sql += "        return @resultado " & vbCrLf
        sql += "        End "

        crearfuncionoactualizarfuncion(conexion, " dbo.udf_Permiso_Carpeta", sql)
    End Sub
    Public Sub crearfuncionPermisoVer_Carpeta(ByVal conexion As ConexionSQL)
        Dim sql As String
        sql = "CREATE FUNCTION [dbo].[udf_PermisoVer_Carpeta] " & vbCrLf
        sql += "( " & vbCrLf
        sql += " @carpeta nvarchar(max), " & vbCrLf
        sql += " @usuario nvarchar(max) " & vbCrLf
        sql += " ) " & vbCrLf
        sql += "RETURNS nvarchar(max)" & vbCrLf
        sql += "AS " & vbCrLf
        sql += "BEGIN" & vbCrLf
        sql += "declare @ver nvarchar(1) " & vbCrLf
        sql += "declare @resultado nvarchar(1) " & vbCrLf

        sql += "select @ver = P.ver FROM [DB_PERMISOS] AS P INNER JOIN [DB_CUADRO] " & vbCrLf
        sql += "AS C ON P.idcuadro = C.id WHERE (P.CodUsuario =@usuario ) AND (C.id =@carpeta) " & vbCrLf
        sql += "GROUP BY P.ver HAVING (P.ver = N'1') " & vbCrLf

        sql += "if @ver is null " & vbCrLf

        sql += "begin" & vbCrLf

        sql += "	select @ver = P.ver " & vbCrLf
        sql += "	FROM DB_PERMISOS AS P INNER JOIN" & vbCrLf
        sql += "              DB_CUADRO AS C ON P.IdCuadro = C.id INNER JOIN" & vbCrLf
        sql += "              DocBoxMain.dbo.G_USRGRUPO AS Grupo ON P.IdGrupo = Grupo.IdGrupo" & vbCrLf
        sql += "	WHERE (C.id = @carpeta) and (P.Ver = N'1') AND (Grupo.CodUsuario = @usuario) AND" & vbCrLf
        sql += "                  ((SELECT     Ver" & vbCrLf
        sql += "                      FROM         DB_PERMISOS AS P1" & vbCrLf
        sql += "                      WHERE     (P.IdCuadro = IdCuadro) AND (CodUsuario = @usuario)" & vbCrLf
        sql += "                      GROUP BY Ver) IS NULL) OR" & vbCrLf
        sql += "           (C.id = @carpeta) AND (P.Ver = N'1') AND (Grupo.CodUsuario = @usuario) AND" & vbCrLf
        sql += "                 ((SELECT     Ver" & vbCrLf
        sql += "                      FROM         DB_PERMISOS AS P1" & vbCrLf
        sql += "                      WHERE     (P.IdCuadro = IdCuadro) AND (CodUsuario = @usuario)" & vbCrLf
        sql += "                      GROUP BY Ver) = '1')" & vbCrLf

        sql += "    if @ver is null " & vbCrLf
        sql += "	    set @resultado = '0'" & vbCrLf
        sql += "Else" & vbCrLf
        sql += "	    set @resultado = @ver	" & vbCrLf
        sql += "    End" & vbCrLf

        sql += "ELSE" & vbCrLf
        sql += "	set @resultado = @ver" & vbCrLf

        sql += "return @resultado " & vbCrLf
        sql += "    End" & vbCrLf
        crearfuncionoactualizarfuncion(conexion, "dbo.udf_PermisoVer_Carpeta", sql)
    End Sub
    Public Sub crearfucionbusqueda(ByVal conexion As ConexionSQL)
        Dim sql As String

        sql = "CREATE FUNCTION [dbo].[udf_select_concat_db_dato] ( @c nvarchar(255) ) RETURNS nVARCHAR(max) AS BEGIN DECLARE @p nVARCHAR(max) ; SET @p = '' ; SELECT @p = @p + Valor + ' ' FROM [db_datos] WHERE idcuadro = @c ; RETURN @p  END  "
        crearfuncionoactualizarfuncion(conexion, "dbo.udf_select_concat_db_dato", sql)

    End Sub

    Public Sub crearfucionbusqueda_objeto(ByVal conexion As ConexionSQL)
        Dim sql As String

        sql = "CREATE FUNCTION [dbo].[udf_select_concat] ( @c nvarchar(255), @n nvarchar(255)) RETURNS VARCHAR(max) AS BEGIN DECLARE @p VARCHAR(max) ; SET @p = '' ; SELECT @p = @p + Valor + ' ' FROM [db_datos_objeto] WHERE Objeto = @c and idobjeto = @n; RETURN @p +' ' + @c END "
        crearfuncionoactualizarfuncion(conexion, "dbo.udf_select_concat", sql)

    End Sub
    Public Sub crearfucionprebusqueda_objeto(ByVal conexion As ConexionSQL)
        Dim sql As String

        sql = "CREATE FUNCTION udf_DatosConcatenados  " & vbCrLf
        sql = sql + "( " & vbCrLf
        sql = sql + ") " & vbCrLf
        sql = sql + "RETURNS TABLE " & vbCrLf
        sql = sql + "AS " & vbCrLf
        sql = sql + "Return " & vbCrLf
        sql = sql + "( " & vbCrLf
        sql = sql + "select objeto , dbo.udf_select_concat(C.Objeto,c,idobjeto ) as datos from db_datos_objeto C  " & vbCrLf
        sql = sql + ") " & vbCrLf

        crearfuncionoactualizarfuncion(conexion, "dbo.udf_DatosConcatenados", sql)

    End Sub

    Public Sub crearfuncionoactualizarvista(ByVal conexion As ConexionSQL, ByVal nombre As String, ByVal cuerpo As String)
        Try

            If conexion.ExisteTablaSQL(nombre) Then
                cuerpo = cuerpo.Replace("CREATE VIEW ", "ALTER VIEW ")
            End If
            conexion.Ejecuta(cuerpo)
        Catch ex As Exception
            Throw New Exception("Error creando VIEW SQL." & nombre & "." & ex.Message & "." & cuerpo)
        End Try
    End Sub
    Public Sub crearfuncionoactualizarfuncion(ByVal conexion As ConexionSQL, ByVal nombre As String, ByVal cuerpo As String)
        Try

            If conexion.ExisteTablaSQL(nombre) Then
                cuerpo = cuerpo.Replace("CREATE FUNCTION", "ALTER FUNCTION")
            End If
            conexion.Ejecuta(cuerpo)
        Catch ex As Exception

        End Try
    End Sub
    Public Function creatablasbd(ByVal nombre As String, ByVal cadena As String, ByVal historico As Boolean) As Boolean
        Dim result As Boolean
        Try
            'tablas base de datos fulltext
            Dim conexion As ConexionSQL
            Dim s As String
            Dim idautonumerico as string = string.empty
            If Not historico Then idautonumerico = " IDENTITY(1,1) "


            '06/11/2013 
            'CAMBIOS EN LA TABLA DB_DATOS Y DB_DATOS_OBJETO
            'CAMPO VALOR ERAN NVARCHAR(MAX) POR NVARCHAR(255)
            'INDICES EN TABLAS:
            'DB_DATOS
            'DB_DATOS_OBJETO
            'DB_CUADRO



            conexion = New ConexionSQL(cadena)
            If conexion.ExisteTablaSQL("DB_CUADRO") = False Then
                'cuadro
                s = "CREATE TABLE DB_CUADRO (id numeric (18, 0) " & idautonumerico & " NOT NULL ," & _
                " Nombre  NVARCHAR(300) NOT NULL, IdPadre numeric (18, 0) NOT NULL,	Orden numeric (18, 0) not NULL," & _
                " IdSerie numeric (18, 0) NULL,	Carpeta nvarchar (1) NOT NULL, [Plantilla] numeric (18, 0) NULL ," & _
                " [IdNorma_dest] [numeric](18, 0) NULL,	[campofecha_dest] [nvarchar](150) NULL, [IdObjeto] [numeric](18, 0) NULL," & _
                " [Objeto] [nvarchar](100) NULL, NombrePadre  NVARCHAR(300) NULL, NombreOriginal nvarchar(300) null, " & _
                " CampoOrdenacion nvarchar(255) NULL,	[CampoOrdenacion_fecha] [datetime] NULL, MetaDatos nvarchar(max) null, MetaObjetos nvarchar(max) null, " & _
                " [Paginas] [numeric](18, 0) NULL PRIMARY KEY CLUSTERED (" & _
                " [ID] Asc ) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS " & _
                " = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
                conexion.Ejecuta(s)

                '12/02/2012 *** INICIO
                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO] ON [dbo].[DB_CUADRO] ([IdPadre] ASC ) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_idpadre_orden] ON [dbo].[DB_CUADRO] ( [IdPadre] ASC,[Orden] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

                '11/11/2013
                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_nombre] ON [dbo].[DB_CUADRO] ( [nombre] ASC ) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                '11/11/2013
                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_NombrePadre] ON [dbo].[DB_CUADRO] ( [NombrePadre] ASC ) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                '11/11/2013
                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_NombreOriginal] ON [dbo].[DB_CUADRO] ( [NombreOriginal] ASC ) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                '11/11/2013
                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_CampoOrdenacion] ON [dbo].[DB_CUADRO] ( [CampoOrdenacion] ASC ) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_CampoOrdenacionfecha] ON [dbo].[DB_CUADRO] ( [CampoOrdenacion_fecha] ASC ) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_idserie_carpeta] ON [dbo].[DB_CUADRO] ( [IdSerie] ASC,[Carpeta] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_idobjeto] ON [dbo].[DB_CUADRO] ( [IdObjeto] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "CREATE NONCLUSTERED INDEX [IX_DB_CUADRO_objeto] ON [dbo].[DB_CUADRO] ( [Objeto] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "sp_fulltext_database enable CREATE FULLTEXT CATALOG DB_DATOSMetadatosCat"
                conexion.Ejecuta(s)
                s = "CREATE UNIQUE INDEX ui_ukDB_DATOSMetadatosCat ON DB_CUADRO(ID)"
                conexion.Ejecuta(s)
                s = "CREATE FULLTEXT INDEX ON dbo.DB_CUADRO([MetaDatos] )" & _
                " KEY INDEX ui_ukDB_DATOSMetadatosCat ON DB_DATOSMetadatosCat WITH CHANGE_TRACKING AUTO"
                conexion.Ejecuta(s)

                s = "CREATE TABLE DB_BORRADOR_CUADRO (id numeric (18, 0) " & idautonumerico & " NOT NULL ," & _
                " Nombre  NVARCHAR(300) NOT NULL, IdPadre numeric (18, 0) NOT NULL,	Orden numeric (18, 0) not NULL," & _
                " IdSerie numeric (18, 0) NULL,	Carpeta nvarchar (1) NOT NULL, [Plantilla] numeric (18, 0) NULL ," & _
                " [IdNorma_dest] [numeric](18, 0) NULL,	[campofecha_dest] [nvarchar](150) NULL, [IdObjeto] [numeric](18, 0) NULL," & _
                " [Objeto] [nvarchar](100) NULL, NombrePadre  NVARCHAR(300) NULL, NombreOriginal nvarchar(300) null, " & _
                " CampoOrdenacion nvarchar(255) NULL, [CampoOrdenacion_fecha] [datetime] NULL, MetaDatos nvarchar(max) null, " & _
                " MetaObjetos nvarchar(max) null, [FechaBorrado] [datetime] NULL,	[Usuario] [nvarchar](50) NULL," & _
                " [IdCuadro] [numeric](18, 0) NOT NULL, [Paginas] [numeric](18, 0) NULL 	" & _
                " PRIMARY KEY CLUSTERED (" & _
                " [ID] Asc ) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS " & _
                " = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
                conexion.Ejecuta(s)

            End If

            'cabecera OBJETO
            If conexion.ExisteTablaSQL("DB_COBJETO") = False Then
                s = "CREATE TABLE DB_COBJETO(id numeric (18, 0) IDENTITY(1,1) NOT NULL PRIMARY KEY," & _
                " Nombre   NVARCHAR(150) NOT NULL,	[Referencia] [nvarchar](150) NULL,[Identificador] [nvarchar](150) NULL,);"
                conexion.Ejecuta(s)
            End If
            'linea OBJETO
            If conexion.ExisteTablaSQL("DB_LOBJETO") = False Then
                s = "CREATE TABLE DB_LOBJETO(id numeric (18, 0) IDENTITY(1,1) NOT NULL PRIMARY KEY," & _
                " IdObjeto numeric (18, 0) NOT NULL, Nombre   NVARCHAR(150) NOT NULL, TipoSql nvarchar (1) NOT NULL," & _
                " VForm nvarchar (1) NOT NULL, TipoCon nvarchar (2) NOT NULL, Tamano numeric (10, 0) NOT NULL," & _
                " Orden numeric (5, 0) NOT NULL, Obligat  nvarchar (2) NOT NULL, Valida nvarchar (400) NULL, Selector ntext NULL);"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_DATOS_OBJETO") = False Then
                s = "CREATE TABLE DB_DATOS_OBJETO(id numeric (18, 0) IDENTITY(1,1) NOT NULL PRIMARY KEY," & _
                " IdObjeto numeric (18, 0) NOT NULL, IdLObjeto numeric (18, 0) NOT NULL, " & _
                " Objeto nvarchar (100) NOT NULL, valor nvarchar(255) NULL);"
                '06/11/2013 *** MAX POR 255
                '" Objeto nvarchar (100) NOT NULL, valor nvarchar(max) NULL);"
                conexion.Ejecuta(s)
                '06/11/2013 *** indices
                s = "CREATE NONCLUSTERED INDEX [IX_DB_DATOS_OBJETO_idobjeto_objeto_valor] ON [dbo].[DB_DATOS_OBJETO] ( [idobjeto] ASC,[objeto] ASC,[valor] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_DATOS_COBJETO") = False Then
                s = "CREATE TABLE DB_DATOS_COBJETO(id numeric (18, 0) IDENTITY(1,1) NOT NULL PRIMARY KEY," & _
                " IdObjeto numeric (18, 0) NOT NULL, " & _
                " Objeto nvarchar (100) NOT NULL,  MetaObjetos nvarchar(max) null);"
                conexion.Ejecuta(s)
                s = "CREATE NONCLUSTERED INDEX [IX_DB_DATOS_COBJETO_idobjeto_objeto] ON [dbo].[DB_DATOS_COBJETO] ( [idobjeto] ASC,[objeto] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
            End If
            'cabecera serie
            If conexion.ExisteTablaSQL("DB_CSERIE") = False Then
                s = "CREATE TABLE DB_CSERIE(id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " Nombre   NVARCHAR(150) NOT NULL, Objeto numeric (18, 0)  NULL, Orden numeric (18, 0)  NULL,	[Titulo] [nvarchar](150) NULL, [OrdenBusquedaCampos] [nvarchar](150) NULL," & _
                "[CampoStamp] [nvarchar](50) NULL,	[ValorCampoStamp1] [nvarchar](50) NULL,	[FicheroStamp1] [nvarchar](200) NULL,	[ValorCampoStamp2] [nvarchar](50) NULL," & _
                "[FicheroStamp2] [nvarchar](200) NULL);"
                conexion.Ejecuta(s)
            End If
            'linea serie
            If conexion.ExisteTablaSQL("DB_LSERIE") = False Then
                s = "CREATE TABLE DB_LSERIE(id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdSerie numeric (18, 0) NOT NULL, Nombre   NVARCHAR(150) NOT NULL, TipoSql nvarchar (1) NOT NULL," & _
                " VForm nvarchar (1) NOT NULL, TipoCon nvarchar (2) NOT NULL, Tamano numeric (10, 0) NOT NULL," & _
                " Orden numeric (5, 0) NOT NULL, Obligat  nvarchar (2) NOT NULL, Valida nvarchar (400) NULL, Selector ntext NULL);"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_DATOS") = False Then
                s = "CREATE TABLE DB_DATOS(id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdCuadro numeric (18, 0) NOT NULL, IdLSerie numeric (18, 0) NOT NULL, " & _
                " valor nvarchar(255) NULL);"
                '06/11/2013 *** MAX POR 255
                '" valor nvarchar(max) NULL);"
                conexion.Ejecuta(s)
                s = "CREATE NONCLUSTERED INDEX [IX_DB_DATOS_valor_idcuadro_idlserie] ON [dbo].[DB_DATOS] ( [valor] ASC,[idlserie] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = ON, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "CREATE TABLE DB_BORRADOR_DATOS(id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdCuadro numeric (18, 0) NOT NULL, IdLSerie numeric (18, 0) NOT NULL, " & _
                " valor nvarchar(255) NULL);"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_FICHEROS") = False Then
                s = "CREATE TABLE DB_FICHEROS(id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdCuadro numeric (18, 0) NOT NULL, " & _
                " FileContent varbinary(max) NULL, NombreFichero nvarchar(100) NOT NULL, NombreOriginal nvarchar(300) NOT NULL,	[FileType] [nvarchar](20) NULL , [Firmado] [nvarchar](1) NULL);"
                conexion.Ejecuta(s)
                s = "sp_fulltext_database enable CREATE FULLTEXT CATALOG DB_FICHEROSFTCat"
                conexion.Ejecuta(s)
                s = "CREATE UNIQUE INDEX ui_ukDB_FICHEROSFTCat ON DB_FICHEROS(ID)"
                conexion.Ejecuta(s)
                s = "CREATE FULLTEXT INDEX ON dbo.DB_FICHEROS([FileContent] TYPE COLUMN FileType)" & _
                " KEY INDEX ui_ukDB_FICHEROSFTCat ON DB_FICHEROSFTCat WITH CHANGE_TRACKING AUTO"
                conexion.Ejecuta(s)

                s = "CREATE TABLE DB_BORRADOR_FICHEROS(id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdCuadro numeric (18, 0) NOT NULL, " & _
                " FileContent varbinary(max) NULL, NombreFichero nvarchar(100) NOT NULL, NombreOriginal nvarchar(300) NOT NULL,	[FileType] [nvarchar](20) NULL, [Firmado] [nvarchar](1) NULL);"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_PERMISOS") = False Then
                s = "CREATE TABLE DB_PERMISOS (id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdCuadro numeric (18, 0) NOT NULL, CodUsuario  NVARCHAR(20) NULL,	IdGrupo numeric (18, 0) NULL," & _
                " Ver [bit] NULL, [Añadir] [bit] NULL, Modificar [bit] NULL, Borrar [bit] NULL, Imprimir [bit] NULL," & _
                " Historico [bit] NULL,	Seguridad [bit] NULL);"
                conexion.Ejecuta(s)
                '12/02/2012 *** FIN

                s = "CREATE NONCLUSTERED INDEX [IX_DB_PERMISOS] ON [dbo].[DB_PERMISOS] ([CodUsuario] ASC,[IdCuadro] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                s = "CREATE NONCLUSTERED INDEX [ver] ON [dbo].[DB_PERMISOS]  ([Ver] asc)  WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                s = "CREATE NONCLUSTERED INDEX [idgrupo,cuadro] ON [dbo].[DB_PERMISOS] ([idgrupo] Asc,[IdCuadro] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                '12/02/2012 *** FIN

                s = "CREATE TABLE DB_BORRADOR_PERMISOS (id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdCuadro numeric (18, 0) NOT NULL, CodUsuario  NVARCHAR(20) NULL,	IdGrupo numeric (18, 0) NULL," & _
                " Ver [bit] NULL, [Añadir] [bit] NULL, Modificar [bit] NULL, Borrar [bit] NULL, Imprimir [bit] NULL," & _
                " Historico [bit] NULL,	Seguridad [bit] NULL);"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_PERMISOS_CARPETAS") = False Then
                s = "CREATE TABLE DB_PERMISOS_CARPETAS (id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdCuadro numeric (18, 0) NOT NULL, CodUsuario  NVARCHAR(20) NULL,	IdGrupo numeric (18, 0) NULL," & _
                " Ver [bit] NULL, [Añadir] [bit] NULL, Modificar [bit] NULL, Borrar [bit] NULL, Imprimir [bit] NULL," & _
                " Historico [bit] NULL,	Seguridad [bit] NULL);"
                conexion.Ejecuta(s)
                '12/02/2012 *** FIN

                s = "CREATE NONCLUSTERED INDEX [IX_DB_PERMISOS_CARPETAS] ON [dbo].[DB_PERMISOS_CARPETAS] ([CodUsuario] ASC,[IdCuadro] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                '12/02/2012 *** FIN
            End If
            If conexion.ExisteTablaSQL("DB_PERMISOS_OBJETO") = False Then
                s = "CREATE TABLE DB_PERMISOS_OBJETO (id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdObjeto numeric (18, 0) NOT NULL, Objeto nvarchar (100) NOT NULL,CodUsuario  NVARCHAR(20) NULL,	IdGrupo numeric (18, 0) NULL," & _
                " Ver [bit] NULL, [Añadir] [bit] NULL, Modificar [bit] NULL, Borrar [bit] NULL, Imprimir [bit] NULL," & _
                " Historico [bit] NULL,	Seguridad [bit] NULL);"
                conexion.Ejecuta(s)
                '12/02/2012 *** FIN

                s = "CREATE NONCLUSTERED INDEX [IX_DB_PERMISOS_OBJETOS] ON [dbo].[DB_PERMISOS_OBJETO] ([CodUsuario] ASC,[IdObjeto] ASC,[Objeto] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                '12/02/2012 *** FIN
            End If
            If conexion.ExisteTablaSQL("DB_PERMISOS_IDOBJETO") = False Then
                s = "CREATE TABLE DB_PERMISOS_IDOBJETO (id numeric (18, 0) " & idautonumerico & " NOT NULL PRIMARY KEY," & _
                " IdObjeto numeric (18, 0) NOT NULL, CodUsuario  NVARCHAR(20) NULL,	IdGrupo numeric (18, 0) NULL," & _
                " Ver [bit] NULL, [Añadir] [bit] NULL, Modificar [bit] NULL, Borrar [bit] NULL, Imprimir [bit] NULL," & _
                " Historico [bit] NULL,	Seguridad [bit] NULL);"
                conexion.Ejecuta(s)
                '12/02/2012 *** FIN

                s = "CREATE NONCLUSTERED INDEX [IX_DB_PERMISOS_IDOBJETOS] ON [dbo].[DB_PERMISOS_IDOBJETO] ([CodUsuario] ASC,[IdObjeto] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)
                '12/02/2012 *** FIN
            End If
            If conexion.ExisteTablaSQL("DB_PLANTILLA") = False Then
                s = " CREATE TABLE [dbo].[DB_PLANTILLA]([Id] [int] " & idautonumerico & " NOT NULL, " & _
                " [Nombre] [nvarchar](100) NOT NULL, [Fichero] [nvarchar](100) NOT NULL, " & _
                " [PosX] [decimal](18, 0) NOT NULL,	[PosY] [decimal](18, 0) NOT NULL, " & _
                " [EspaciadoH] [decimal](18, 2) NOT NULL, [EspaciadoV] [decimal](18, 2) NOT NULL, " & _
                " [TamFont] [int] NOT NULL,	[Negrita] [bit] NOT NULL) ON [PRIMARY]"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_BUSQUEDA") = False Then
                s = " CREATE TABLE [dbo].[DB_BUSQUEDA]([Id] [int] " & idautonumerico & " NOT NULL," & _
                " [Nombre] [nvarchar](100) NOT NULL, [Serie] [numeric](18, 0) NULL, " & _
                " [Objeto] [numeric](18, 0) NULL,[Identificador] [nvarchar](100) NULL, " & _
                " [ComboCarpeta] [bit] NULL, [SignoCarpeta] [nvarchar](50) NULL, " & _
                " [ValorCarpeta] [nvarchar](100) NULL, [CampoDato1] [nvarchar](150) NULL, " & _
                " [SignoDato1] [nvarchar](50) NULL,	[ValorDato1] [nvarchar](100) NULL, " & _
                " [AndOr1] [nvarchar](50) NULL,	[CampoDato2] [nvarchar](150) NULL, " & _
                " [SignoDato2] [nvarchar](50) NULL,	[ValorDato2] [nvarchar](100) NULL, " & _
                " [AndOr2] [nvarchar](50) NULL, [CampoDato3] [nvarchar](150) NULL, " & _
                " [SignoDato3] [nvarchar](50) NULL, [ValorDato3] [nvarchar](100) NULL, " & _
                " [Andor3] [nvarchar](50) NULL,	[CampoDato4] [nvarchar](150) NULL, " & _
                " [SignoDato4] [nvarchar](50) NULL,	[ValorDato4] [nvarchar](100) NULL, 	[Usuario] [nvarchar](50) NULL, " & _
                " [TextoCompleto] [bit] NULL, " & _
                " [CamposSerie] [bit] NULL) ON [PRIMARY]"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_CONSULTA") = False Then
                s = "CREATE TABLE [dbo].[DB_CONSULTA]([Id] [int] IDENTITY(1,1) NOT NULL,[Nombre] [nvarchar](100) NOT NULL,[Consulta] [nvarchar](max) NULL,[Usuario] [nvarchar](50) NULL) ON [PRIMARY]"
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_DATOS_REG") = False Then
                s = "CREATE TABLE [dbo].[DB_DATOS_REG](	[id] [numeric](18, 0) " & idautonumerico & " NOT NULL, " & _
                "[IdCuadro] [numeric](18, 0) NOT NULL, [IdLSerie] [numeric](18, 0) NOT NULL, " & _
                "[valor] [nvarchar](255) NULL, [FechaRegistro] [datetime] NULL, [UsuarioRegistro] " & _
                "[nvarchar](50) NULL, PRIMARY KEY CLUSTERED([ID] Asc )WITH (PAD_INDEX  = OFF, " & _
                "STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, " & _
                "ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "CREATE NONCLUSTERED INDEX [datosreg] ON [dbo].[DB_DATOS_REG] (	[IdCuadro] ASC,	[IdLSerie] ASC," & _
                "[FechaRegistro] ASC, [UsuarioRegistro] Asc )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = " & _
                "OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, " & _
                "ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                conexion.Ejecuta(s)

            End If
            If conexion.ExisteTablaSQL("DB_ACCION_REG") = False Then
                s = "CREATE TABLE [dbo].[DB_ACCION_REG]([id] [numeric](18, 0) " & idautonumerico & " NOT NULL, " & _
                "[IdCuadro] [numeric](18, 0) NOT NULL,	[Accion] [nvarchar](100) NULL,[FechaRegistro] [datetime] NULL, " & _
                "[UsuarioRegistro] [nvarchar](50) NULL, PRIMARY KEY CLUSTERED([ID] Asc )WITH (PAD_INDEX  = OFF, " & _
                "STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, " & _
                "ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
                conexion.Ejecuta(s)

                s = "CREATE NONCLUSTERED INDEX [accion] ON [dbo].[DB_ACCION_REG] " & _
                "(	[IdCuadro] ASC,	[FechaRegistro] ASC,  [UsuarioRegistro] Asc )WITH (PAD_INDEX  = OFF," & _
                "STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF," & _
                "ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                conexion.Ejecuta(s)
            End If
            If conexion.ExisteTablaSQL("DB_DATOS_OBJETO_REG") = False Then
                s = "CREATE TABLE [dbo].[DB_DATOS_OBJETO_REG](	[id] [numeric](18, 0) " & idautonumerico & " NOT NULL, " & _
                "[IdObjeto] [numeric](18, 0) NOT NULL, [IdLObjeto] [numeric](18, 0) NOT NULL,Objeto nvarchar (100) NOT NULL, " & _
                "[valor] [nvarchar](255) NULL, [FechaRegistro] [datetime] NULL, [UsuarioRegistro] " & _
                "[nvarchar](50) NULL, PRIMARY KEY CLUSTERED([ID] Asc )WITH (PAD_INDEX  = OFF, " & _
                "STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, " & _
                "ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
                conexion.Ejecuta(s)
                s = "CREATE NONCLUSTERED INDEX [IX_DB_DATOS_OBJETO_REG] ON [dbo].DB_DATOS_OBJETO_REG " & _
                "([IdObjeto] ASC,[Objeto] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, " & _
                " SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, " & _
                " ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
                conexion.Ejecuta(s)

            End If
            If conexion.ExisteTablaSQL("DB_ACCION_OBJETO_REG") = False Then
                s = "CREATE TABLE [dbo].[DB_ACCION_OBJETO_REG]([id] [numeric](18, 0) " & idautonumerico & " NOT NULL, " & _
                "[IdObjeto] [numeric](18, 0) NOT NULL,Objeto nvarchar (100) NOT NULL,	[Accion] [nvarchar](100) NULL,[FechaRegistro] [datetime] NULL, " & _
                "[UsuarioRegistro] [nvarchar](50) NULL, PRIMARY KEY CLUSTERED([ID] Asc )WITH (PAD_INDEX  = OFF, " & _
                "STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, " & _
                "ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
                conexion.Ejecuta(s)
                s = "CREATE NONCLUSTERED INDEX [IX_DB_ACCION_OBJETO_REG] ON [dbo].DB_ACCION_OBJETO_REG " & _
                "([IdObjeto] ASC,[Objeto] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF," & _
                " SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, " & _
                " ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                conexion.Ejecuta(s)

            End If

            If conexion.ExisteTablaSQL("DB_NORMA_DEST") = False Then
                s = "CREATE TABLE [dbo].[DB_NORMA_DEST]([Id] [numeric](18, 0) " & idautonumerico & " NOT NULL," & _
                "[nombre] [nvarchar](50) NOT NULL,	[meses] [int] NOT NULL, " & _
                " [tipo] [nvarchar](2) NOT NULL CONSTRAINT [DF_DB_NORMA_DEST_tipo]  DEFAULT ('D'), CONSTRAINT [PK_DB_NORMA_DEST] PRIMARY KEY CLUSTERED " & _
                "([ID] Asc) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS " & _
                "= ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
                conexion.Ejecuta(s)
                If Not historico Then
                    s = "insert into [dbo].[DB_NORMA_DEST]([nombre], [meses], [tipo]) values ('Indefinido',0,'H') "
                    conexion.Ejecuta(s)
                End If
            End If
            'CREO ESTA TABLA PORQUE SI NO DA ERROR AL ENTRAR EN EJECUCION DE NORMAS DE CONSERVACION 
            'AUNQUE CREO QUE ESE APARTADO FUE SUSTITUIDO POR LA APLICACION EXTERNA DE MIGUE
            '(SOLO HE ENCONTRADO LA TABLA EN LA BD DOCBOX_REGISTRO)
            '(NO SE CREA TAMPOCO EN LA APLICACION DE MIGUE DE NORMAS DE CONSERVACION)
            If conexion.ExisteTablaSQL("DB_EJEC_NORMA_DEST") = False Then
                s = "CREATE TABLE [dbo].[DB_EJEC_NORMA_DEST]([ID] [numeric](18, 0) IDENTITY(1,1) NOT NULL, " & _
                "[IDNorma] [numeric](18, 0) NOT NULL, [Meses] [int] NOT NULL, [Tipo] [varchar](2) NOT NULL, " & _
                "[CodUsuario] [nvarchar](20) NOT NULL, [Fecha] [datetime] NOT NULL,	[Procesado] [bit] NOT NULL, " & _
                "[FechaProcesado] [datetime] NULL, CONSTRAINT [PK_DB_EJEC_NORMA_DEST] PRIMARY KEY CLUSTERED ([ID] Asc " & _
                ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, " & _
                "ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] ) ON [PRIMARY]"
                conexion.Ejecuta(s)
            End If

            If conexion.ExisteTablaSQL("DB_LOGDESTRUCCION") = False Then
                s = "CREATE TABLE [dbo].[DB_LOGDESTRUCCION]( [ID] [numeric](18, 0) " & idautonumerico & " NOT NULL, " & _
                "[IDCUADRO] [numeric](18, 0) NOT NULL,	[Nombre] [nvarchar](300) NOT NULL, " & _
                " [Fecha] [datetime] NOT NULL,	[Usuario] [nvarchar](20) NOT NULL, " & _
                " CONSTRAINT [PK_DB_LogDestruccion] PRIMARY KEY CLUSTERED ([ID] Asc  " & _
                " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS " & _
                "= ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
                conexion.Ejecuta(s)
            End If
            'PARA PGR
            If conexion.ExisteTablaSQL("DB_DOCEXP") = False Then
                s = " CREATE TABLE [dbo].[DB_DOCEXP]([Id] [numeric](18, 0) IDENTITY(1,1) NOT NULL," & _
                " [Nombre] [nvarchar](150) NULL) ON [PRIMARY]"
                conexion.Ejecuta(s)
                rellenadatostablaDB_DOCEXP(conexion)
            End If


            'crear funciones
            crearfuncionPermisoVer_serie(conexion)
            crearfuncionPermisoVer_Carpeta(conexion)
            crearfuncionPermiso(conexion)
            crearfuncionPermiso_Carpeta(conexion)
            crearfuncionPermiso_objeto(conexion)
            crearfuncionPermiso_IDobjeto(conexion)
            crearfincionBuscarDato(conexion)
            crearfuncionPermiso_ver_nodorapido(conexion)
            crearfuncionf_generaarbol(conexion)
            'crearfuncionf_generaarbol_objeto_old(conexion)
            crearfuncionf_generaarbol_objeto(conexion)
            crearfuncionf_generaarbol_hijos(conexion)
            crearfuncionf_generaarbol_mas(conexion)
            crearfuncionf_generaarbol_menos(conexion)
            crearvista_cuenta_carpeta(conexion)
            crearfuncionf_generaarbol_carpetas(conexion)

            crearfuncionf_generaarbol_hijos_carpetas(conexion)
            crearfuncionf_generaarbol_mas_carpetas(conexion)
            crearfuncionf_generaarbol_menos_carpetas(conexion)
            crearfucionbusqueda(conexion)
            crearfucionbusqueda_objeto(conexion)
            crearvista_permisos(conexion)
            conexion.CerrarConexion()
            conexion = Nothing
            result = True

        Catch ex As Exception

            result = False

        Finally
            ' Por si se produce un error,
            ' comprobar si la conexión está abierta

        End Try
        Return result
    End Function

    Public Function permiso_usuario_sistema(ByVal usuario As String, ByVal basedatos As Long, ByVal conexion As String) As Boolean
        Dim sen As SentenciaSQL
        Dim con As ConexionSQL
        'pemiso para administrar 
        sen = New SentenciaSQL
        With sen
            .sql_select = CAMPO_SISTEMA
            .sql_from = TABLA_ROLUSRBD
            .add_condicion(CAMPO_CODUSUARIO, usuario)
            .add_condicion(CAMPO_IDBD, basedatos)
        End With
        con = New ConexionSQL(conexion)
        If con.ejecuta1v_long(sen.texto_sql) = Permiso_Administrador Then Return True Else Return False
    End Function

End Module
