Option Strict Off
Imports System.Data
Imports System.Web.UI.Page
Imports System.Web
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Diagnostics.Debug
Imports System.Text

Public Module GesCampusDocbox

    Public Sub SolicitudesAsignaRango(fechadesde As Date, fechahasta As Date, solidesde As Long, solihasta As Long, con As ConexionSQL, Real As Boolean)

        Dim dt As DataTable
        Dim dr As DataRow
        Dim cadena As String
        Dim nsoli As Long
        Dim montacadena As String



        Dim observaciones As String = VACIO

        observaciones = "ASIGNACION AUTOMATICA DE SOLICTIUDES"
        'TODO LOGININS
        If Real = True Then
            montacadena = " SELECT SOLICITUDES.NSolicitud, SOLICITUDES.CodTurno " & _
          " FROM SOLICITUDES LEFT OUTER JOIN SOLICITUDES_TURNOS ON SOLICITUDES.NSolicitud = SOLICITUDES_TURNOS.NSolicitud LEFT OUTER JOIN" & _
          " TURNOS ON SOLICITUDES_TURNOS.CodColonia = TURNOS.CodColonia AND SOLICITUDES_TURNOS.CodTurno = TURNOS.CodTurno LEFT OUTER JOIN " & _
          " COLONIAS ON SOLICITUDES_TURNOS.CodColonia = COLONIAS.CodColonia where (SOLICITUDES_TURNOS.Prioridad = 1)"


            cadena = montacadena & " AND  (SOLICITUDES.nsolicitud >= " & solidesde & " and SOLICITUDES.nsolicitud <= " & solihasta & ") " & _
                    " and (FechaEnvio >= " & Util.fechasql(fechadesde) & " and FechaEnvio <= " & Util.fechasql(fechahasta) & ") " & _
                    " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                    " and PteConsulta = 0 and FechaBaja is null order by FechaEnvio, SOLICITUDES.NSolicitud"



        Else
            montacadena = " SELECT SOLICITUDESSimulacion.NSolicitud, SOLICITUDESSimulacion.CodTurno " & _
          " FROM SOLICITUDESSimulacion LEFT OUTER JOIN SOLICITUDES_TURNOSSimulacion ON SOLICITUDESSimulacion.NSolicitud = SOLICITUDES_TURNOSSimulacion.NSolicitud LEFT OUTER JOIN" & _
          " TURNOSSimulacion ON SOLICITUDES_TURNOSSimulacion.CodColonia = TURNOSSimulacion.CodColonia AND SOLICITUDES_TURNOSSimulacion.CodTurno = TURNOSSimulacion.CodTurno LEFT OUTER JOIN " & _
          " COLONIAS ON SOLICITUDES_TURNOSSimulacion.CodColonia = COLONIAS.CodColonia where (SOLICITUDES_TURNOSSimulacion.Prioridad = 1)"


            cadena = montacadena & " AND  (SOLICITUDESSimulacion.nsolicitud >= " & solidesde & " and SOLICITUDESSimulacion.nsolicitud <= " & solihasta & ") " & _
                    " and (FechaEnvio >= " & Util.fechasql(fechadesde) & " and FechaEnvio <= " & Util.fechasql(fechahasta) & ") " & _
                    " and (SOLICITUDESSimulacion.CodTurno is null or SOLICITUDESSimulacion.CodTurno =0) " & _
                    " and PteConsulta = 0 and FechaBaja is null order by FechaEnvio, NSolicitud"
        End If
        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            nsoli = CStr(dr("NSolicitud"))
            AsignaSolicitud(nsoli, con, Real)
        Next

    End Sub

    Public Sub SolicitudesDesasigna(solidesde As Long, solihasta As Long, con As ConexionSQL, Real As Boolean)

        Dim dt As DataTable
        Dim dr As DataRow
        Dim cadena As String
        Dim nsoli As Long
        Dim sexo As String
        Dim codcolonia As Long
        Dim codturno As Long
        Dim regimen As String = VACIO
        Dim coddormitorio As String
        Dim listaespera As Integer
        Dim observaciones As String = VACIO
        If Real = True Then
            cadena = "SELECT  nsolicitud,codcolonia,coddormitorio,listaespera,regimen,codturno,sexo FROM SOLICITUDES where (nsolicitud >= " & _
                    solidesde & " and nsolicitud <= " & solihasta & ") " & _
                    " and (SOLICITUDES.CodColonia is not null and SOLICITUDES.CodColonia <>0) and  SOLICITUDES.estado = 'A' " & _
                    " and ((SOLICITUDES.CodTurno is not null and SOLICITUDES.CodTurno <>0) or( listaespera is not null and ListaEspera <> 0))  " & _
                    " and FechaRecepcion is null order by FechaEnvio, NSolicitud"
        Else
            cadena = "SELECT  nsolicitud,codcolonia,coddormitorio,listaespera,regimen,codturno,sexo FROM SOLICITUDESSimulacion as Solicitudes where (nsolicitud >= " & _
                  solidesde & " and nsolicitud <= " & solihasta & ") " & _
                  " and (SOLICITUDES.CodColonia is not null and SOLICITUDES.CodColonia <>0) and  SOLICITUDES.estado = 'A' " & _
                  " and ((SOLICITUDES.CodTurno is not null and SOLICITUDES.CodTurno <>0) or( listaespera is not null and ListaEspera <> 0)) " & _
                  " and FechaRecepcion is null order by FechaEnvio, NSolicitud"
        End If
        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            nsoli = CStr(dr("NSolicitud"))
            sexo = CStr(dr("Sexo"))
            codcolonia = (dr("codcolonia"))
            If dr("codturno").ToString = "" Then codturno = 0 Else codturno = (dr("codturno"))
            If dr("coddormitorio").ToString = "" Then coddormitorio = VACIO Else coddormitorio = (dr("coddormitorio").ToString)
            If dr("regimen").ToString = "" Then regimen = VACIO Else regimen = (dr("regimen").ToString)
            If dr("listaespera").ToString = "" Then listaespera = 0 Else listaespera = (dr("listaespera"))
            DesasignaSolicitud(nsoli, sexo, codcolonia, codturno, regimen, coddormitorio, listaespera, con, Real)
        Next

    End Sub

    Public Sub AsignaSolicitud(solicitud As Long, con As ConexionSQL, Real As Boolean)

        Dim dt As DataTable
        Dim dr As DataRow
        Dim cadena As String
        Dim sexo As String
        Dim vincu As String
        Dim regimen As String = VACIO
        Dim observaciones As String = VACIO

        If Real = True Then
            cadena = "SELECT CodVinculo,Sexo FROM SOLICITUDES where nsolicitud = " & solicitud
        Else
            cadena = "SELECT CodVinculo,Sexo FROM SOLICITUDESSimulacion where nsolicitud = " & solicitud
        End If
        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            sexo = CStr(dr("Sexo"))
            vincu = CStr(dr("CodVinculo").ToString)

            If vincu = VACIO Then
                AsignaSolicitudSola(solicitud, sexo, con, Real)
            Else
                If Real = True Then
                    cadena = "select count(*) from solicitudes where codvinculo ='" & vincu & "' and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                        "and PteConsulta = 0  and FechaBaja is null"
                Else
                    cadena = "select count(*) from solicitudesSimulacion as solicitudes where codvinculo ='" & vincu & "' and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                        "and PteConsulta = 0  and FechaBaja is null"
                End If
                If con.ejecuta1v_long(cadena) = 1 Then 'porque si no entra en bucle 
                    AsignaSolicitudSola(solicitud, sexo, con, Real)
                Else
                    AsignaSolicitudGrupo(solicitud, vincu, con, Real)

                End If
            End If
        Next

    End Sub
    Public Function HayPlazas(solicitud As Long, sexov As Integer, sexom As Integer, turno As Long, regimen As String, colonia As Long, con As ConexionSQL, Real As Boolean) As Integer
        Dim dt As DataTable
        Dim dr As DataRow
        Dim cadena As String
        Dim valor As Integer
        Dim libresmasc As Integer
        Dim libresfem As Integer
        Dim libresmascex As Integer
        Dim libresfemex As Integer
        If Real = True Then
            cadena = "select * from TurnoS where CodColonia = " & colonia & "and CodTurno = " & turno
        Else
            cadena = "select * from TurnoSSimulacion where CodColonia = " & colonia & "and CodTurno = " & turno
        End If
        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            libresmasc = CLng(dr("LibresMasc"))
            libresfem = CLng(dr("Libresfem"))
            libresmascex = CLng(dr("LibresMascex"))
            libresfemex = CLng(dr("Libresfemex"))
            Select Case regimen
                Case "I"
                    If libresmasc >= sexov And libresfem >= sexom Then valor = 1 Else valor = 0
                Case "E"
                    If libresmascex >= sexov And libresfemex >= sexom Then valor = 1 Else valor = 0
            End Select
        Next
        Return valor
    End Function

    Public Function AumentaPlaza(codcolonia As Long, codturno As Long, regimen As String, sexov As Integer, sexom As Integer, con As ConexionSQL, Real As Boolean) As Integer
        Dim dt As DataTable
        Dim dr As DataRow
        Dim cadena As String
        Dim valor As Boolean
        Dim cerrado As Boolean
        Dim libresmasc As Integer
        Dim libresfem As Integer
        Dim libresmascex As Integer
        Dim libresfemex As Integer
        Dim difemasc As Integer = 0
        Dim difefem As Integer = 0
        Dim difemascex As Integer = 0
        Dim difefemex As Integer = 0
        If Real = True Then
            cadena = "select * from Turnos where CodColonia = " & codcolonia & "and CodTurno = " & codturno
        Else
            cadena = "select * from TurnosSimulacion where CodColonia = " & codcolonia & "and CodTurno = " & codturno
        End If
        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            cerrado = CStr(dr("cerrado"))
            libresmasc = CLng(dr("LibresMasc"))
            libresfem = CLng(dr("Libresfem"))
            libresmascex = CLng(dr("LibresMascex"))
            libresfemex = CLng(dr("Libresfemex"))

            If cerrado Then
                valor = 0
            Else
                If regimen = "I" Then
                    If libresmasc < sexov Then difemasc = sexov - libresmasc Else difemasc = 0
                Else
                    If libresmascex < sexov Then difemascex = sexov - libresmascex Else difemascex = 0
                End If
                If regimen = "I" Then
                    If libresfem < sexom Then difefem = sexom - libresfem Else difefem = 0
                Else
                    If libresfemex < sexom Then difefemex = sexom - libresfemex Else difefemex = 0
                End If
                If Real = True Then
                    cadena = "update turnos set   numplazasmasc = numplazasmasc + " & difemasc & "," & _
                                                "numplazasfem = numplazasfem + " & difefem & "," & _
                                                "libresmasc = libresmasc + " & difemasc & "," & _
                                                "libresfem = libresfem + " & difefem & "," & _
                                                "numplazasmascex = numplazasmascex + " & difemascex & "," & _
                                                "numplazasfemex = numplazasfemex + " & difefemex & "," & _
                                                "libresmascex = libresmascex + " & difemascex & "," & _
                                                "libresfemex = libresfemex + " & difefemex & _
                                                " where codcolonia = " & codcolonia & " and codturno = " & codturno
                Else
                    cadena = "update turnosSimulacion set   numplazasmasc = numplazasmasc + " & difemasc & "," & _
                                               "numplazasfem = numplazasfem + " & difefem & "," & _
                                               "libresmasc = libresmasc + " & difemasc & "," & _
                                               "libresfem = libresfem + " & difefem & "," & _
                                               "numplazasmascex = numplazasmascex + " & difemascex & "," & _
                                               "numplazasfemex = numplazasfemex + " & difefemex & "," & _
                                               "libresmascex = libresmascex + " & difemascex & "," & _
                                               "libresfemex = libresfemex + " & difefemex & _
                                               " where codcolonia = " & codcolonia & " and codturno = " & codturno
                End If
                con.Ejecuta(cadena)
                valor = 1
            End If

        Next
        Return valor
    End Function
    Public Sub AsignaSolicitudGrupo(solicitud As Long, vinculo As String, con As ConexionSQL, Real As Boolean)
        Dim CodColonia As Long
        Dim CodTurno As Long
        Dim Regimen As String = VACIO
        Dim Forzado As Integer
        Dim resultado As Integer
        Dim observaciones As String
        Dim menor As Integer
        Dim sexov As Integer
        Dim sexom As Integer

        menor = MenorPreferente(solicitud, con, Real)

        If menor = 0 Then
            observaciones = "GR " & vinculo
            'TODO LOGININS
            Exit Sub
        End If

        CuentaGrupo(vinculo, sexov, sexom, con, Real)

        resultado = BuscaTurno(solicitud, sexov, sexom, CodColonia, CodTurno, Regimen, Forzado, con, Real)

        If resultado = 1 Then
            AsignaPlazaGrupo(vinculo, CodColonia, CodTurno, Regimen, Forzado, con, Real)
        Else
            PendienteGrupo(vinculo, CodColonia, CodTurno, Regimen, con, Real)
        End If


    End Sub


    Public Sub AsignaPlazaGrupo(vinculo As String, codcolonia As Long, codturno As Long, regimen As String, forzado As Integer, con As ConexionSQL, Real As Boolean)
        Dim solicitud As Long
        Dim observaciones As String
        Dim sexo As String = VACIO
        Dim cadena As String
        Dim sexov As Integer
        Dim sexom As Integer
        Dim dt As DataTable
        Dim dr As DataRow

        observaciones = "GR " & vinculo & "C" & codcolonia.ToString.Trim & "T" & codturno.ToString.Trim & "R" & regimen

        If forzado = 0 Then
            observaciones = observaciones & " Preferente"
        Else
            observaciones = observaciones & " Forzada"
        End If
        If Real = True Then
            cadena = "SELECT  nsolicitud,SEXO FROM SOLICITUDES where codvinculo = " & v(vinculo) & _
            " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
            " and PteConsulta = 0 and FechaBaja is null "
        Else
            cadena = "SELECT  nsolicitud,SEXO FROM SOLICITUDESSimulacion as solicitudes where codvinculo = " & v(vinculo) & _
                   " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                   " and PteConsulta = 0 and FechaBaja is null "
        End If
        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            sexo = CStr(dr("sexo"))
            solicitud = CLng(dr("nsolicitud"))
            If sexo = "V" Then sexov = 1 : sexom = 0 Else sexov = 0 : sexom = 1
            'TODO LOGININS
            AsignaPlaza(solicitud, sexov, sexom, codcolonia, codturno, regimen, forzado, "G", con, Real)
        Next

    End Sub
    Public Sub PendienteGrupo(vinculo As String, codcolonia As Long, codturno As Long, regimen As String, con As ConexionSQL, Real As Boolean)
        Dim solicitud As Long
        Dim observaciones As String
        Dim cadena As String
        Dim dt As DataTable
        Dim dr As DataRow

        observaciones = "GR " & vinculo & "C" & codcolonia.ToString.Trim & "T" & codturno.ToString.Trim & "R" & regimen & "Lista Espera"
        If Real = True Then



            cadena = "SELECT  nsolicitud FROM SOLICITUDES where codvinculo =" & v(vinculo) & _
            " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
            " and PteConsulta = 0 and FechaBaja is null "
        Else
            cadena = "SELECT  nsolicitud FROM SOLICITUDESSimulacion as solicitudes where codvinculo =" & v(vinculo) & _
           " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
           " and PteConsulta = 0 and FechaBaja is null "
        End If
        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            'TODO LOGININS
            solicitud = dr("nsolicitud").ToString
            RechazoPlaza(solicitud, codcolonia, codturno, regimen, con, Real)
        Next

    End Sub

    Public Sub CuentaGrupo(vinculo As String, ByRef sexov As Integer, ByRef sexom As Integer, con As ConexionSQL, Real As Boolean)
        If Real = True Then
            sexov = con.ejecuta1v_long("select count(*) from solicitudes where codvinculo =" & v(vinculo) & _
                         " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                         " and PteConsulta = 0  and FechaBaja is null and sexo='V'")

            sexom = con.ejecuta1v_long("select count(*) from solicitudes where codvinculo =" & v(vinculo) & _
                   " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                   " and PteConsulta = 0  and FechaBaja is null and sexo <>'V'")
        Else
            sexov = con.ejecuta1v_long("select count(*) from solicitudesSimulacion as solicitudes where codvinculo =" & v(vinculo) & _
                    " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                    " and PteConsulta = 0  and FechaBaja is null and sexo='V'")

            sexom = con.ejecuta1v_long("select count(*) from solicitudesSimulacion as solicitudes where codvinculo =" & v(vinculo) & _
                   " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                   " and PteConsulta = 0  and FechaBaja is null and sexo <>'V'")
        End If

    End Sub
    Public Function MenorPreferente(solicitud As Long, con As ConexionSQL, Real As Boolean) As Integer

        Dim cadena As String
        Dim fechaenvio As Date
        Dim posteriores As Integer
        Dim vinculo As String = VACIO
        Dim dt As DataTable
        Dim dr As DataRow
        If Real = True Then
            cadena = "SELECT fechaenvio,codvinculo FROM SOLICITUDES where nsolicitud = " & solicitud
        Else
            cadena = "SELECT fechaenvio,codvinculo FROM SOLICITUDESSimulacion where nsolicitud = " & solicitud
        End If
        dt = con.SelectSQL(cadena)
        con.Ejecuta(cadena)
        For Each dr In dt.Rows
            fechaenvio = Format(dr("FechaEnvio"), "dd/MM/yyyy HH:mm:ss")
            vinculo = CStr(dr("CodVinculo"))
        Next
        If Real = True Then
            posteriores = con.ejecuta1v_long("select count(*) from solicitudes where codvinculo =" & v(vinculo) & _
                         " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                         " and PteConsulta = 0  and FechaBaja is null and fechaenvio > " & Util.fechasql(fechaenvio, tipofecha.yyyyddMMmmss))
        Else
            posteriores = con.ejecuta1v_long("select count(*) from solicitudesSimulacion as solicitudes where codvinculo =" & v(vinculo) & _
                       " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) " & _
                       " and PteConsulta = 0  and FechaBaja is null and fechaenvio > " & Util.fechasql(fechaenvio, tipofecha.yyyyddMMmmss))
        End If
        If posteriores = 0 Then Return 1 Else Return 0

    End Function

    Public Sub SolicitudesAsignaTuno(solicitud As Long, codcolonia As Long, codturno As Long, regimen As String, con As ConexionSQL, Real As Boolean, CodPrecio As Integer)
        Dim ValorReq As String
        Dim resultado As Integer
        Dim observaciones As String
        Dim sexov As Integer
        Dim sexom As Integer
        Dim sexo As String
        If Real = True Then
            sexo = con.ejecuta1v_string("select sexo from solicitudes where nsolicitud =" & solicitud)
        Else
            sexo = con.ejecuta1v_string("select sexo from solicitudesSimulacion where nsolicitud =" & solicitud)
        End If

        If sexo = "V" Then
            sexov = 1 : sexom = 0
        Else
            sexov = 0 : sexom = 1
        End If

        resultado = HayPlazas(solicitud, sexov, sexom, codturno, regimen, codcolonia, con, Real)

        If resultado = 0 Then
            resultado = AumentaPlaza(codcolonia, codturno, regimen, sexov, sexom, con, Real)
            If resultado = 0 Then
                Exit Sub
            End If
        End If

        observaciones = "ASIGNACION MANUAL DE TURNO"
        'TODO LOGININS
        If Real = True Then
            ValorReq = con.ejecuta1v_string("select valorrequerido FROM SOLICITUDES_TURNOS WHERE codprecio=" & CodPrecio & " and  NSOLICITUD =" & solicitud)
            If ValorReq = "" Then
                If con.ejecuta1v_string("select  REQUIEREINFORMACION  from precios where codprecio = " & CodPrecio) = "True" Then ValorReq = "."
            End If
            con.Ejecuta("DELETE FROM SOLICITUDES_TURNOS WHERE NSOLICITUD =" & solicitud)

            con.Ejecuta("INSERT INTO SOLICITUDES_TURNOS (NSOLICITUD, CODCOLONIA,CODTURNO,REGIMEN,PRIORIDAD,codprecio,valorrequerido) VALUES (" & solicitud & "," & codcolonia & "," & codturno & "," & v(regimen) & ",1," & CodPrecio & ",'" & ValorReq & "')")
        Else
            con.Ejecuta("DELETE FROM SOLICITUDES_TURNOSSimulacion WHERE NSOLICITUD =" & solicitud)

            con.Ejecuta("INSERT INTO SOLICITUDES_TURNOSSimulacion (NSOLICITUD, CODCOLONIA,CODTURNO,REGIMEN,PRIORIDAD,CODPRECIO) VALUES (" & solicitud & "," & codcolonia & "," & codturno & "," & v(regimen) & ",1," & CodPrecio & ")")
        End If

        AsignaPlaza(solicitud, sexov, sexom, codcolonia, codturno, regimen, 0, "D", con, Real)
        If Real = True Then
            con.Ejecuta("UPDATE SOLICITUDES SET CodVinculo = '' WHERE NSOLICITUD =" & solicitud)
        Else
            con.Ejecuta("UPDATE SOLICITUDESSimulacion SET CodVinculo = '' WHERE NSOLICITUD =" & solicitud)
        End If



    End Sub

    Public Sub AsignaSolicitudSola(solicitud As Long, sexo As String, con As ConexionSQL, Real As Boolean)
        Dim CodColonia As Long
        Dim CodTurno As Long
        Dim Regimen As String = VACIO
        Dim Forzado As Integer
        Dim resultado As Integer
        Dim observaciones As String
        Dim sexov As Integer
        Dim sexom As Integer

        If sexo = "V" Then
            sexov = 1 : sexom = 0
        Else
            sexov = 0 : sexom = 1
        End If

        resultado = BuscaTurno(solicitud, sexov, sexom, CodColonia, CodTurno, Regimen, Forzado, con, Real)

        If resultado = 1 Then
            'HAY PLAZA
            If Forzado = 0 Then
                observaciones = "C" & CodColonia.ToString.Trim & "T" & CodTurno.ToString.Trim & "R" & Regimen & " Preferente"
            Else
                observaciones = "C" & CodColonia.ToString.Trim & "T" & CodTurno.ToString.Trim & "R" & Regimen & " Forzada"
            End If

            'TODO LOGININS
            AsignaPlaza(solicitud, sexov, sexom, CodColonia, CodTurno, Regimen, Forzado, "S", con, Real)
        Else
            'NO HAY PLAZAS
            observaciones = "C" & CodColonia.ToString.Trim & "T" & CodTurno.ToString.Trim & "R" & Regimen & " Lista Espera"

            'TODO LOGININS
            RechazoPlaza(solicitud, CodColonia, CodTurno, Regimen, con, Real)
        End If

    End Sub

    Public Sub AsignaPlaza(solicitud As Long, sexov As Integer, sexom As Integer, codcolonia As Long, codturno As Long, regimen As String, forzado As Integer, tipoAsignacion As String, con As ConexionSQL, Real As Boolean)
        Dim CodColoniaAnt As Long
        Dim CodTurnoAnt As Long
        Dim RegimenAnt As String = VACIO
        Dim ListaEsperaAnt As Long = 0
        Dim CodDormitorioAnt As String = VACIO
        Dim sexo As String = VACIO
        Dim cadena As String
        Dim fecha As Date

        Dim dt As DataTable
        Dim dr As DataRow

        If Real = True Then
            cadena = "select * from solicitudes where nsolicitud =" & solicitud
        Else
            cadena = "select * from solicitudesSimulacion where nsolicitud =" & solicitud
        End If

        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            If dr("CodColonia").ToString <> "" Then
                CodColoniaAnt = CStr(dr("CodColonia").ToString)
                CodTurnoAnt = CStr(dr("CodTurno").ToString)
                RegimenAnt = CStr(dr("Regimen").ToString)
                CodDormitorioAnt = CStr(dr("CodDormitorio").ToString)
                ListaEsperaAnt = CStr(dr("ListaEspera").ToString)
            End If
            sexo = CStr(dr("Sexo").ToString)
        Next

        If CodColoniaAnt <> 0 Then
            DesasignaSolicitud(solicitud, sexo, CodColoniaAnt, CodTurnoAnt, RegimenAnt, CodDormitorioAnt, ListaEsperaAnt, con, Real)
        End If

        fecha = Today

        EstadoSolicitud(solicitud, con, Real, codcolonia, codturno, regimen, forzado, fecha, tipoAsignacion, , , , "A")

        'TODO SI HAY QUE HACER UNA PREASIGNACION
        If Real = True Then
            If regimen = "I" Then
                cadena = "UPDATE TURNOS SET LibresMasc = LibresMasc - " & sexov & ", LibresFem = LibresFem - " & sexom
            Else
                cadena = "UPDATE TURNOS SET LibresMascex = LibresMascex- " & sexov & ", LibresFemex = LibresFemex - " & sexom
            End If
        Else
            If regimen = "I" Then
                cadena = "UPDATE TURNOSSimulacion SET LibresMasc = LibresMasc - " & sexov & ", LibresFem = LibresFem - " & sexom
            Else
                cadena = "UPDATE TURNOSSimulacion SET LibresMascex = LibresMascex- " & sexov & ", LibresFemex = LibresFemex - " & sexom
            End If
        End If

        cadena = cadena & " where codcolonia = " & codcolonia & " and codturno = " & codturno

        con.Ejecuta(cadena)

    End Sub
    Public Sub ActualizaFechaCarta(fechadesde As Date, fechahasta As Date, solidesde As Long, solihasta As Long, con As ConexionSQL)

        Dim dt As DataTable
        Dim dr As DataRow
        Dim cadena As String
        Dim nsoli As Long


        'TODO LOGININS

        cadena = "SELECT  nsolicitud,codturno FROM SOLICITUDES where (nsolicitud >= " & solidesde & " and nsolicitud <= " & solihasta & ") " & _
                " and (FechaEnvio >= " & Util.fechasql(fechadesde) & " and FechaEnvio <= " & Util.fechasql(fechahasta) & ") " & _
                " and (SOLICITUDES.CodTurno is not null or SOLICITUDES.CodTurno >0) " & _
                " and PteConsulta = 0 and FechaBaja is null and Fechacarta is null"

        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            nsoli = CStr(dr("NSolicitud"))
            con.Ejecuta("UPDATE SOLICITUDES SET             FechaCarta =" & Util.fechasql(Today) & " WHERE NSolicitud=" & nsoli & "")
            '  envía email admitido
        Next


        cadena = "SELECT  nsolicitud,codturno FROM SOLICITUDES where (nsolicitud >= " & solidesde & " and nsolicitud <= " & solihasta & ") " & _
                " and (FechaEnvio >= " & Util.fechasql(fechadesde) & " and FechaEnvio <= " & Util.fechasql(fechahasta) & ") " & _
                " and (SOLICITUDES.CodTurno is null or SOLICITUDES.CodTurno =0) and solicitudes.listaespera >1 " & _
                " and PteConsulta = 0 and FechaBaja is null and Fechacarta is null"

        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows
            nsoli = CStr(dr("NSolicitud"))
            con.Ejecuta("UPDATE SOLICITUDES SET             FechaCarta =" & Util.fechasql(Today) & " WHERE NSolicitud=" & nsoli & "")
            '  envía email No admitido
        Next

    End Sub

    Public Sub EstadoSolicitud(solicitud As Long, con As ConexionSQL, Real As Boolean, Optional codcolonia As Long = 0, Optional codturno As Long = 0, Optional regimen As String = "", Optional forzado As Integer = 0, Optional fechaconfirma As String = VACIO, Optional tipoAsignacion As String = "", Optional PteConsulta As Integer = 0, Optional ObsPteConsulta As String = "", Optional ListaEspera As Integer = 0, Optional EstadoSoli As String = "P")

        Dim cadena As String
        Dim fechaasignacion As Date
        Dim CodPrecio As Long

        If tipoAsignacion <> "" Then
            fechaasignacion = Today
        End If

        'NO ENTIENDO NADA

        'TODO SI HAY QUE HACER UNA PREASIGNACION
        If Real = True Then
            cadena = "UPDATE SOLICITUDES SET Estado='" & EstadoSoli & "',PteConsulta = " & PteConsulta & ",  " & _
                                            "CodColonia = " & codcolonia & ", ListaEspera = " & ListaEspera & ", " & _
                                            "CodTurno = " & codturno & ", " & "Forzado = " & forzado & ", " & "TipoAsignacion = " & v(tipoAsignacion)


        Else
            cadena = "UPDATE SOLICITUDESSimulacion SET Estado='" & EstadoSoli & "',PteConsulta = " & PteConsulta & ", " & _
                                            "CodColonia = " & codcolonia & ", ListaEspera = " & ListaEspera & ", " & _
                                            "CodTurno = " & codturno & ", " & "Forzado = " & forzado & ", " & "TipoAsignacion = " & v(tipoAsignacion)
        End If

        If regimen <> "" Then cadena = cadena & ", Regimen = " & v(regimen)

        If IsDate(fechaconfirma) And tipoAsignacion <> "N" Then cadena = cadena & ", FechaConfirma = " & Util.fechasql(fechaconfirma)



        If tipoAsignacion <> "" And tipoAsignacion <> "N" Then
            fechaasignacion = Today
            cadena = cadena & ", fechaasignacion = " & Util.fechasql(fechaasignacion)
        End If


        If tipoAsignacion <> "S" And tipoAsignacion <> "D" And tipoAsignacion <> "G" Then cadena = cadena & ", FechaCarta  = null "

        'tipo N desasignacion actualizao codprecio a null
        If tipoAsignacion = "N" Then cadena = cadena & ", fechaasignacion  = null , FechaConfirma = null, codprecio = 0"

        cadena = cadena & " where nsolicitud = " & solicitud

        con.Ejecuta(cadena)


        If EstadoSoli = "A" Then
            If Real = True Then
                CodPrecio = con.ejecuta1v_long("select codprecio from Solicitudes_Turnos where nsolicitud =" & solicitud & " and CodTurno = " & codturno & "")
                con.Ejecuta("Update SOLICITUDES set codprecio=" & CodPrecio & " where nsolicitud=" & solicitud & "")


                'If Real Then
                '    ' BUSCAR CODTURNO PRIORIDAD 1

                '    Dim priori As Integer
                '    priori = con.ejecuta1v_long("select prioridad from solicitudes_turnos where codturno =  " & codturno & "  where nsolicitud = " & solicitud)

                '    If priori <> 1 Then
                '        'actualizo el codturno a la prioridad 1
                '        con.Ejecuta("update solicitudes_turnos set prioridad = 9 where prioridad = 1 and nsolicitud = " & solicitud)
                '        con.Ejecuta("update solicitudes_turnos set prioridad = 1 where codturno =  " & codturno & "  where nsolicitud = " & solicitud)
                '        con.Ejecuta("update solicitudes_turnos set prioridad =  " & priori & " where prioridad = 9 and nsolicitud = " & solicitud)
                '    Else
                '        ' SI ES EL MISMO NADA 
                '    End If

                'End If



            Else

                CodPrecio = con.ejecuta1v_long("select codprecio from Solicitudes_TurnosSimulacion where nsolicitud =" & solicitud & " and CodTurno = " & codturno & "")
                con.Ejecuta("Update SOLICITUDESSimulacion set codprecio=" & CodPrecio & " where nsolicitud=" & solicitud & "")

            End If
        Else
            If Real = True Then
                If EstadoSoli = "P" Then
                    'estadosoli = 'P' then elimino si tiene codprecio
                    con.Ejecuta("delete precios where nsolicitud=" & solicitud & " and especial = 1 ")
                End If
            End If
        End If




            If Real = True Then
                If ListaEspera > 0 Then
                    ' con.Ejecuta("UPDATE SOLICITUDES SET             FechaCarta =" & Util.fechasql(Today) & " WHERE NSolicitud=" & solicitud & "")
                    '  envía email lista espera
                con.Ejecuta("UPDATE SOLICITUDES SET Codcolonia =NUll WHERE NSolicitud=" & solicitud & "")

                'con.Ejecuta("delete SOLICITUDESCARTA  where nsolicitud =" & solicitud)
                'con.Ejecuta("INSERT INTO SOLICITUDESCARTA  (Nsolicitud,tipo) values(" & solicitud & "," & v("P") & ")")

                ElseIf EstadoSoli = "A" Then

                    'con.Ejecuta("UPDATE SOLICITUDES SET             FechaCarta =" & Util.fechasql(Today) & " WHERE NSolicitud=" & solicitud & "")
                '  enviaremail(solicitud, sess, appli, con)
                con.Ejecuta("delete SOLICITUDESCARTA  where nsolicitud =" & solicitud)
                con.Ejecuta("INSERT INTO SOLICITUDESCARTA  (Nsolicitud,tipo) values(" & solicitud & "," & v("C") & ")")
                End If

            End If



    End Sub


    Public Function BuscaTurno(solicitud As Long, sexov As Integer, sexom As Integer, ByRef codcolonia As Long, ByRef codturno As Long, ByRef regimen As String, ByRef forzado As Integer, con As ConexionSQL, Real As Boolean) As Integer
        Dim encontrado As Integer = 0
        Dim PCodColonia As Long
        Dim PCodTurno As Long
        Dim PRegimen As String = VACIO
        Dim HayPlaza As Integer
        Dim cadena As String
        forzado = 0
        Dim dt As DataTable
        Dim dr As DataRow
        If Real = True Then
            cadena = "select codcolonia, codturno,regimen,prioridad from solicitudes_turnos where nsolicitud =" & solicitud & _
                   " and codturno in ( select codturno from  turnos where (TURNOS.NumPlazasFem + TURNOS.NumPlazasFemEx + TURNOS.NumPlazasMasc + TURNOS.NumPlazasMascEx > 0)) order by prioridad"
        Else
            cadena = "select codcolonia, codturno,regimen,prioridad from solicitudes_turnossimulacion where nsolicitud =" & solicitud & _
                 " and codturno in ( select codturno from  turnossimulacion as turnos where (TURNOS.NumPlazasFem + TURNOS.NumPlazasFemEx + TURNOS.NumPlazasMasc + TURNOS.NumPlazasMascEx > 0)) order by prioridad"
        End If

        dt = con.SelectSQL(cadena)

        For Each dr In dt.Rows
            If CStr(dr("prioridad")) = 1 Then
                PCodTurno = CStr(dr("CodTurno"))
                PCodColonia = CStr(dr("CodColonia"))
                PRegimen = CStr(dr("Regimen"))
            End If
            codturno = CStr(dr("CodTurno"))
            codcolonia = CStr(dr("CodColonia"))
            regimen = CStr(dr("Regimen"))
            HayPlaza = HayPlazas(solicitud, sexov, sexom, codturno, regimen, codcolonia, con, Real)
            If HayPlaza = 1 Then
                encontrado = 1
                Exit For
            Else
                forzado = 1
            End If
        Next

        If encontrado = 0 Then
            codcolonia = PCodColonia
            codturno = PCodTurno
            regimen = PRegimen
        End If
        Return encontrado
    End Function

    Public Sub RechazoPlaza(solicitud As Long, codcolonia As Long, codturno As Long, regimen As String, con As ConexionSQL, Real As Boolean)
        Dim CodColoniaAnt As Long
        Dim CodTurnoAnt As Long
        Dim RegimenAnt As String = VACIO
        Dim ListaEsperaAnt As String = VACIO
        Dim CodDormitorioAnt As String = VACIO
        Dim sexo As String = VACIO
        Dim cadena As String
        Dim fecha As Date
        Dim sexov As Integer
        Dim sexom As Integer

        Dim dt As DataTable
        Dim dr As DataRow
        If Real = True Then
            cadena = "select * from solicitudes where nsolicitud =" & solicitud
        Else
            cadena = "select * from solicitudesSimulacion where nsolicitud =" & solicitud
        End If
        dt = con.SelectSQL(cadena)
        For Each dr In dt.Rows

            'CodColoniaAnt = CStr(dr("CodColonia").ToString)
            'CodTurnoAnt = CStr(dr("CodTurno").ToString)
            'CAMBIO REVISION ENMA

            If CStr(dr("CodColonia").ToString) = "" Then
                CodColoniaAnt = 0
            Else
                CodColoniaAnt = CLng(CStr(dr("CodColonia").ToString))
            End If
            If CStr(dr("CodTurno").ToString) = "" Then
                CodTurnoAnt = 0
            Else
                CodTurnoAnt = CLng(CStr(dr("CodTurno").ToString))
            End If


            RegimenAnt = CStr(dr("Regimen").ToString)
            CodDormitorioAnt = CStr(dr("CodDormitorio").ToString)
            ListaEsperaAnt = CStr(dr("ListaEspera").ToString)
            sexo = CStr(dr("Sexo").ToString)
        Next

        If CodColoniaAnt <> 0 Then

            DesasignaSolicitud(solicitud, sexo, CodColoniaAnt, CodTurnoAnt, RegimenAnt, CodDormitorioAnt, ListaEsperaAnt, con, Real)
        End If

        fecha = Today

        If sexo = "V" Then sexov = 1 : sexom = 0 Else sexov = 0 : sexom = 1


        EstadoSolicitud(solicitud, con, Real, codcolonia, , regimen, , , , , , codturno)

        If Real = True Then
            If regimen = "I" Then
                cadena = "UPDATE TURNOS SET RvaMasc = RvaMasc + " & sexov & ", RvaFem = RvaFem + " & sexom
            Else
                cadena = "UPDATE TURNOS SET RvaMascex = RvaMascex + " & sexov & ", RvaFemex = RvaFemex + " & sexom
            End If
        Else

            If regimen = "I" Then
                cadena = "UPDATE TURNOSSimulacion SET RvaMasc = RvaMasc + " & sexov & ", RvaFem = RvaFem + " & sexom
            Else
                cadena = "UPDATE TURNOSSimulacion SET RvaMascex = RvaMascex + " & sexov & ", RvaFemex = RvaFemex + " & sexom
            End If
        End If

        cadena = cadena & " where codcolonia = " & codcolonia & " and codturno = " & codturno

        con.Ejecuta(cadena)

    End Sub

    Public Sub DesasignaSolicitud(solicitud As Long, sexo As String, codcolonia As Long, codturno As Long, regimen As String, CodDormitorio As String, listaespera As Integer, con As ConexionSQL, Real As Boolean)
        Dim cadena As String = VACIO
        Dim sexov As Integer
        Dim sexom As Integer
        'observaciones = "DESASIGNACION"
        'TODO LOGININS

        EstadoSolicitud(solicitud, con, Real, , , , , , "N", , , , "P")

        If sexo = "V" Then sexov = 1 : sexom = 0 Else sexov = 0 : sexom = 1

        If codturno <> 0 Then
            If Real = True Then
                If regimen = "I" Then
                    cadena = " UPDATE TURNOS SET LibresMasc = LibresMasc + " & sexov & ", LibresFem = LibresFem + " & sexom
                Else
                    cadena = " UPDATE TURNOS SET LibresMascex = LibresMascex + " & sexov & ", LibresFemex = LibresFemex + " & sexom
                End If
            Else
                If regimen = "I" Then
                    cadena = " UPDATE TURNOSsimulacion SET LibresMasc = LibresMasc + " & sexov & ", LibresFem = LibresFem + " & sexom
                Else
                    cadena = " UPDATE TURNOSsimulacion SET LibresMascex = LibresMascex + " & sexov & ", LibresFemex = LibresFemex + " & sexom
                End If
            End If
            cadena = cadena & " where codcolonia = " & codcolonia & " and codturno = " & codturno
            con.Ejecuta(cadena)
        End If
        If listaespera <> 0 Then
            If Real = True Then
                If regimen = "I" Then
                    cadena = " UPDATE TURNOS SET RvaMasc = RvaMasc - " & sexov & ", RvaFem = RvaFem - " & sexom
                Else
                    cadena = " UPDATE TURNOS SET RvaMascex = RvaMascex - " & sexov & ", RvaFemex = RvaFemex - " & sexom
                End If
            Else
                If regimen = "I" Then
                    cadena = " UPDATE TURNOSSimulacion SET RvaMasc = RvaMasc - " & sexov & ", RvaFem = RvaFem - " & sexom
                Else
                    cadena = " UPDATE TURNOSSimulacion SET RvaMascex = RvaMascex - " & sexov & ", RvaFemex = RvaFemex - " & sexom
                End If
            End If

            cadena = cadena & " where codcolonia = " & codcolonia & " and codturno = " & listaespera
            con.Ejecuta(cadena)
        End If
            If CodDormitorio <> VACIO Then
            DormitoriosSolicitud(solicitud, codcolonia, codturno, CodDormitorio, con, Real, 1)
            End If

    End Sub

    Public Sub CreaTablasSimulacion(con As ConexionSQL)

        con.Ejecuta("IF OBJECT_ID('SolicitudesSimulacion', 'U') IS NOT NULL drop table SolicitudesSimulacion")
        con.Ejecuta("IF OBJECT_ID('DormitoriosSimulacion', 'U') IS NOT NULL    drop table DormitoriosSimulacion")
        con.Ejecuta("IF OBJECT_ID('TurnosSimulacion', 'U') IS NOT NULL    drop table TurnosSimulacion")
        con.Ejecuta("IF OBJECT_ID('Solicitudes_TurnosSimulacion', 'U') IS NOT NULL    drop table Solicitudes_TurnosSimulacion")
        con.Ejecuta("select * into SolicitudesSimulacion from SOLICITUDES")
        con.Ejecuta("select * into DormitoriosSimulacion from Dormitorios")
        con.Ejecuta("select * into TurnosSimulacion from Turnos")
        con.Ejecuta("select * into Solicitudes_TurnosSimulacion from SOLICITUDES_Turnos")


    End Sub
    Public Sub DormitoriosSolicitud(solicitud As Long, codcolonia As Long, codturno As Long, CodDormitorio As String, con As ConexionSQL, Real As Boolean, Optional DesAsigna As Integer = 0)
        Dim cadena As String = VACIO

        If DesAsigna = 0 Then
            If Real = True Then
                cadena = "update solicitudes set coddormitorio = " & v(CodDormitorio) & " where nsolicitud  = " & solicitud
            Else
                cadena = "update solicitudesSimulacion set coddormitorio = " & v(CodDormitorio) & " where nsolicitud  = " & solicitud
            End If
            con.Ejecuta(cadena)

            'observaciones = "Asignado dormitorio a solicitud"
            'TODO LOGININS
            If Real = True Then
                cadena = "update dormitorios set camaslibres = camaslibres - 1 where  codcolonia = " & codcolonia & " and codturno = " & codturno & " and coddormitorio = " & v(CodDormitorio)
            Else
                cadena = "update dormitoriosSimulacion set camaslibres = camaslibres - 1 where  codcolonia = " & codcolonia & " and codturno = " & codturno & " and coddormitorio = " & v(CodDormitorio)
            End If
            con.Ejecuta(cadena)

            'observaciones = "Asignado solicitud a dormitorio"
            'TODO LOGININS
        Else
            If Real = True Then
                cadena = "update solicitudes set coddormitorio = '' where nsolicitud  = " & solicitud

            Else
                cadena = "update solicitudesSimulacion set coddormitorio = '' where nsolicitud  = " & solicitud
            End If
            con.Ejecuta(cadena)

            'observaciones = "Deasignado dormitorio a solicitud"
            'TODO LOGININS
            If CodDormitorio.Trim <> "" Then
                If Real = True Then
                    cadena = "update dormitorios set camaslibres = camaslibres + 1 where  codcolonia = " & codcolonia & " and codturno = " & codturno & " and coddormitorio = " & v(CodDormitorio)
                Else
                    cadena = "update dormitoriosSimulacion set camaslibres = camaslibres + 1 where  codcolonia = " & codcolonia & " and codturno = " & codturno & " and coddormitorio = " & v(CodDormitorio)
                End If
                con.Ejecuta(cadena)
            End If

                'observaciones = "Deasignado solicitud a dormitorio"
                'TODO LOGININS
            End If
    End Sub

End Module
