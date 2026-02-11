<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ListRegistroEmail.aspx.vb" Inherits="NuevaPlantilla.ListRegistroEmail" title="Listar Registros" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link rel="shortcut icon" type="image/x-icon" href="ImagenesDocbox/favicon.ico" />
    <link rel="apple-touch-icon" href="apple-touch-icon.png" />
    <link rel="apple-touch-icon" sizes="114x114" href="apple-touch-icon.png" />
    <link rel="apple-touch-icon" sizes="72x72" href="apple-touch-icon.png" />
    <link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <link href="css/Master.css" rel="stylesheet" type="text/css" />
    <script src="scripts/jquery-2.1.3.min.js" type="text/javascript"></script>
    <script src="scripts/bootstrap.min.js" type="text/javascript"></script>

    <title>Registro</title>
    <script type="text/javascript">

        // Muestra el mensaje de Alerta
        function MostrarAlerta() {

            document.getElementById("alert").style.display = "block";
        }


        function activapanel() {
            // Activamos el panel detalle
            $('#myTab a[href="#detalle"]').show();

            // Hacemos que se convierta en el principal
            $('#myTab a[href="#detalle"]').tab('show');
        }

        function activapanellista() {
            // Activamos el panel lista
            $('#myTab a[href="#lista"]').tab('show');
        }

        function ocultarpanel(id_panel) {
            // Ocultamos el panel detalle
            $('#myTab a[href="#detalle"]').hide();
        }

        function mostrar_mensaje(mensaje) {

            document.getElementById("mensaje-popup").innerHTML += "<p id='msgSolicitud'>" + mensaje + "</p>";
            // Mostramos el mensaje
            $('#mensaje-popup').show();
        }


        function ocultar_mensaje() {

            // Ocultamos el div
            $('#mensaje-popup').hide();

            // Eliminamos el mensaje del div(Etiqueta <p>)
            // Para que el proximo mensaje no sea incremental, sino único.
            document.getElementById("mensaje-popup").removeChild(document.getElementById("msgSolicitud"));

        }

        function Solo_Numerico(variable) {
            Numer = parseInt(variable);
            if (isNaN(Numer)) {
                return "";
            }
            return Numer;
        }

        function ValNumero(Control) {
            Control.value = Solo_Numerico(Control.value);
        }


        function pulsar(e) {
            if (e.keyCode === 13 && !e.shiftKey) {
                e.preventDefault();
                var boton = document.getElementById("Filtrar");
                angular.element(boton).triggerHandler('click');
            }
        }

        function activar_calendario(texto) {
            $(texto).datepicker({
                format: "dd/mm/yyyy"
            })
                .on('changeDate', function (ev) {
                    $(texto).datepicker('hide');
                });
        }

        var primerslap = false;
        var segundoslap = false;
        function formateafecha(fecha) {
            var long = fecha.length;
            var dia;
            var mes;
            var ano;
            if ((long >= 2) && (primerslap == false)) {
                dia = fecha.substr(0, 2);
                if ((IsNumeric(dia) == true) && (dia <= 31) && (dia != "00")) { fecha = fecha.substr(0, 2) + "/" + fecha.substr(3, 7); primerslap = true; }
                else { fecha = ""; primerslap = false; }
            }
            else {
                dia = fecha.substr(0, 1);
                if (IsNumeric(dia) == false) { fecha = ""; }
                if ((long <= 2) && (primerslap = true)) { fecha = fecha.substr(0, 1); primerslap = false; }
            }
            if ((long >= 5) && (segundoslap == false)) {
                mes = fecha.substr(3, 2);
                if ((IsNumeric(mes) == true) && (mes <= 12) && (mes != "00")) { fecha = fecha.substr(0, 5) + "/" + fecha.substr(6, 4); segundoslap = true; }
                else { fecha = fecha.substr(0, 3);; segundoslap = false; }
            }
            else { if ((long <= 5) && (segundoslap = true)) { fecha = fecha.substr(0, 4); segundoslap = false; } }
            if (long >= 7) {
                ano = fecha.substr(6, 4);
                if (IsNumeric(ano) == false) { fecha = fecha.substr(0, 6); }
                else { if (long == 10) { if ((ano == 0) || (ano < 1900) || (ano > 2100)) { fecha = fecha.substr(0, 6); } } }
            }
            if (long >= 10) {
                fecha = fecha.substr(0, 10);
                dia = fecha.substr(0, 2);
                mes = fecha.substr(3, 2);
                ano = fecha.substr(6, 4);
                // Año no viciesto y es febrero y el dia es mayor a 28 
                if ((ano % 4 != 0) && (mes == 02) && (dia > 28)) { fecha = fecha.substr(0, 2) + "/"; }
            }
            return (fecha);
        }


    </script>

    <style type="text/css">
        .auto-style1 {
            width: 228px;
            height: 120px;
        }
    </style>

</head>

<body style="background:#252149;padding:10px" >
    <form id="form1" runat="server">
            <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePartialRendering="true">
        </asp:ScriptManager>
    <!--Cabecera-->
    <div class="container theme-showcase" role="main" style="padding-right: 0px; padding-left: 0px;">
        <div class="page-header">
            <div class="row">
                   <div class="col-sm-12">
                    <div class="well well-sm">
                        <div class="form-horizontal">
                            <div class="form-group">

                                <div class="col-md-12" style="padding-top :20px;">
                                    <center>
                                    <asp:label runat="server" ID="lblnombre" style="font-size:large;color:#252149;" ></asp:label>
                                        </center>
                                </div>                             
                            </div>
                        </div>
                    </div>
                </div>
             
                <div class="col-sm-12">
                    <div class="well well-sm">
                        <div class="form-horizontal">
                            <div class="form-group">

                                <div class="col-md-2">
                                    <label>Fecha Inicio</label>
                                    <asp:TextBox runat="server" onKeyUp="this.value=formateafecha(this.value);" onclick="activar_calendario('#ContentPlaceHolder1_fecha1')" type="date" min="1900-01-01" max="2100-01-01" ID="fecha1" MaxLength="10" CssClass="form-control" placeholder=""></asp:TextBox>
                                </div>
                                <div class="col-md-2">
                                    <label>Fecha Fin</label>
                                    <asp:TextBox runat="server" onKeyUp="this.value=formateafecha(this.value);" onclick="activar_calendario('#ContentPlaceHolder1_fecha2')" type="date" min="1900-01-01" max="2100-01-01" ID="fecha2" MaxLength="10" CssClass="form-control" placeholder=""></asp:TextBox>
                                </div>
                                <asp:HiddenField ID="txtnombre" runat="server" Value="valor inicial" />
                              
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-sm-12">
                    <div class="form-group">
                        <asp:Button ID="Filtrar" formnovalidate runat="server" CausesValidation="False" Text="buscar" class="btn btn-normal" />
                        <asp:Button ID="Limpiar" formnovalidate runat="server" Text="Limpiar" class="btn btn-normal" />
                        <asp:Button ID="Salir" formnovalidate runat="server" Text="Salir" class="btn btn-normal" />

                    </div>
                </div>
            </div>
            <div class="row">
                <div class="col-sm-12">
<%--                    <asp:Panel ID="mensajepopup" runat="server" Visible="true">
                        <asp:Label runat="server" ID="msgSolicitud" Text="" />
                    </asp:Panel>--%>
                    <asp:Panel ID="mensajepopup" runat="server" Visible="false" CssClass="alert alert-danger" style="margin-top:10px;">
                            <button type="button" class="close" data-dismiss="alert">&times;</button>
                            <asp:Label runat="server" ID="msgSolicitud" Text="" />
                        </asp:Panel>

                </div>
                <div class="col-sm-12">
                    <!-- Tab panes -->
                    <asp:Panel ID="panelresultado" Visible="false" runat="server">
                        <!-- Panel de Tabulación -->
                        <div role="tabpanel" id="myTab">
                            <ul class="nav nav-tabs" role="tablist">
                                <li role="presentation" class="active"><a href="#lista" aria-controls="Registro" role="tab" data-toggle="tab">Registro</a></li>
                            </ul>
                        
                            <asp:Panel ID="Panel1" runat="server" >
                                <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                                    <ContentTemplate>
                                        <div class="tab-content">
                                            <%-- tap lista--%>
                                            <div class="tab-pane active" id="lista">
                                                <div class="row">
                                                    <div class="col-md-12 table-responsive" style="border: 0px solid #ddd;" >
                                                        <asp:SqlDataSource ID="SqlDS" runat="server"></asp:SqlDataSource>
                                                        <asp:SqlDataSource ID="Sqlempresa" runat="server"></asp:SqlDataSource>
                                                        <asp:SqlDataSource ID="Sqlcentro" runat="server"></asp:SqlDataSource>
                                                        <asp:SqlDataSource ID="Sqlemail" runat="server"></asp:SqlDataSource>
                                                        <asp:SqlDataSource ID="SqlDetail" runat="server"></asp:SqlDataSource>
                                                        <center>
                                                            <asp:Label ID="resultado" runat="server" Font-Bold="True"></asp:Label>
                                                        </center>

                                                        <asp:GridView ID="matriz"  runat="server" BorderStyle="Solid"
                                                            background-color ="White"  BorderWidth="1px" 
                                                            style="background-color: white"
                                                            CssClass="table table-hover table-condensed"                                                            
                                                            AutoGenerateColumns="False" GridLines="None" AllowSorting="True" 
                                                            AllowPaging="true"  DataKeyNames="Fecha" DataSourceID="SqlDS">

                                                            <RowStyle />
                                                            <HeaderStyle BackColor="Silver" />
                                                            <EmptyDataRowStyle />


                                                            <Columns>                                                                
<%--                                                                <asp:BoundField DataField="Nempresa"   HeaderText ="Empresa" ReadOnly="True" SortExpression="Nempresa">
                                                                <ItemStyle />
                                                                </asp:BoundField>                                                                                                                             
                                                                <asp:BoundField DataField="Centro"   HeaderText ="Centro" ReadOnly="True" SortExpression="Centro">
                                                                <ItemStyle />
                                                                </asp:BoundField>--%>
                                                                <asp:BoundField DataField="Fecha" dataformatstring="{0:dd/MM/yyyy}"  HeaderText ="Fecha" ReadOnly="True" SortExpression="Fecha">
                                                                <ItemStyle />
                                                               <%-- </asp:BoundField>
                                                                 <asp:BoundField DataField="Nombre" HeaderText="Nombre" ReadOnly="True" SortExpression="Nombre">
                                                                <ItemStyle />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="Email" HeaderText="Email" ReadOnly="True" SortExpression="email">
                                                                <ItemStyle />--%>

                                                                </asp:BoundField>
                                                                     <asp:BoundField DataField="Dia" HeaderText="Día" ReadOnly="True" SortExpression="Dia">
                                                                <ItemStyle />
                                                                </asp:BoundField>
                                                                     <asp:BoundField DataField="Lugar" HeaderText="Lugar" ReadOnly="True" SortExpression="Lugar">
                                                                <ItemStyle />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="Entrada" HeaderText="Entrada" SortExpression="Entrada">
                                                                <ItemStyle />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="Salida" HeaderText="Salida" SortExpression="Salida">
                                                                <ItemStyle />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="Horas" HeaderText="Horas" SortExpression="Horas">
                                                                <ItemStyle />
                                                                </asp:BoundField>                                                               
<%--                                                                <asp:TemplateField>
                                                                    <ItemTemplate>
                                                                        <asp:LinkButton ID="btnEdit" BorderStyle ="Solid" BorderWidth ="1"
                                                                            runat="server"
                                                                            CssClass="btn btn-sm"
                                                                            OnClick="btnEdit_Click">Detalle                                                       
                                                                        </asp:LinkButton>
                                                                    
                                                                    </ItemTemplate>

                                                                    <ItemStyle  ForeColor ="DarkBlue"  />
                                                                </asp:TemplateField>--%>


                                                            </Columns>


                                                            <FooterStyle />

                                                            <PagerSettings Mode="NumericFirstLast" />
                                                            <PagerStyle CssClass="pagination-ys" HorizontalAlign="Center" />
                                                            <SelectedRowStyle />
                                                            <HeaderStyle />
                                                            <EditRowStyle />
                                                            <AlternatingRowStyle />
                                                        </asp:GridView>

                                                        <asp:SqlDataSource ID="SqlDSexcel" runat="server"></asp:SqlDataSource>


                                                        <asp:GridView ID="matrizexcel" Visible="false" runat="server" BorderStyle="Solid" BorderWidth="1px" CssClass="table table-hover table-condensed" GridLines="None" AllowSorting="True" DataKeyNames="fecha" DataSourceID="SqlDSexcel">
                                                            <RowStyle />
                                                            <HeaderStyle BackColor="Silver" />
                                                            <EmptyDataRowStyle />
                                                            <Columns>
                                                            </Columns>
                                                            <FooterStyle />

                                                            <SelectedRowStyle />
                                                            <HeaderStyle />
                                                            <EditRowStyle />
                                                            <AlternatingRowStyle />
                                                        </asp:GridView>





                                                    </div>
                                                </div>
                                            </div>

                                        </div>
                                    </ContentTemplate>
                                    <Triggers>
                                        <asp:AsyncPostBackTrigger ControlID="Matriz" EventName="Sorting" />
                        
                                    </Triggers>
                                </asp:UpdatePanel>
                            </asp:Panel>
                        </div>
                    </asp:Panel>
                </div>
            </div>
        </div>
    </div> 
    

      </form>
</body>
</html>







