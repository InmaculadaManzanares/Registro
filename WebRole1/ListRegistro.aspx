<%@ Page Language="VB" MasterPageFile="~/PaginaPrincipal.master" AutoEventWireup="false" codeBehind="ListRegistro.aspx.vb" Inherits="NuevaPlantilla.ListRegistro" title="Listar Registros" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" defaultbutton="Filtrar">
    <script language="javascript" type="text/javascript">


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



        function pageLoad(sender, args) {
            if (!args.get_isPartialLoad()) {
                //  add our handler to the document's
                //  keydown event
                $addHandler(document, "keydown", onKeyDown);
            }
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
                if (IsNumeric(dia) == false)
                { fecha = ""; }
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

    <!--Cabecera-->
    <div class="container theme-showcase" role="main" style="padding-right: 0px; padding-left: 0px;">
        <div class="page-header">
            <div class="row">
                <asp:ScriptManagerProxy ID="ScriptManagerProxy2" runat="server"></asp:ScriptManagerProxy>
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
                                <div class="col-md-2">
                                    <label>Empresa</label>
                                    <asp:DropDownList ID="comboempresa" runat="server" class="form-control" AutoPostBack="True"></asp:DropDownList>
                                </div>
                               <div class="col-md-2">
                                    <label>Centro</label>
                                    <asp:DropDownList ID="combocentro" runat="server" class="form-control" AutoPostBack="True"></asp:DropDownList>
                                </div>
                                <div class="col-md-2">
                                    <label>Nombre</label>
                                    <asp:DropDownList ID="combonombre" runat="server" class="form-control" AutoPostBack="True"></asp:DropDownList>
                                </div>
                                <div class="col-md-2">
                                    <label>Situación</label>
                                    <asp:DropDownList ID="informes" runat="server" AutoPostBack="True" CssClass="form-control">
                                          <asp:ListItem Value=""></asp:ListItem>
                                                          <asp:ListItem Value="Vacaciones">Vacaciones</asp:ListItem>
                                                          <asp:ListItem Value="Baja">Baja</asp:ListItem>
                                                          <asp:ListItem Value="Otros">Otros</asp:ListItem>
                                    </asp:DropDownList>
                                </div>

                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-sm-12">
                    <div class="form-group">
                        <asp:Button ID="Filtrar" formnovalidate runat="server" CausesValidation="False" Text="buscar" class="btn btn-normal" />
                        <asp:Button ID="Limpiar" formnovalidate runat="server" Text="Limpiar" class="btn btn-normal" />
                        <asp:Button ID="Excel" formnovalidate runat="server" Text="Exportar Excel" class="btn btn-normal" />
                        <asp:Button ID="ExcelD" formnovalidate runat="server" Text="Exportar Excel Detalle" class="btn btn-normal" />

                    </div>
                </div>
            </div>
            <div class="row">
                <div class="col-sm-12">
                    <asp:Panel ID="mensajepopup" runat="server" Visible="true">
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
                                <li role="presentation"><a href="#detalle" aria-controls="detalle" role="tab" data-toggle="tab">Detalle</a></li>
                            </ul>
                        
                            <asp:Panel ID="Panel1" runat="server" >
                                <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                                    <ContentTemplate>
                                        <div class="tab-content">
                                            <%-- tap lista--%>
                                            <div class="tab-pane active" id="lista">
                                                <div class="row">
                                                    <div class="col-md-12 table-responsive" >
                                                        <asp:SqlDataSource ID="SqlDS" runat="server"></asp:SqlDataSource>
                                                        <asp:SqlDataSource ID="Sqlempresa" runat="server"></asp:SqlDataSource>
                                                        <asp:SqlDataSource ID="Sqlcentro" runat="server"></asp:SqlDataSource>
                                                        <asp:SqlDataSource ID="Sqlemail" runat="server"></asp:SqlDataSource>
                                                        <asp:SqlDataSource ID="SqlDetail" runat="server"></asp:SqlDataSource>
                                                        <center>
                                                            <asp:Label ID="resultado" runat="server" Font-Bold="True"></asp:Label>
                                                        </center>

                                                        <asp:GridView ID="matriz"  runat="server" BorderStyle="Solid" BorderWidth="1px" CssClass="table table-hover table-condensed" AutoGenerateColumns="False" GridLines="None" AllowSorting="True" AllowPaging="true"  DataKeyNames="Fecha" DataSourceID="SqlDS">

                                                            <RowStyle />
                                                            <HeaderStyle BackColor="Silver" />
                                                            <EmptyDataRowStyle />


                                                            <Columns>                                                                
                                                                <asp:BoundField DataField="Nempresa"   HeaderText ="Empresa" ReadOnly="True" SortExpression="Nempresa">
                                                                <ItemStyle />
                                                                </asp:BoundField>                                                                                                                             
                                                                <asp:BoundField DataField="Centro"   HeaderText ="Centro" ReadOnly="True" SortExpression="Centro">
                                                                <ItemStyle />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="Fecha" dataformatstring="{0:dd/MM/yyyy}"  HeaderText ="Fecha" ReadOnly="True" SortExpression="Fecha">
                                                                <ItemStyle />
                                                                </asp:BoundField>
                                                                 <asp:BoundField DataField="Nombre" HeaderText="Nombre" ReadOnly="True" SortExpression="Nombre">
                                                                <ItemStyle />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="Email" HeaderText="Email" ReadOnly="True" SortExpression="email">
                                                                <ItemStyle />

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
                                                                <asp:TemplateField>
                                                                    <ItemTemplate>
                                                                        <asp:LinkButton ID="btnEdit" BorderStyle ="Solid" BorderWidth ="1"
                                                                            runat="server"
                                                                            CssClass="btn btn-sm"
                                                                            OnClick="btnEdit_Click">Detalle                                                       
                                                                        </asp:LinkButton>
                                                                    
                                                                    </ItemTemplate>

                                                                    <ItemStyle  ForeColor ="DarkBlue"  />
                                                                </asp:TemplateField>


                                                            </Columns>


                                                            <FooterStyle />
<%--                                                            <PagerSettings Mode="NumericFirstLast" />
                                                            <PagerStyle CssClass="pagination-ys" HorizontalAlign="Center" />--%>
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

                                            <%-- tap detalle--%>
                                            <div class="tab-pane" id="detalle">
                                                <h2><small></small></h2>
                                                <div class="row">
                                                    <div class="col-md-12 table-responsive">
                                                        <asp:GridView ID="matrizdetail" runat="server" BorderStyle="Solid" BorderWidth="1px" CssClass="table table-hover table-condensed" GridLines="None" AllowSorting="True" DataKeyNames="fecha" DataSourceID="SqlDetail">
                                                            <RowStyle />
                                                            <HeaderStyle BackColor="Silver" />
                                                            <EmptyDataRowStyle />
                                                            <Columns>

                                                                <asp:TemplateField>
                                                                    <ItemTemplate>
                                                                    </ItemTemplate>
                                                                </asp:TemplateField>

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
                                        <asp:AsyncPostBackTrigger ControlID="matrizdetail" EventName="Sorting" />
                                    </Triggers>
                                </asp:UpdatePanel>
                            </asp:Panel>
                        </div>
                    </asp:Panel>
                </div>
            </div>
        </div>
    </div> 
    
</asp:Content>
   








