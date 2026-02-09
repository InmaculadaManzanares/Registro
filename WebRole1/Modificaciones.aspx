<%@ Page Language="vb" MasterPageFile="~/PaginaPrincipal.master" AutoEventWireup="false" CodeBehind="Modificaciones.aspx.vb" Inherits="NuevaPlantilla.Modificaciones" title="Modificaciones" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <script language="javascript" type="text/javascript">

        function activar_calendario(texto) {
            $(texto).datepicker({
                format: "dd/mm/yyyy"
            })
                .on('changeDate', function (ev) {
                    $(texto).datepicker('hide');
                });
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
        function stopRKey(evt) {
            var evt = (evt) ? evt : ((event) ? event : null);
            var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
            if ((evt.keyCode == 13) && (node.type == "text")) { return false; }
        }
        document.onkeypress = stopRKey;

        function ocultar_mensaje() {

            // Ocultamos el div
            $('#mensaje-popup').hide();

            // Eliminamos el mensaje del div(Etiqueta <p>)
            // Para que el proximo mensaje no sea incremental, sino único.
            document.getElementById("mensaje-popup").removeChild(document.getElementById("msgSolicitud"));

        }


        function limpiar_campos() {
            document.getElementById("ContentPlaceHolder1_txt_valorescampo").value = "";
        }

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
                <div class="col-md-12">
                    <div class="well well-sm">
                        <label class ="label-rosa-filtro">Filtros</label>  
                        <div class="form-inline">

                            <div class="form-group">
                                <asp:DropDownList ID="combo_campo" runat="server" AutoPostBack="True" class="form-control"></asp:DropDownList>
                            </div>
                            <div class="form-group">
                                <asp:DropDownList ID="combo_operador" runat="server" AutoPostBack="True" class="form-control">
                                    <asp:ListItem>=</asp:ListItem>
                                    <asp:ListItem>&lt;</asp:ListItem>
                                    <asp:ListItem>&gt;</asp:ListItem>
                                    <asp:ListItem Value="&gt;=">=&gt;</asp:ListItem>
                                    <asp:ListItem>&lt;=</asp:ListItem>
                                    <asp:ListItem Value="LIKE">Contiene</asp:ListItem>
                                    <asp:ListItem Value="&lt;&gt;">Distinto</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                            <div class="form-group">
                                <asp:TextBox ID="txt_valorescampo" runat="server" class="form-control" ></asp:TextBox>
                                <asp:TextBox ID="TextBox2" value="" runat="server"  Text=""  Width="0" BorderStyle="None" ReadOnly="True" Font-Size="XX-Small" BackColor="#EEEEEE"></asp:TextBox>
                            </div>
                            <asp:Button ID="Filtrar" formnovalidate runat="server" CausesValidation="False" Text="Filtrar" class="btn btn-normal" />
                            <asp:Button ID="Limpiar" formnovalidate runat="server" Text="Borrar Filtro" class="btn btn-normal" />
                            <asp:LinkButton ID="nuevo"
                                    runat="server"
                                    CssClass="btn btn-success-docBox"> <i class="glyphicon glyphicon-plus"></i>Nuevo Registro </asp:LinkButton>

                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Panel de Tabulación -->
        <div role="tabpanel" id="myTab">
           <%-- <div id="mensaje-popup" class="alert alert-danger">
                <p id='msgSolicitud' ></p>
                <a href="#" class="close" data-dismiss="alert">&times;</a>
            </div>--%>
            <asp:Panel ID="Panel1" runat="server">
                <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                    <ContentTemplate>
            <div id="mensaje-popup" class="alert mensaje-popup">
                <p id='msgSolicitud'></p>
                <a href="#" class="close" data-dismiss="alert">&times;</a>
            </div>
            <ul class="nav nav-tabs" role="tablist">
                <li role="presentation" class="active"><a href="#lista" aria-controls="lista" role="tab" data-toggle="tab">Lista</a></li>
                <li role="presentation"><a href="#detalle" aria-controls="detalle" role="tab" data-toggle="tab">Detalle</a></li>
            </ul>
            <!-- Tab panes -->
            
                   <div class="tab-content">
                            <div class="tab-pane active" id="lista">
                                <div class="row">
                                    <div class="col-md-12 table-responsive">

                                        <asp:SqlDataSource ID="SqlDS" runat="server"></asp:SqlDataSource>
                                        <asp:SqlDataSource ID="SqlCombousuario" runat="server"></asp:SqlDataSource>                                    
                                     
                                        <asp:GridView ID="matriz" runat="server" CssClass="table table-hover table-condensed" GridLines="None" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" DataKeyNames="ID" DataSourceID="SqlDS">
                                            <RowStyle />
                                            <EmptyDataRowStyle />
                                            <FooterStyle />
                                            <SelectedRowStyle />
                                            <EditRowStyle />
                                            <AlternatingRowStyle />
                                            <Columns>
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="btnEdit" runat="server" CssClass="btn btn-sm" OnClick="btnEdit_Click">
                                                            <i class="glyphicon glyphicon-folder-open"></i>                               
                                                        </asp:LinkButton>
                                                        <asp:LinkButton ID="btnDelete" runat="server" CssClass="btn btn-sm" Onclick="btnDelete_Click" OnClientClick="javascript:return confirm('¿Está seguro que desea borrar?');">
                                                            <i class="glyphicon glyphicon-remove"></i>
                                                       </asp:LinkButton>
                                                    </ItemTemplate>
                                                    <ItemStyle />
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="ID" HeaderText="ID" ReadOnly="True" SortExpression="ID">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Email" HeaderText="Email" ReadOnly="True" SortExpression="email">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Nombre" HeaderText="Nombre" SortExpression="Nombre">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                     <asp:BoundField DataField="empresa" HeaderText="Empresa" SortExpression="empresa">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                     <asp:BoundField DataField="centro" HeaderText="Centro" SortExpression="centro">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Fecha" HeaderText="Fecha" SortExpression="Fecha">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="inOut" HeaderText="Entrada/Salida" SortExpression="inOut">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Oficina" HeaderText="Lugar" SortExpression="Oficina">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Obs" HeaderText="Observaciones" SortExpression="Obs">
                                                <ItemStyle />
                                                </asp:BoundField>
                                            </Columns>
                                            <HeaderStyle BackColor="#ebebeb" />
                                            <PagerStyle CssClass="pagination-ys" />
                                        </asp:GridView>
                                    </div>
                                </div>
                            </div>

                            <%-- tap detalle--%>
                             <div class="tab-pane" id="detalle">                                
                                <div class="row" style="padding-top: 10px;">
                                    <div id="lg" class="form-horizontal">


                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-4 control-label pq">id</label>
                                                <div class="col-md-3">
                                                    <asp:TextBox ID="txtid"  MaxLength="100" type="text" runat="server" autocomplete="off" CssClass="form-control"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-4 control-label pq">Usuario</label>
                                                <div class="col-md-8">
                                                    <asp:DropDownList ID="cmbusuario" required="true"  AutoPostBack="true" runat="server" CssClass="form-control pq">
                                                      </asp:DropDownList>                                                
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-4 control-label pq">Fecha</label>
                                                <div class="col-md-3">
                                                    <asp:TextBox ID="txtfecha" onKeyUp="this.value=formateafecha(this.value);" MaxLength="10" required="true" type="text" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                 <asp:label runat="server" ID="lbldias" CssClass ="col-md-2 control-label pq">Hora</asp:label>
                                                <div class="col-md-3">
                                                    <asp:TextBox ID="txthora" required="true" type="text" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-4 control-label pq">Entrada/Salida</label>
                                                <div class="col-md-3">
                                                    <asp:DropDownList ID="EntradaSalida" AutoPostBack="true" required="true"  runat="server" CssClass="form-control pq">
                                                          <asp:ListItem Value=""></asp:ListItem>
                                                          <asp:ListItem Value="Entrada">Entrada</asp:ListItem>
                                                          <asp:ListItem Value="Salida">Salida</asp:ListItem>
                                                     </asp:DropDownList>                                                
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-4 control-label pq">Lugar</label>
                                                <div class="col-md-5">
                                                   <asp:TextBox ID="txtlugar" required="true" text ="Oficina" type="text" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                 </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-4 control-label pq">Observaciones</label>
                                                <div class="col-md-8">
                                                    <asp:TextBox ID="txtobs" required="true" MaxLength="250" type="text" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>                                      
                                   
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <div class="col-md-12 pq" align="center">
                                                    <asp:Button ID="btnsave" runat="server" OnClick="btnsave_Click" Text="Guardar" class="btn btn-sm btn-success-docBox" />
                                                    <asp:Button ID="btncancel" formnovalidate runat="server" Text="Cancelar" class="btn btn-sm btn-cancel-docBox" />
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <div class="col-xs-4 col-xs-offset-4 pq">
                                                    <%--<asp:Label ID="mensaje" style ="white-space: normal;" multiline ="true"  runat="server" CssClass="label"></asp:Label>--%>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                    </ContentTemplate>
                    <Triggers>
                        <asp:AsyncPostBackTrigger ControlID="Matriz" EventName="Sorting" />
                        <asp:AsyncPostBackTrigger ControlID="btnSave" EventName="Click" />
                    </Triggers>
                </asp:UpdatePanel>
            </asp:Panel>
        </div>
    </div>
</asp:Content>
   








