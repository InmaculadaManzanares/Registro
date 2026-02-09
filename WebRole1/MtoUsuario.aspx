<%@ Page Language="VB" MasterPageFile="~/PaginaPrincipal.master" AutoEventWireup="false" codeBehind="MtoUsuario.aspx.vb" Inherits="NuevaPlantilla.MtoUsuario" title="Usuarios" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
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
                                    CssClass="btn btn-success-docBox"> <i class="glyphicon glyphicon-plus"></i>Nuevo Usuario </asp:LinkButton>

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
                                        <asp:SqlDataSource ID="SqlComboempresas" runat="server"></asp:SqlDataSource>
                                        <asp:SqlDataSource ID="SqlCombocentros" runat="server"></asp:SqlDataSource>
                                        <asp:SqlDataSource ID="SqlDSestado" runat="server"></asp:SqlDataSource>
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
                                                <asp:BoundField DataField="Verificado" HeaderText="Verificado" SortExpression="Verificado">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Administrador" HeaderText="Administrador" SortExpression="v">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Empresa" HeaderText="Empresa" SortExpression="Empresa">
                                                <ItemStyle />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Centro" HeaderText="Centro" SortExpression="Centro">
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
                                                <label class="col-md-3 control-label pq">id</label>
                                                <div class="col-md-3">
                                                    <asp:TextBox ID="txtid" required="true" MaxLength="100" type="text" runat="server" autocomplete="off" CssClass="form-control"></asp:TextBox>
                                                </div>                                          
                                                <div class="col-md-3">
                                                    <asp:checkbox ID="txtnotifica"  runat="server" Text="&nbsp;&nbsp;Notificaciones"  ></asp:checkbox>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">email</label>
                                                <div class="col-md-9">
                                                    <asp:TextBox ID="txtemail" required="true" MaxLength="100" type="text" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                         <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">password</label>
                                                <div class="col-md-9">
                                                    <asp:TextBox ID="txtPassword" required="true" MaxLength="100" type="text" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">nombre</label>
                                                <div class="col-md-9">
                                                    <asp:TextBox ID="txtnombre" required="true" MaxLength="100" type="text" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">Verificado</label>
                                                <div class="col-md-3">
                                                    <asp:DropDownList ID="cmbveri" AutoPostBack="true" runat="server" CssClass="form-control pq">
                                                        <asp:ListItem Value="Si">Si</asp:ListItem>
                                                        <asp:ListItem Value="No">No</asp:ListItem>
                                                    </asp:DropDownList>
                                                </div>
                                     
                                                    <label class="col-md-3 control-label pq">Administrador</label>
                                                    <div class="col-md-3">
                                                        <asp:DropDownList ID="cmbadmini" AutoPostBack="true" runat="server" CssClass="form-control pq">
                                                            <asp:ListItem Value="Si">Si</asp:ListItem>
                                                            <asp:ListItem Value="No">No</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </div>
                                                </div>                                          
                                        </div>
                                         <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">Empresa</label>
                                                <div class="col-md-3">
                                                    <asp:DropDownList ID="cmbempresa" AutoPostBack="true" runat="server" CssClass="form-control pq">
                                                          </asp:DropDownList>                                                
                                                </div>
                                    
                                                <label class="col-md-3 control-label pq">Centro</label>
                                                <div class="col-md-3">
                                                    <asp:DropDownList ID="cmbcentro" AutoPostBack="true" runat="server" CssClass="form-control pq">
                                                          </asp:DropDownList>                                                
                                                </div>
                                            </div>
                                        </div>
                                        <br />

                                       <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">LUNES </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Mañana hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="lunes1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Mañana hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="lunes2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                       <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq"> </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Tarde hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tlunes1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Tarde hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tlunes2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                       <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">MARTES </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Mañana hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="martes1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Mañana hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="martes2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>

                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq"> </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Tarde hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tmartes1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Tarde hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tmartes2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                       <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">MIERCOLES </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Mañana hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="miercoles1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Mañana hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="miercoles2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>

                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq"> </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Tarde hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tmiercoles1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Tarde hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tmiercoles2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                       <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">JUEVES </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Mañana hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="jueves1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Mañana hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="jueves2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>

                                            </div>
                                        </div>
                                         <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq"> </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Tarde hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tjueves1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Tarde hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tjueves2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">VIERNES </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Mañana hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="viernes1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Mañana hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="viernes2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>

                                            </div>
                                        </div>
                                         <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq"> </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Tarde hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tviernes1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Tarde hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tviernes2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                                                           <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">SABADO </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Mañana hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="sabado1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Mañana hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="sabado2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>

                                            </div>
                                        </div>
                                         <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq"> </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Tarde hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tsabado1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Tarde hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tsabado2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq">DOMINGO </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Mañana hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="domingo1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Mañana hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="domingo2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>

                                            </div>
                                        </div>
                                         <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <label class="col-md-3 control-label pq"> </label>
                                                <label class="col-md-3 control-label pq" style="text-align: left;">Tarde hora inicio:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tdomingo1"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                                <label class="col-md-2 control-label pq" style="text-align: left;">Tarde hora fin:</label>
                                                <div class="col-md-2">
                                                    <asp:TextBox ID="tdomingo2"  MaxLength="100" type="time" runat="server" CssClass="form-control pq"></asp:TextBox>
                                                </div>
                                            </div>
                                        </div>

                                        <div class="col-sm-12 col-lg-8">
                                            <div class="form-group">
                                                <div class="col-md-12 pq" align="center">
                                                    <asp:Button ID="btnsave" runat="server" OnClick="btnsave_Click" Text="Guardar" class="btn btn-sm btn-success-docBox" />
                                                    <asp:Button ID="btncancel" formnovalidate runat="server" Text="Cancelar" class="btn btn-sm btn-cancel-docBox" />
                                                    <asp:Button ID="btnemail" formnovalidate runat="server" Text="Enviar email con contraseña" class="btn btn-sm btn-cancel-docBox" />
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
   








