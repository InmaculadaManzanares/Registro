<%@ Page Language="VB" MasterPageFile="~/PaginaPrincipal.master" AutoEventWireup="false" CodeBehind="Consultas.aspx.vb" Inherits="NuevaPlantilla.Consultas" Title="Consultas" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">

<script language="javascript" type="text/javascript">

    function activapanellista() {
    $('#myTab a[href="#lista"]').show();
    $('#myTab a[href="#lista"]').tab('show');
    }
    
    function ocultarpanellista() {
        $('#myTab a[href="#lista"]').hide();
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

</script>
        <!--Cabecera-->
    <asp:SqlDataSource ID="SqlDS" runat="server" OldValuesParameterFormatString="original_{0}"></asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDSBus" runat="server" OldValuesParameterFormatString="original_{0}"></asp:SqlDataSource>
    <div class="container theme-showcase" role="main">
        <div class="page-header">
            <div class="row">
                <asp:ScriptManagerProxy ID="ScriptManagerProxy2" runat="server"></asp:ScriptManagerProxy>
                <div class="col-md-4">
                    <h3>Consultas SQL</h3>
                </div>     
                <div class="col-md-8">
                    <div class="well well-sm">
                        <div class="form-inline">
                            <div class="form-group">
                                <asp:Label ID="Label10" runat="server" Text="PANTILLA CONSULTA" CssClass ="control-label pq"></asp:Label>
                            </div>
                            <div class="form-group">
                                <asp:DropDownList ID="combo_consulta" runat="server" AutoPostBack="True"  class="form-control"> </asp:DropDownList>
                              </div>
                            <div class="form-group">
                                <asp:TextBox ID="nuevaconsulta" type="text" runat="server"  class="form-control" placeholder="Nombre Nueva Consulta"></asp:TextBox>
                            </div>
                            <asp:Button ID="btnSave" runat="server" Text="Grabar Consulta" cssclass="btn btn-sm btn-cancel-docBox" />
                        </div>
                    </div>
                </div>
                <div class="col-md-12">
                    <div class="well well-sm">
                        <div class="form-horizontal">
                              <div class="col-md-12">  
                                <div class="form-group">
                                   
                                        <asp:TextBox ID="consulta"  type="text" runat="server" TextMode="MultiLine" cssclass="form-control" placeholder="Introduzca Consulta SQL"></asp:TextBox><br />
                                   
                                </div>
                              </div>
                             
                                <div class="form-group">
                                   
                                         <asp:Label ID="Label1" runat="server" cssclass="col-md-1 control-label" Text="FILTRO1"></asp:Label>
                                         <div class="col-md-2">
                                             <asp:TextBox ID="t_filtro1"  type="text" runat="server" cssclass="form-control" ></asp:TextBox>
                                         </div>
                                         <asp:Label ID="Label2" runat="server" cssclass="col-md-1 control-label" Text="FILTRO2"></asp:Label>
                                         <div class="col-md-2">
                                            <asp:TextBox ID="t_filtro2"  type="text" runat="server" cssclass="form-control" ></asp:TextBox>
                                         </div>
                                        <asp:Button ID="BBuscar" formnovalidate runat="server" text="Ejecutar" onclick ="BBuscar_click"  cssclass="btn btn-sm btn-success-docBox" />
                                        <asp:Button ID="bLimpiar" formnovalidate runat="server" text="Limpiar" onclick ="bLimpiar_click" cssclass="btn btn-sm btn-cancel-docBox" />
                                        <asp:Button ID="ImageButton1" runat="server" text="Excel" onclick ="ImageButton1_Click" cssclass="btn btn-sm btn-cancel-docBox" />
                                        <asp:Button ID="salir" runat="server" text="Salir" onclick ="salir_Click" cssclass="btn btn-sm btn-cancel-docBox" />
                                        <asp:Label ID="numero" runat="server" cssclass="control-label pq" ></asp:Label>
                                   
                                </div>
                            
                        </div>
                    </div>
                 </div>

             </div>
        </div>
        <!-- Panel de Tabulación -->
        <div role="tabpanel" id="myTab">
            <div id="mensaje-popup" class="alert mensaje-popup">
                <a href="#" class="close" data-dismiss="alert">&times;</a>
            </div>
            <ul class="nav nav-tabs" role="tablist">
                <li role="presentation" class="active"><a href="#lista" aria-controls="lista" role="tab" data-toggle="tab">Lista</a></li>
            </ul>
            <!-- Tab panes -->
            <div class="tab-content">
                <div class="tab-pane active" id="lista">
                    <div class="row">
                        <div class="col-md-12 table-responsive"> 
                            <asp:GridView ID="matriz" runat="server" CssClass="table table-hover table-condensed" GridLines="None" AllowPaging="True" AllowSorting="True" DataSourceID="SqlDS">
                                <RowStyle />
                                <EmptyDataRowStyle />
                                <Columns>
                                </Columns>
                                <FooterStyle />
                                <PagerStyle CssClass="pagination-ys" />
                                <SelectedRowStyle />
                                <HeaderStyle />
                                <EditRowStyle />
                                <AlternatingRowStyle />
                            </asp:GridView>
                        </div>
                    </div>
                </div>
            </div> 
       </div>
    </div>

</asp:Content>

