<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RegistroHorario.aspx.vb" Inherits="NuevaPlantilla.RegistroHorario" %>


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
    <link href="css/Login.css" rel="stylesheet" type="text/css" />
    <title>Registro</title>
    <script type="text/javascript">

        // Muestra el mensaje de Alerta
        function MostrarAlerta() {
       
            document.getElementById("alert").style.display = "block";
        }
        function pageLoad(sender, args) {
            if (!args.get_isPartialLoad()) {
                //  add our handler to the document's
                //  keydown event
                $addHandler(document, "keydown", onKeyDown);
            }
        }

    </script>

    <style type="text/css">
        .auto-style1 {
            width: 228px;
            height: 120px;
        }
    </style>

</head>
<body style="background:#252149;" >
    
    <div class="container" >
        <br />

        <form id="form1" runat="server" defaultbutton="ImageButton1" style="min-width:300px;">
            <div class="login">
                <div align="center">
                 <br />
                             <br />
                    <%--<img alt="" class="auto-style1" src="ImagenesDocbox/logotipogrupoland.png" /></div>--%>
                   <img alt="Logotipo"  src="ImagenesDocbox/logocorielespaceland.png" data-at2x="ImagenesDocbox/logocorielespaceland.png" />
                <br />
                <br />
          
                <div id="alert" class="alert alert-danger" style="display:none;">
                  <button type="button" class="close" data-dismiss="alert">&times;</button>
                  <asp:Label ID="mensaje" style ="white-space: normal;" multiline ="true"  runat="server" CssClass="label"></asp:Label>
                </div>
                    <asp:Panel ID ="panel1" runat="server" >
                        <div class="inner-addon left-addon">
                            <i class="glyphicon glyphicon-user"></i>
                            <asp:TextBox ID="txtusuario" runat="server" CssClass="form-control" placeholder="Usuario"></asp:TextBox>
                        </div>
                        <div class="inner-addon left-addon">
                            <i class="glyphicon glyphicon-lock"></i>
                            <asp:TextBox ID="txtcontraseña" runat="server" CssClass="form-control" TextMode="Password" placeholder="Password"></asp:TextBox>
                        </div>
                        <br />                        
                        <div align="center">
                            &nbsp;</div> 
                        <center>
                            <asp:LinkButton ID="ImageButton1"
                                runat="server"
                                CssClass="btn btn-login btn-sm" color="#fff"   background-color="#21B2BA"  border-color="#21B2BA" >             
                                <i class="glyphicon glyphicon-off"></i> Acceso
                            </asp:LinkButton>
                        </center>  
                        </asp:Panel> 
                    <!-- Link Cambiar/Restablecer -->
                    <div style="margin-top:10px; text-align:center;">
                        <asp:LinkButton ID="lnkReset" runat="server" CausesValidation="False">
                            Cambiar/Restablecer contraseña
                        </asp:LinkButton>
                    </div>

                    <!-- Panel reset (oculto al inicio) -->
                    <asp:Panel ID="panelReset" runat="server" Visible="false" style="margin-top:15px;">
                      <div id="divResetEmail" runat="server" class="inner-addon left-addon">
                            <i class="glyphicon glyphicon-envelope"></i>
                            <asp:TextBox ID="txtResetEmail" runat="server"
                                CssClass="form-control"
                                placeholder="Correo electrónico"></asp:TextBox>
                        </div>

                        <center style="margin-top:10px;">
                            <asp:Button ID="btnResetEnviar" runat="server"
                                CssClass="btn btn-login btn-sm"
                                Text="Enviar enlace" />
                        </center>

                        <center style="margin-top:10px;">
                            <asp:LinkButton ID="lnkVolverAcceso" runat="server" CausesValidation="False">
                                Volver al acceso
                            </asp:LinkButton>
                        </center>
                    </asp:Panel>









                         <asp:Panel ID ="panel2" runat="server" >
                        <div class="inner-addon left-addon">                       
                            <asp:label ID="txtnombre" runat="server" style="font-weight: bold;font-size: 18px;" >

                            </asp:label>
                             <br />
                               <br />
                            <asp:label runat="server" id="lblubicacion" style="">Ubicación</asp:label>
                            <asp:label ID="txtubicacion" visible="false"  Text ="Web" runat="server" style="font-weight: bold;font-size: 18px;" ></asp:label>
                        </div>
                     
                        <div class="inner-addon left-addon">
                            <asp:TextBox ID="txtlugar" runat="server"  CssClass="form-control"  text="Oficina"></asp:TextBox>
                        </div>
                        <br />
                        <div align="center">
                            </div> 
                        <center>
                            <asp:Button ID="EntradaSalida"
                                runat="server"
                                CssClass="btn btn-login btn-sm" color="#fff" style="padding-left: 30px;padding-right: 30px;" Text="Entrada" background-color="#21B2BA"  border-color="#21B2BA" >             
                             </asp:Button>
                        </center> 
               </asp:Panel>
                <div align="center">
                    &nbsp;</div>  
                <div align="center">
                    &nbsp;<asp:Button ID="list"
                                runat="server"
                                CssClass="btn btn-login btn-sm" color="#fff" style="padding-left: 15px;padding-right: 15px;" Visible ="false" Text="Listar Registro" background-color="#21B2BA"  border-color="#21B2BA" >             
                             </asp:Button>
                        </div>              
                </div>
            </div> 
        </form>

        <script src="scripts/jquery-2.1.3.min.js" type="text/javascript"></script>
        <script src="scripts/bootstrap.min.js" type="text/javascript"></script>

        <asp:SqlDataSource ID="SqlDSCentro" runat="server"></asp:SqlDataSource>
    </div>
</body>
</html>
