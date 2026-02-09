<%@ Page Language="VB" AutoEventWireup="false" CodeBehind="InicioGestion.aspx.vb" Inherits="NuevaPlantilla.InicioGestion" title="Registro" %>

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
<body>

    <div class="container">

        <form id="form1" runat="server" defaultbutton="ImageButton1">
            <div class="login">
                <div align="center">
                 <br />
                    <%--<img alt="" class="auto-style1" src="ImagenesDocbox/logotipogrupoland.png" /></div>--%>
                   <img alt="Logotipo"  src="ImagenesDocbox/logocorielespaceland.png" data-at2x="ImagenesDocbox/logocorielespaceland.png" />
                <br />
                <br />
                <div align="center">
                    &nbsp;</div> 
                <center>
                <div id="alert" class="alert alert-danger">
                  <button type="button" class="close" data-dismiss="alert">&times;</button>
                  <asp:Label ID="mensaje" style ="white-space: normal;" multiline ="true"  runat="server" CssClass="label"></asp:Label>
                </div>
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
                        CssClass="btn btn-login btn-sm">             
                        <i class="glyphicon glyphicon-off"></i> Acceder
                    </asp:LinkButton>
                </center>  
                <div align="center">
                    &nbsp;</div>  
                <div align="center">
                    &nbsp;</div>              
            </div>
        </form>

        <script src="scripts/jquery-2.1.3.min.js" type="text/javascript"></script>
        <script src="scripts/bootstrap.min.js" type="text/javascript"></script>

        <asp:SqlDataSource ID="SqlDSCentro" runat="server"></asp:SqlDataSource>
    </div>
</body>
</html>
