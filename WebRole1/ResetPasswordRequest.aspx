<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ResetPasswordRequest.aspx.vb" Inherits="NuevaPlantilla.ResetPasswordRequest" %>

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

    <title>Recuperar contraseña</title>

    <script type="text/javascript">
        function MostrarAlerta() {
            document.getElementById("alert").style.display = "block";
        }
    </script>
</head>

<body style="background:#252149;">
    <div class="container">
        <br />

        <form id="form1" runat="server" style="min-width:300px;">
            <div class="login">
                <div align="center">
                    <br /><br />
                    <img alt="Logotipo" src="ImagenesDocbox/logocorielespaceland.png"
                         data-at2x="ImagenesDocbox/logocorielespaceland.png" />
                    <br /><br />

                    <div id="alert" class="alert alert-danger" style="display:none;">
                        <button type="button" class="close" data-dismiss="alert">&times;</button>
                        <asp:Label ID="mensaje" style="white-space: normal;" runat="server" CssClass="label"></asp:Label>
                    </div>

                    <div id="ok" class="alert alert-success" style="display:none;">
                        <button type="button" class="close" data-dismiss="alert">&times;</button>
                        <asp:Label ID="mensajeOk" style="white-space: normal;" runat="server" CssClass="label"></asp:Label>
                    </div>

                    <asp:Panel ID="panel1" runat="server">
                        <h4 style="color:#252149; margin-bottom:15px;">Recuperar contraseña</h4>

                        <div class="inner-addon left-addon">
                            <i class="glyphicon glyphicon-user"></i>
                            <asp:TextBox ID="txtusuario" runat="server" CssClass="form-control" placeholder="Usuario (email)"></asp:TextBox>
                        </div>

                        <br />

                        <center>
                            <asp:LinkButton ID="btnEnviar"
                                runat="server"
                                CssClass="btn btn-login btn-sm"
                                style="background-color:#21B2BA; border-color:#21B2BA; color:#fff;">
                                <i class="glyphicon glyphicon-envelope"></i> Enviar enlace
                            </asp:LinkButton>
                        </center>

                        <br />
                        
                    </asp:Panel>
                    <asp:Panel ID="panelVolver" runat="server" >
                        <br />
                        <center>
                            <asp:HyperLink ID="lnkVolver" runat="server" NavigateUrl="~/RegistroHorario.aspx"
                                style="color:#252149; text-decoration:underline;">
                                Volver al acceso
                            </asp:HyperLink>
                        </center>
                    </asp:Panel>


                </div>
            </div>
        </form>

        <script src="scripts/jquery-2.1.3.min.js" type="text/javascript"></script>
        <script src="scripts/bootstrap.min.js" type="text/javascript"></script>
    </div>
</body>
</html>
