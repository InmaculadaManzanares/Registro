<%@ Page Language="VB" AutoEventWireup="false" CodeBehind="ResetPassword.aspx.vb" Inherits="NuevaPlantilla.ResetPassword" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />

    <link rel="shortcut icon" type="image/x-icon" href="ImagenesDocbox/favicon.ico" />
    <link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <link href="css/Login.css" rel="stylesheet" type="text/css" />

    <title>Nueva contraseña</title>

    <script type="text/javascript">
        function MostrarAlerta() {
            document.getElementById("alert").style.display = "block";
        }
    </script>

    <style type="text/css">
        /* Forzamos contraste y visibilidad dentro del login */
        .titulo-reset { color: #ffffff !important; font-weight: bold; font-size: 20px; margin: 0; }
        .subtitulo-reset { color: #dddddd !important; margin: 6px 0 14px 0; }
        .reglas-box {
            background: #ffffff;
            border-radius: 6px;
            padding: 12px 14px;
            margin: 10px 0 14px 0;
            text-align: left;
            color: #222;
            font-size: 13px;
        }
        .reglas-box ul { margin: 6px 0 0 18px; }
        .reglas-box li { margin: 2px 0; }
    </style>
</head>

<body style="background:#252149;">
    <div class="container">
        <br />

        <form id="form1" runat="server" style="min-width:300px;">
            <div class="login">
                <div align="center">
                    <br />
                    <br />

                    <img alt="Logotipo" src="ImagenesDocbox/logocorielespaceland.png"
                         data-at2x="ImagenesDocbox/logocorielespaceland.png" style="padding-bottom : 20px" />

                         <br />

                    <!-- TITULO -->
                    <asp:Label ID="labelmensaje" runat="server"  text="Introduce una nueva contraseña para tu acceso" />

                    <!-- ALERTA -->
                    <div id="alert" class="alert alert-danger" style="display:none;">
                        <button type="button" class="close" data-dismiss="alert">&times;</button>
                        <asp:Label ID="mensaje" style="white-space: normal;" runat="server" CssClass="label"></asp:Label>
                    </div>

                    <!-- PANEL RESET -->
                    <asp:Panel ID="panelReset" runat="server">

                        <!-- REGLAS (SIEMPRE VISIBLE) -->
                        <div class="reglas-box">
                            <strong>La contraseña debe cumplir:</strong>
                            <ul>
                                <li>Mínimo <strong>8 caracteres</strong></li>
                                <li>Al menos <strong>una letra mayúscula</strong></li>
                                <li>Al menos <strong>una letra minúscula</strong></li>
                                <li>Al menos <strong>un número</strong></li>
                                <li>Al menos <strong>un carácter especial</strong> (por ejemplo: ! @ # $ %)</li>
                                <li>Sin espacios</li>
                            </ul>
                        </div>

                        <div class="inner-addon left-addon">
                            <i class="glyphicon glyphicon-lock"></i>
                            <asp:TextBox ID="txtNewPass" runat="server" CssClass="form-control"
                                TextMode="Password" placeholder="Nueva contraseña"></asp:TextBox>
                        </div>

                        <div class="inner-addon left-addon" style="margin-top:10px;">
                            <i class="glyphicon glyphicon-lock"></i>
                            <asp:TextBox ID="txtNewPass2" runat="server" CssClass="form-control"
                                TextMode="Password" placeholder="Repetir nueva contraseña"></asp:TextBox>
                        </div>

                        <br />

                        <center>
                            <asp:Button ID="btnSet" runat="server"
                                CssClass="btn btn-login btn-sm"
                                Text="Cambiar contraseña"
                                style="padding-left:30px;padding-right:30px;" />
                        </center>

                        <br />

                        <center>
                            <asp:LinkButton ID="lnkVolver" runat="server"
                                CssClass="btn btn-login btn-sm" CausesValidation="False">
                                <i class="glyphicon glyphicon-home"></i> Volver
                            </asp:LinkButton>
                        </center>

                    </asp:Panel>

                    <!-- PANEL OK -->
                    <asp:Panel ID="panelOk" runat="server" Visible="false">
                        <div class="alert alert-success" style="margin-top:10px;">
                            <asp:Label ID="lblOk" runat="server" style="white-space: normal;"></asp:Label>
                        </div>

                        <center>
                            <asp:LinkButton ID="lnkIrLogin" runat="server"
                                CssClass="btn btn-login btn-sm" CausesValidation="False">
                                <i class="glyphicon glyphicon-log-in"></i> Ir al acceso
                            </asp:LinkButton>
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
