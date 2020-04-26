<!doctype html>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>Carolina Betancourt - Bienestar, Salud y Belleza</title>
	<link rel="icon" href="images/favicon.png" type="image/png" />
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="keywords" content="Carolina Betancourt, Bienestar, Salud, Belleza, Botox, Acido Hialuronico, tratramientos, estetica, faciales, corporales, hilos tensores, limpieza facial, antiembejecimiento" />
	<script type="application/x-javascript"> addEventListener("load", function() { setTimeout(hideURLbar, 0); }, false); function hideURLbar(){ window.scrollTo(0,1); } </script>
	<link href='http://fonts.googleapis.com/css?family=Oswald:400,300,700' rel='stylesheet' type='text/css'>
	<link href='http://fonts.googleapis.com/css?family=Niconne' rel='stylesheet' type='text/css'>
 <div id="fb-root"></div>
	<script>(function(d, s, id) {
	  var js, fjs = d.getElementsByTagName(s)[0];
	  if (d.getElementById(id)) return;
	  js = d.createElement(s); js.id = id;
	  js.src = "//connect.facebook.net/es_LA/sdk.js#xfbml=1&version=v2.4";
	  fjs.parentNode.insertBefore(js, fjs);
	}(document, 'script', 'facebook-jssdk'));
	</script>
</head>
<body>
<%
	If Request.ServerVariables("REQUEST_METHOD")="POST" Then
		'ENVIO DEL FORMULARIO DE CONTACTO
		sch = "http://schemas.microsoft.com/cdo/configuration/"
		Set cdoConfig = CreateObject("CDO.Configuration")
		With cdoConfig.Fields
			.Item(sch & "sendusing") = 2
			'.Item(sch & "smtpserverpickupdirectory") = "C:\inetpub\mailroot\pickup" 
			.Item(sch & "smtpserver") = "localhost"
			.Item(sch & "smtpserverport") = 25
			.Item(sch & "smtpconnectiontimeout") = 40
			.Item(sch & "smtpauthenticate") = 1
			.Item(sch & "sendusername") = "carolinabetancourt_88@hotmail.com"
			.Item(sch & "sendpassword") = "Carito2004"
			.update
		End With

		Set MailObject = Server.CreateObject("CDO.Message")
		Set MailObject.Configuration = cdoConfig
		'MailObject.BodyFormat = 0
		'MailObject.mailformat = 0
		MailObject.From	= "web@carolinabetancourt.com.co"
		MailObject.To	= "carolinabetancourt_88@hotmail.com"
		MailObject.Subject = "Contacto DESDE PAGINA WEB"
		nombre = Request.Form("nombre")
		Cuerpo = "Nombre: " & Request.Form("nombre") & "<br>"
		Cuerpo = Cuerpo & "Email: " & Request.Form("email") & "<br>"
		Cuerpo = Cuerpo & "Celular: " & Request.Form("celular") & "<br>"
		Cuerpo = Cuerpo & "Mensaje: " & Request.Form("mensaje") & "<br>"
		MailObject.HTMLBody = Cuerpo
		MailObject.Send
   if Err ><0 then 
        rspta = "Error, no se ha podido completar la operacion" 
    else 
        rspta = "Gracias por escribirnos: " & nombre & ", el mensaje se ha enviado correctamente." 
    end if 
    
	Set MailObject = Nothing
	Set cdoConfig = Nothing
End If
%>

<footer id="contact" class="footer_wrapper">
    <div class="container" align="center"  >
        <div class="footer_bottom">
            <div class="fb-like" data-href="https://www.facebook.com/Carolina-Betancourt-Bienestar-Salud-Y-Belleza-2071904196414111" data-layout="button_count" data-action="like" data-show-faces="false" data-share="false"></div>
        </div>
    </div>
</footer>
  
 <div class="container" align="center" >
    <div class="container">
        <br>
        <h4><% response.write rspta %></h4>
        <br>
        <a href="http://www.carolinabetancourt.com.co"><h2>Volver</h2></a>
        <br>
    </div>
  
 </div>


    <section class="container"  style="background: #E0D4BE; MARGIN-TOP: 15PX; MARGIN-BOTTOM:15PX; color:#902B05 ">
		<div class="container">
			<div class="col-md-4 text-left" style="MARGIN-TOP: 5PX;">
                <p>Dise�o 2020 por: <a href="http://www.sabet.com.co">SABET Ingenieros</a></p>
			</div>
            <div class="col-md-4 text-center" style="MARGIN-TOP: 25PX;">
                <div class="fb-like" data-href="https://www.facebook.com/Carolina-Betancourt-Bienestar-Salud-Y-Belleza-2071904196414111" data-layout="button_count" data-action="like" data-show-faces="false" data-share="false"></div>
			</div>
			<div class="col-md-4 text-right">
                <p  style="MARGIN-TOP: 10PX;">LL�menos ahora!:</p>
                <h3 style="MARGIN-TOP: 10PX;">311-200-1707</h3>
			</div>
		</div>
	</section>


<script type="text/javascript" src="js/jquery-1.11.0.min.js"></script>
<script type="text/javascript" src="js/bootstrap.min.js"></script>
<script type="text/javascript" src="js/jquery-scrolltofixed.js"></script>
<script type="text/javascript" src="js/jquery.nav.js"></script> 
<script type="text/javascript" src="js/jquery.easing.1.3.js"></script>
<script type="text/javascript" src="js/jquery.isotope.js"></script>
<script type="text/javascript" src="js/wow.js"></script> 
<script type="text/javascript" src="js/custom.js"></script>

</body>
</html>

