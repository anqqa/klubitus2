<html>
<head>
    <title>palaute</title>
	<link rel="stylesheet" href="punainen.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_punainen.jpg" bgproperties="fixed">

<%
if request("nimi") <> "" then
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.FromName = Request("nimi")
	Mailer.FromAddress = Request("email")
	Mailer.RemoteHost = "smtpgw.activeisp.com"
	Mailer.AddRecipient "anqqa", "webmaster@anqqa.com"
	Mailer.Subject = Request("aihe")
	Mailer.BodyText = Request("palaute")
	if Mailer.SendMail then %>
<table cellpadding="0" cellspacing="5" border="0">
<tr>
	<td>
kiitoksia palautteesta, palaillaan astialle asap.
<p>
<a href="javascript:history.go(-2)">takas</a>
	</td>
</tr>
</table>
<%
	else
		Response.Write "reisille meni. syy: " & Mailer.Response
	end if
else
%>
<form action="palaute.asp" method="POST">
<table cellpadding="0" cellspacing="5" border="0">
<tr>
	<td class="selite">kuka:</td>
	<td><input type="Text" name="nimi" size="25" maxlength="50" value="<%=session("nikki") %>"></td>
</tr>
<tr>
	<td class="selite">e-mail:</td>
	<td><input type="Text" name="email" size="25" maxlength="50" value="<%=session("email") %>"></td>
</tr>
<tr>
	<td class="selite">aihe:</td>
	<td><input type="Text" name="aihe" size="25" maxlength="50"></td>
</tr>
<tr>
	<td class="selite" valign="top">palaute:</td>
	<td><textarea name="palaute" cols="50" rows="10" wrap="virtual"></textarea></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
<input style="background:" type="Image" name="" src="pics/icon_ok.gif" border="0" width="28" height="35" alt="oujeah!"> <a 	href="javascript:history.go(-1)"><img src="pics/icon_pyh.gif" width=28 height=36 border=0 alt="antaa olla"></a>
	</td>
</tr>
</table>
</form>
<%end if %>
<div align="right">
<!--WEBBOT bot="HTMLMarkup" startspan ALT="Site Meter" -->
<script type="text/javascript" language="JavaScript">var site="sm3klubitus"</script>
<script type="text/javascript" language="JavaScript1.2" src="http://sm3.sitemeter.com/js/counter.js?site=sm3klubitus">
</script>
<noscript>
<a href="http://sm3.sitemeter.com/stats.asp?site=sm3klubitus" target="_top">
<img src="http://sm3.sitemeter.com/meter.asp?site=sm3klubitus" alt="Site Meter" border=0></a>
</noscript>
<!-- Copyright (c)2000 Site Meter -->
<!--WEBBOT bot="HTMLMarkup" Endspan -->
</div>
</body>
</html>