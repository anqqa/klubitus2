<%
if session("admin") <> "true" then response.redirect "virhe.asp"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>forum . alue. lisää</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#002200" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<form action="forum.asp" method="post">
<table cellspacing="5" cellpadding="0" border="0">
<tr>
	<td class="selite">alue:</td>
	<td><input type="text" maxlength="100" size="80" value="" name="alue"></td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<tr>
	<td class="selite" valign="top">selitys:</td>
	<td><textarea cols="80" rows="5" name="selitys"></textarea></td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<tr>
	<td class="selite" valign="top">tyyppi:</td>
	<td>
<select name="tyyppi">
	<option value="1" SELECTED>julkinen</option>
	<option value="2">vain luku</option>
	<option value="3">rekisteröityneille</option>
	<option value="5">privaatti</option>
</select>
	</td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<tr>
	<td align="right" colspan="2">
<input type="hidden" name="lisays" value="true">
<input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok">&nbsp;<a href="javascript:history.go(-1)"><img height="36" alt="antaa olla" src="pics/icon_pyh.gif" width="28" border="0"></a>
	</td>
</tr>
</table>
</form>
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