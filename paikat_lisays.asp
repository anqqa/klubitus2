<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.CacheControl="no-cache"
%>
<html>
<head>
	<title>paikat . lisäys</title>
<link rel="stylesheet" href="liila.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_liila.jpg" bgproperties="fixed">

<%
if session("admin") = "" then response.redirect "virhe.asp"

nimi = lcase(request.form("nimi"))
kotisivu = formatURL(request.form("kotisivu"))
ikaraja = request.form("ikaraja")
auki1 = request.form("auki1")
auki2 = request.form("auki2")
kaupunki = request.form("kaupunki")
osoite = request.form("osoite")
puhelin = request.form("puhelin")
puhelin2 = request.form("puhelin2")
fax = request.form("fax")
' lisäys!
readyDBCon
Set RS = Server.CreateObject( "ADODB.Recordset" )
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
RS.LockType = adLockOptimistic
RS.Open "SELECT * FROM places", Con, 2, 2
RS.AddNew
RS("name") = nimi
RS("homepage") = kotisivu
RS("city") = kaupunki
RS("open_from") = auki1
RS("open_to") = auki2
if ikaraja <> "" then RS("age") = ikaraja
if osoite <> "" then RS("address") = osoite
if puhelin <> "" then RS("phone1") = puhelin
if puhelin2 <> "" then RS("phone2") = puhelin2
if fax <> "" then RS("fax") = fax
RS.Update
RS.Close
Con.Close
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
bueno. lisääminen luonnistu.
<p>
<a href="paikat2.asp" class="navi">takas</a>
	</td>
</tr>
</table>

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
