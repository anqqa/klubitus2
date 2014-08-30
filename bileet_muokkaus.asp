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
	<title>bileet . muokkaus</title>
<link rel="stylesheet" href="liila.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_liila.jpg" bgproperties="fixed">

<%
if session("nikki") = "" then response.redirect "virhe.asp"

id = request.form("id")
p = request.form("p")
kk = request.form("kk")
v = request.form("v")
nimi = lcase(request.form("nimi"))
nimi_url = formatURL(request.form("nimi_url"))
paikka = lcase(zeroString(request.form("paikka")))
paikka_url = formatURL(request.form("paikka_url"))
ikaraja = request.form("ikaraja")
auki1 = request.form("auki1")
auki2 = request.form("auki2")
kaupunki = request.form("kaupunki")
dj = lcase(zeroString(request.form("dj")))
musiikki = lcase(zeroString(request.form("musiikki")))
muuta = lcase(zeroString(request.form("muuta")))
hinta = request.form("hinta")

' puuttuuko kohtia?
dim puuttuu
if nimi = "" then puuttuu = "nimi "
if kaupunki = "" then puuttuu = puuttuu + "kaupunki "
if puuttuu <> "" then
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
<b>erhe!</b><br>et täyttäny kaikkia vaadittuja kohtia: <b><%=puuttuu %></b>
<p>
<a href="javascript:history.go(-1)" class="navi">takas</a>
	</td>
</tr>
</table>
<%
else
' muokkaus!
	readyDBCon
	if ikaraja = "-8" then
		Con.Execute "DELETE * FROM parties WHERE id = " & id
	else
		Set RS = Server.CreateObject( "ADODB.Recordset" )
		RS.ActiveConnection = Con
		RS.CursorType = adOpenStatic
		RS.LockType = adLockOptimistic
		RS.Open "SELECT * FROM parties WHERE id = " & id, Con, 2, 2
		RS("day") = p
		RS("month") = kk
		RS("year") = v
		RS("name") = nimi
		RS("name_url") = nimi_url
		RS("place") = paikka
		RS("place_url") = paikka_url
		RS("city") = kaupunki
		RS("dj") = dj
		RS("hours_from") = auki1
		RS("hours_to") = auki2
		if ikaraja <> "" then RS("age") = ikaraja else RS("age") = 0
		if hinta <> "" then RS("price") = hinta else RS("price") = -1
		RS("music") = musiikki
		RS("info") = muuta
		RS("modified_by") = session("id")
		RS("modified") = Now()
		RS("mods") = int(RS("mods")) + 1
		RS.Update
		RS.Close
	end if
	Con.Execute "UPDATE members SET mods = mods+1 WHERE id = " & session("id")
	Con.Close
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
bueno. muokkaus luonnistu.
<p>
<a href="bileet.asp?kk=<%=kk %>&v=<%=v %>" class="navi">takas</a>
	</td>
</tr>
</table>
<%
end if
%>
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
