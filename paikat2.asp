<% if not session("index") = true then response.redirect "default.asp?sivu=liikkeet.asp" %>
<!-- #INCLUDE FILE="dbfuncs.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>paikat</title>
	<link rel="stylesheet" href="liila.css" type="text/css">
	<base target="_new">
<%
readyDBCon
set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
SQL = "SELECT * FROM places ORDER BY city, name"
RS.Open SQL
yht = int(RS.RecordCount)
if yht > 0 then paikat = RS.GetRows() ' kenttä, numero
RS.Close
Con.Close
%>
</head>

<body bgcolor="#000000" background="pics/tausta_liila.jpg" bgproperties="fixed">

<table width="100%" cellpadding="0" cellspacing="5" border="0">
<% if session("admin") = "true" then response.write "<tr><td colspan=""4"" align=""right""><a href=""paikat_lisaa.asp"" target=""main"" class=""navi"">uus paikka</a></td></tr>" & vbnewline
response.write "<tr>" & vbnewline

kaupunki = ""
numero = 1
for i = 0 to yht - 1
	if paikat(8, i) <> kaupunki then
		kaupunki = paikat(8, i)
		if numero > 1 then response.write "</tr>" & vbnewline & "<tr>" & vbnewline
		response.write "<td class=""head"" colspan=""4"">" & paikat(8, i) & "<br>&nbsp;</td>" & vbnewline
		response.write "</tr>" & vbnewline & "<tr>" & vbnewline
		numero = 1
	end if
	response.write "<td width=""25%"" class=""sub1"" valign=""top"">"
	if len(paikat(2, i)) > 10 then 
		response.write "<a href=""" & paikat(2, i) & """ class=""sub1"">" & paikat(1, i) & "</a><br>" & vbnewline
	else
		response.write paikat(1, i) & "<br>" & vbnewline
	end if
	response.write "<div class=""sub2"">"
	if paikat(4, i) <> "" then response.write paikat(4, i) & "<br>"
	if paikat(5, i) <> "" then response.write "puh. " & paikat(5, i) & "<br>"
	if paikat(6, i) <> "" then response.write "puh. " & paikat(6, i) & "<br>"
	if paikat(7, i) <> "" then response.write "fax " & paikat(7, i) & "<br>"
	if paikat(9, i) > 0 then response.write "auki: " & zero(paikat(9, i)) & " - " & zero(paikat(10, i)) & "<br>"
	if paikat(11, i) > 0 then response.write "ikäraja: " & paikat(11, i)
	response.write "</div></td>" & vbnewline
	if numero = 4 then
		response.write "</tr><tr><td colspan=""4"">&nbsp;</td></tr><tr>" & vbnewline
		numero = 1
	else
		numero = numero + 1
	end if
next
%>
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
