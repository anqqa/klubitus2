<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = -1000 

lisays = request.form("lisays")
alue = lcase(zeroString(request.form("alue")))
selitys = lcase(zeroString(request.form("selitys")))
tyyppi = request.form("tyyppi")
readyDBCon

if lisays = "true" then
	Set RS = Server.CreateObject( "ADODB.Recordset" )
	RS.Open "SELECT * FROM message_areas", Con, 2, 2
	RS.AddNew
	RS("topic") = alue
	RS("desc") = selitys
	RS("type") = tyyppi
	RS("author_id") = session("id")
	RS.Update
	RS.Close
end if
set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
SQL = "SELECT a.id, a.topic, a.desc, a.type, COUNT(m.id) AS viesteja, a.last_message_entered, a.last_message_topic, a.last_message_author FROM message_areas AS a, messages AS m WHERE a.type < 10 AND m.area = a.id GROUP BY a.id, a.topic, a.desc, a.type, a.last_message_entered, a.last_message_topic, a.last_message_author ORDER BY a.type, a.id"
SQL2 = "SELECT a.id, COUNT(m.id) AS aiheita FROM messages AS m, message_areas AS a WHERE m.reply_to = -1 AND m.area = a.id GROUP BY a.id, a.type ORDER BY a.type, a.id"
RS.Open SQL
yht = int(RS.RecordCount)
viesteja = 0
if yht > 0 then 
	alueet = RS.GetRows()
	for i = 0 to yht - 1
		viesteja = viesteja + alueet(4, i)
	next
end if
RS.Close
RS.Open SQL2
yht = int(RS.RecordCount)
aiheita = 0
if yht > 0 then 
	aiheet = RS.GetRows()
	for i = 0 to yht - 1
		aiheita = aiheita + aiheet(1, i)
	next
end if
RS.Close
Con.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>forum</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#002200" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<table cellspacing="2" cellpadding="3" border="0">
<tr>
	<td class="head">klubitus forumit</td>
	<td align="right" colspan="3">
aiheita yhteensä: <%= aiheita %> - viestejä yhteensä: 
<%
response.write viesteja 
if session("admin") = "true" then response.write " · <a href=""forum_alue_lisaa.asp"" class=""sub2"">uus alue</a>"
%>
	</td>
</tr>
<tr bgcolor="#009900">
	<td class="sub1">alue</td>
	<td class="sub1" align="right">aiheita</td>
	<td class="sub1" align="right">viestejä</td>
	<td class="sub1" align="left">uusin</td>
</tr>
<%
if yht > 0 then
	for i = 0 to yht - 1
		if int(alueet(3, i)) = 1 then
			response.write "<tr>" & vbNewline
			response.write "<td><b><a href=""forum_alue.asp?alue=" & alueet(0, i) & """ class=""sub2"">" & alueet(1, i) & "</a></b><br>" & alueet(2, i) & "</td>" & vbNewline
			response.write "<td align=""right"" valign=""top"">" & aiheet(1, i) & "</td>" & vbNewline
			response.write "<td align=""right"" valign=""top"" class=""sub1"">" & alueet(4, i) & "</td>" & vbNewline
			response.write "<td class=""sub2"">" & alueet(5, i) & " - " & alueet(7, i) & "<br>aihe: " & alueet(6, i) & "</td>" & vbNewline
			response.write "</tr>" & vbNewline
		end if
	next
end if
%>
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
