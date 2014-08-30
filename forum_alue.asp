<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = -1000 

alue = request.querystring("alue")
readyDBCon

set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
SQL = "SELECT m.id, topic, author, entered, replies, last_reply, replied FROM messages AS m WHERE m.status = 10 AND m.reply_to = -1 AND m.area = " & alue & " ORDER BY replied DESC"
RS.Open SQL
yht = int(RS.RecordCount)
viesteja = 0
if yht > 0 then 
	viestit = RS.GetRows()
	for v = 0 to yht - 1
		viesteja = viesteja + viestit(4, v) + 1
	next
end if
RS.Close
SQL = "SELECT id, topic, type FROM message_areas WHERE type < 10 ORDER BY type, id"
RS.Open SQL
alueita = int(RS.RecordCount)
if alueita > 0 then alueet = RS.GetRows()
RS.Close
Con.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>forum . alue</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#002200" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<table cellspacing="2" cellpadding="3" border="0">
<tr>
	<td class="head">
<% 
for a = 0 to alueita - 1
	if int(alue) = int(alueet(0, a)) then response.write alueet(1, a)
next
%>
	</td>
<form name="alue">
	<td>
<select name="pika" onChange="window.location.href=this.options[this.selectedIndex].value">
<option value="forum.asp">f o r u m i t</option>
<%
for a = 0 to alueita - 1
	if int(alueet(2, a)) = 1 then
		if int(alue) = int(alueet(0, a)) then
			response.write "<option value=""forum_alue.asp?alue=" & alueet(0, a) & """ selected>" & alueet(1, a) & "</option>" & vbNewline
		else
			response.write "<option value=""forum_alue.asp?alue=" & alueet(0, a) & """>" & alueet(1, a) & "</option>" & vbNewline
		end if
	end if
next
%>
</select>
	</td>
</form>
	<td colspan="2" align="right">
viestejä yhteensä: 
<%
response.write viesteja 
%>
 · <a href="forum_viesti_lisaa.asp?alue=<%=alue %>" class="sub2">uus aihe</a>
	</td>
</tr>
<tr bgcolor="#009900">
	<td class="sub1">aihe</td>
	<td class="sub1" align="center">kirjoittaja</td>
	<td class="sub1">vast.</td>
	<td class="sub1">uusin</td>
</tr>
<%
if yht > 0 then
	for i = 0 to yht - 1
		response.write "<tr>" & vbNewline
		response.write "<td><a href=""forum_viesti_lue.asp?alue=" & alue & "&viesti=" & viestit(0, i) & """ class=""sub2"">" & viestit(1, i) & "</a></td>" & vbNewline
		response.write "<td align=""center"">" & viestit(2, i) & "</td>" & vbNewline
		response.write "<td align=""right"" class=""sub1"">" & viestit(4, i) & "</td>" & vbNewline
'		if int(viestit(4, i)) then
'			response.write "<td>" & paiva(viestit(3, i)) & " - " & viestit(2, i) & "</td>" & vbNewline
'		else
			response.write "<td class=""sub2"">" & viestit(6, i) & " - " & viestit(5, i) & "</td>" & vbNewline
'		end if
		response.write "</tr>" & vbNewline
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
