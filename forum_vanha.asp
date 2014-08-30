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
	RS.ActiveConnection = Con
	RS.CursorType = adOpenStatic
	RS.LockType = adLockOptimistic
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
SQL = "SELECT id, topic, replies, last_reply, replied, status FROM messages WHERE status < 10 ORDER BY status, id"
RS.Open SQL
yht = int(RS.RecordCount)
viesteja = 0
if yht > 0 then 
	alueet = RS.GetRows()
	for i = 0 to yht - 1
		viesteja = viesteja + alueet(2, i)
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
	<td align="right" colspan="2">
viestejä yhteensä: 
<%
response.write viesteja 
if session("admin") = "true" then response.write " · <a href=""forum_alue_lisaa.asp"" class=""sub2"">uus alue</a>"
%>
	</td>
</tr>
<tr bgcolor="#009900">
	<td class="sub1">alue</td>
	<td class="sub1">viestejä</td>
	<td class="sub1">uusin</td>
</tr>
<%
if yht > 0 then
	for i = 0 to yht - 1
		if int(alueet(5, i)) < 3 or session("nikki") <> "" then
			response.write "<tr>" & vbNewline
			response.write "<td><a href=""forum_alue.asp?alue=" & alueet(0, i) & """ class=""sub2"">" & alueet(1, i) & "</a></td>" & vbNewline
			response.write "<td align=""right"">" & alueet(2, i) & "</td>" & vbNewline
			response.write "<td>" & paiva(alueet(4, i)) & " - " & alueet(3, i) & "</td>" & vbNewline
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
