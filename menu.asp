<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.CacheControl="no-cache"
Response.Buffer = True 

If Session("nikki") = "" Then
	Response.Redirect "login.asp"
End If

readyDBCon
%>

<html>
<head>
	<title>menu</title>
	<link rel="stylesheet" href="menu.css" type="text/css">
	<meta http-equiv="refresh" content="180">
</head>

<body bgcolor="#000000" text="#FFFFFF" background="pics/tausta_sivu.jpg">

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<!-- logout -->
<tr>
	<td valign="top">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td><a href="default.asp?logout=true" target="_top"><img onmouseover="this.src='pics/link_logout2.gif'" onmouseout="this.src='pics/link_logout.gif'" src="pics/link_logout.gif" width="57" height="11" alt="logout" border="0" vspace="10"></a></td>
</tr>
</table>
	</td>
</tr>

<!-- uutiset -->
<tr>
	<td valign="top">
<%
set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
RS.MaxRecords = 5
SQL = "SELECT id, topic, entered FROM messages WHERE status = 11 ORDER BY entered DESC"
RS.Open SQL
yht = int(RS.RecordCount)
if yht > 0 then uutiset = RS.GetRows()
RS.Close
%>
<table cellpadding="0" cellspacing="1" border="0">
<tr>
	<td colspan="2">uusimmat uutiset</td>
</tr>
<%
for i = 0 to 4
	response.write "<tr><td width=""30"" align=""right"">" & day(uutiset(2, i)) & "." & month(uutiset(2, i)) & ".</td>"
	if len(uutiset(1, i)) > 12 then uutiset(1, i) = left(uutiset(1, i), 12) & ".."
	response.write "<td class=""sub1""><nobr><a href=""uutiset.asp"" class=""sub1"" target=""main"">" & uutiset(1, i) & "</a></nobr></td></tr>" & vbNewline
next
%>
	</td>
</tr>
</table>
	</td>
</tr>

<!-- lisätty -->
<tr>
	<td valign="top">
<%
set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
RS.MaxRecords = 5
SQL = "SELECT day, month, name, name_url FROM parties ORDER BY added DESC"
RS.Open SQL
bileet = RS.GetRows()
RS.Close
%>
<br>
<table cellpadding="0" cellspacing="1" border="0">
<tr>
	<td colspan="2">viimeks lisätty</td>
</tr>
<%
for i = 0 to 4
	response.write "<tr><td width=""30"" align=""right"">" & bileet(0, i) & "." & bileet(1, i) & ".</td>"
	if len(bileet(2, i)) > 12 then bileet(2, i) = left(bileet(2, i), 12) & ".."
	if bileet(3, i) = " " then
		response.write "<td class=""sub1""><nobr>" & bileet(2, i) & "</nobr></td></tr>" & vbNewline
	else
		response.write "<td class=""sub1""><nobr><a href=""" & bileet(3, i) & """ class=""sub1"" target=""_blank"">" & bileet(2, i) & "</a></nobr></td></tr>" & vbNewline
	end if
next
%>
	</td>
</tr>
</table>
	</td>
</tr>

<!-- muokattu -->
<tr>
	<td valign="top">
<%
set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
RS.MaxRecords = 5
SQL = "SELECT day, month, name, name_url FROM parties ORDER BY modified DESC"
RS.Open SQL
bileet = RS.GetRows()
RS.Close
Con.Close
%>
<br>
<table cellpadding="0" cellspacing="1" border="0">
<tr>
	<td colspan="2">viimeks muutettu</td>
</tr>
<%
for i = 0 to 4
	response.write "<tr><td width=""30"" align=""right"">" & bileet(0, i) & "." & bileet(1, i) & ".</td>"
	if len(bileet(2, i)) > 12 then bileet(2, i) = left(bileet(2, i), 12) & ".."
	if bileet(3, i) = " " then
		response.write "<td class=""sub1"">" & bileet(2, i) & "</nobr></td></tr>" & vbNewline
	else
		response.write "<td class=""sub1""><nobr><a href=""" & bileet(3, i) & """ class=""sub1"" target=""_blank"">" & bileet(2, i) & "</a></nobr></td></tr>" & vbNewline
	end if
next
%>
	</td>
</tr>
</table>

	</td>
</tr>

<!-- chat -->
<tr>
	<td align="left">
<br>
porukkaa sivuilla: <%= linjoilla(request.servervariables("REMOTE_ADDR")) %><br>
<br>
porukkaa chatissa: <%= application("chat_porukkaa") %>
	</td>
</tr>
<!-- /chat -->
</table>

</body>
</html>


