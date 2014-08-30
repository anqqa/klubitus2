<!-- #INCLUDE FILE="dbfuncs.asp" -->
<% 
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.CacheControl="no-cache"
if session("nikki") = "" then response.redirect "virhe.asp"
%>
<html>
<head>
	<title>bileet . muokkaa</title>
	<link rel="stylesheet" href="liila.css" type="text/css">
<%
id = request.querystring("id")

readyDBCon
SQL = "SELECT p.day, p.month, p.year, p.name, p.name_url, p.place, p.place_url, p.city, p.dj, p.hours_from, p.hours_to, p.age, p.price, p.music, p.info, p.added, m.nick FROM parties AS p, members AS m WHERE (p.id = " & id & ") AND (p.added_by = m.id)"
set RS = Con.Execute(SQL)
%>
</head>

<body bgcolor="#000000" background="pics/tausta_liila.jpg" bgproperties="fixed">

<form action="bileet_muokkaus.asp" method="post">
<table cellspacing="0" cellpadding="0" border="0">
<tr>
	<td valign="top" align="left" width="50%">
	<table cellspacing="5" cellpadding="0" width="100%" border="0">
	<tr>
		<td class="selite"><b>p‰iv‰m‰‰r‰:</b></td>
		<td>
<select name="p">
<%
for i = 1 to 31
	if RS("day") = i then
		response.write "<option value='" & i & "' selected>" & i & ".</option>"
	else
		response.write "<option value='" & i & "'>" & i & ".</option>"
	end if
next
%>
</select> 

<select name="kk">
<option value="1" <%if RS("month") = 1 then response.write("selected") %>>tammikuuta</option>
<option value="2" <%if RS("month") = 2 then response.write("selected") %>>helmikuuta</option>
<option value="3" <%if RS("month") = 3 then response.write("selected") %>>maaliskuuta</option>
<option value="4" <%if RS("month") = 4 then response.write("selected") %>>huhtikuuta</option>
<option value="5" <%if RS("month") = 5 then response.write("selected") %>>toukokuuta</option>
<option value="6" <%if RS("month") = 6 then response.write("selected") %>>kes‰kuuta</option>
<option value="7" <%if RS("month") = 7 then response.write("selected") %>>hein‰kuuta</option>
<option value="8" <%if RS("month") = 8 then response.write("selected") %>>elokuuta</option>
<option value="9" <%if RS("month") = 9 then response.write("selected") %>>syyskuuta</option>
<option value="10" <%if RS("month") = 10 then response.write("selected") %>>lokakuuta</option>
<option value="11" <%if RS("month") = 11 then response.write("selected") %>>marraskuuta</option>
<option value="12" <%if RS("month") = 12 then response.write("selected") %>>joulukuuta</option>
</select>

<select name="v">
<option value="2000" <%if RS("year") = 2000 then response.write("selected") %>>2000</option>
<option value="2001" <%if RS("year") = 2001 then response.write("selected") %>>2001</option>
<option value="2002" <%if RS("year") = 2002 then response.write("selected") %>>2002</option>
<option value="2003" <%if RS("year") = 2003 then response.write("selected") %>>2003</option>
</select>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><b>nimi:</b></td>
		<td><input type="text" maxlength="100" size="30" value="<%=RS("name") %>" name="nimi"></td>
	</tr>
	<tr>
		<td class="selite"> - kotisivu:</td>
		<td><input type="text" maxlength="100" size="30" value="<%=unzeroString(RS("name_url")) %>" name="nimi_url"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite" valign="top">paikka:</td>
		<td><input type="text" maxlength="100" size="30" value="<%=unzeroString(RS("place")) %>" name="paikka"></td>
	</tr>
	<tr>
		<td class="selite" valign="top"> - kotisivu:</td>
		<td><input type="text" maxlength="100" size="30" value="<%=unzeroString(RS("place_url")) %>" name="paikka_url"></td>
	</tr>
	<tr>
		<td class="selite">ik‰raja:</td>
		<td>K<input type="text" maxlength="2" size="1" name="ikaraja" value="<%=RS("age") %>"></td>
	</tr>
	<tr>
		<td class="selite">auki:</td>
		<td>
		<table cellpadding="0" cellspacing="0" border="0">
		<tr>
		<td>
<select name="auki1">
<option value="-1" <%if RS("hours_from") = -1 then response.write("selected") %>>--</option>
<%
for i = 23 to 0 step -1
	if RS("hours_from") = i then
		response.write "<option value='" & i & "' selected>" & zero(i) & "</option>"
	else
		response.write "<option value='" & i & "'>" & zero(i) & "</option>"
	end if
next
%>
</select>
		</td>
		<td>&nbsp;-&nbsp;</td>
		<td>
<select name="auki2">
<option value="-1" <%if RS("hours_to") = -1 then response.write("selected") %>>--</option>
<%
for i = 0 to 23
	if RS("hours_to") = i then
		response.write "<option value='" & i & "' selected>" & zero(i) & "</option>"
	else
		response.write "<option value='" & i & "'>" & zero(i) & "</option>"
	end if
next
%>
</select>
		</td>
		</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite" valign="top"><b>kaupunki:</b></td>
		<td>
<select name="kaupunki">
<option value="helsinki" <%if RS("city") = "helsinki" then response.write "selected" %>>helsinki</option>
<option value="j‰rvenp‰‰" <%if RS("city") = "j‰rvenp‰‰" then response.write "selected" %>>j‰rvenp‰‰</option>
<option value="tampere" <%if RS("city") = "tampere" then response.write "selected" %>>tampere</option>
<option value="turku" <%if RS("city") = "turku" then response.write "selected" %>>turku</option>
<option value="espoo" <%if RS("city") = "espoo" then response.write "selected" %>>espoo</option>
<option value="jyv‰skyl‰" <%if RS("city") = "jyv‰skyl‰" then response.write "selected" %>>jyv‰skyl‰</option>
<option value="kotka" <%if RS("city") = "kotka" then response.write "selected" %>>kotka</option>
<option value="kouvola" <%if RS("city") = "kouvola" then response.write "selected" %>>kouvola</option>
<option value="kuopio" <%if RS("city") = "kuopio" then response.write "selected" %>>kuopio</option>
<option value="lahti" <%if RS("city") = "lahti" then response.write "selected" %>>lahti</option>
<option value="lappeenranta" <%if RS("city") = "lappeenranta" then response.write "selected" %>>lappeenranta</option>
<option value="oulu" <%if RS("city") = "oulu" then response.write "selected" %>>oulu</option>
<option value="pori" <%if RS("city") = "pori" then response.write "selected" %>>pori</option>
<option value="vaasa" <%if RS("city") = "vaasa" then response.write "selected" %>>vaasa</option>
<option value="-muu-" <%if RS("city") = "-muu-" then response.write "selected" %>>-muu-</option>
</select>
		</td>
	</tr>
	</table>
	</td>

	<td valign="top" align="left" width="50%">
	<table cellspacing="5" cellpadding="0" width="100%" border="0">
	<tr>
		<td class="selite">dj:</td>
		<td><input type="text" maxlength="250" size="30" name="dj" value="<%=unzeroString(RS("dj")) %>"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite">musiikkityyli:</td>
		<td><input type="text" maxlength="60" size="30" name="musiikki" value="<%=unzeroString(RS("music")) %>"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite">muuta infoa:</td>
		<td><textarea name="muuta" cols="30" rows="3" wrap="virtual"><%=RS("info") %></textarea></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite">normaali hinta:</td>
		<td><input type="text" maxlength="3" size="1" name="hinta" value="<%=RS("price") %>">mk</td>
	</tr>
	<tr>
<%
if session("admin") = "true" then
	response.write "<td>lis‰‰j‰:</td><td>" & RS("nick") & " (" & RS("added") & ")</td>" & vbnewline
else
	response.write "<td colspan=""2"">&nbsp;</td>" & vbnewline
end if
%>
	</tr>
	<tr>
		<td align="right" colspan="2">
<input type="hidden" name="id" value="<%=id %>">
<input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok">&nbsp;<a href="javascript:history.go(-1)"><img height="36" alt="antaa olla" src="pics/icon_pyh.gif" width="28" border="0"></a>
		</td>
	</tr>
	</table>
	</td>
</tr>
</table>
</form>
<%
RS.Close
Con.Close
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