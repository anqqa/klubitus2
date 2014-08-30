<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.CacheControl="no-cache"
Response.Buffer = True 
nimi = Request.Form("nimi")
salasana = Request.Form("salasana")

If Session("nikki") <> "" Then
	Response.Redirect "menu.asp"
End If
%>

<html>
<head>
	<title>login</title>
<link rel="stylesheet" href="login.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_sivu.jpg">

<%
If salasana = "" Then
	ShowLogin
Else
	CheckLogin
End If

Sub CheckLogin
	readyDBCon
	SQL = "SELECT id, password, email, config FROM members WHERE nick ='" & nimi & "' AND password ='" & salasana & "'"
	set RS = Con.Execute(SQL)
	if RS.BOF and RS.EOF then
		ShowLogin
	else
    	Session("nikki") = nimi
    	Session("id") = int(RS("id"))
		Session("email") = RS("email")
		session("config") = RS("config")
		if nimi = "anqqa" then session("admin") = "true"
%>
<script language="javascript">
<!--
	top.document.location = "default.asp";
//-->
</script>
<%
	End If
End Sub

Sub ShowLogin
%>
<form action="login2.asp" method="post">
<table width="100%" height="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td valign="top">

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td><img src="pics/text_nimi.gif" width="98" height="11" alt="nimi:" border="0" vspace="5"></td>
</tr>
<tr>
	<td><input type="text" name="nimi" size="10" maxlength="20"></td>
</tr>
<tr>
	<td><img src="pics/text_salasana.gif" width="98" height="11" alt="salasana:" border="0" vspace="5"></td>
</tr>
<tr>
	<td><input type="password" name="salasana" size="10" maxlength="20"></td>
</tr>
<tr>
	<td><input type="image" name="login" src="pics/link_login.gif" alt="login" width="98" height="11" vspace="10" border="0" onmouseover="this.src='pics/link_login2.gif'" onmouseout="this.src='pics/link_login.gif'"></td>
</tr>
</table>
</form>

	</td>
</tr>
</table>
<%
end sub
%>

</body>
</html>


