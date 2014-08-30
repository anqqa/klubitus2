<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<%
kuva = request.querystring("kuva")
bileet = request.querystring("bileet")
num = cint(request.querystring("num"))
yht = cint(request.querystring("yht"))
%>
	<title><%=bileet %>&nbsp;<%=num %> / <%=yht %></title>
<link rel="stylesheet" href="sininen.css" type="text/css">
</head>

<body bgcolor="#000000" text="#FFFFFF" background="pics/tausta_sininen.jpg" bgproperties="fixed" leftmargin="0">

<table width="100%" height="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td colspan="2" width="100%" align="center"><img src="photot/<%=kuva & num %>.jpg" border="0"></td>
</tr>
<tr>
	<td align="right" width="50%" valign="bottom">
<% if (num > 1) then %>
<a class="sub1" href="kuva.asp?kuva=<%=kuva %>&bileet=<%=pileet %>&num=<%=(num - 1) %>&yht=<%=yht %>">edellinen</a>&nbsp;
<% end if %>
	</td>
	<td align="left" width="50%" valign="bottom">
<% if (num < yht) then %>
&nbsp;<a class="sub1" href="kuva.asp?kuva=<%=kuva %>&bileet=<%=pileet %>&num=<%=(num + 1) %>&yht=<%=yht %>">seuraava</a>
<% end if %>
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
