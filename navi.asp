<!doctype html public "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>navigaatio</title>
<script language="javascript">
<!--
if (document.images) {
	kuva1a = new Image(); kuva1a.src = "pics/link_bileet.gif";
	kuva1b = new Image(); kuva1b.src = "pics/link_bileet2.gif";
	kuva2a = new Image(); kuva2a.src = "pics/link_uutiset.gif";
	kuva2b = new Image(); kuva2b.src = "pics/link_uutiset2.gif";
	kuva3a = new Image(); kuva3a.src = "pics/link_kuvat.gif";
	kuva3b = new Image(); kuva3b.src = "pics/link_kuvat2.gif";
	kuva4a = new Image(); kuva4a.src = "pics/link_palaute.gif";
	kuva4b = new Image(); kuva4b.src = "pics/link_palaute2.gif";
	kuva5a = new Image(); kuva5a.src = "pics/link_paikat.gif";
	kuva5b = new Image(); kuva5b.src = "pics/link_paikat2.gif";
	kuva6a = new Image(); kuva6a.src = "pics/link_forum.gif";
	kuva6b = new Image(); kuva6b.src = "pics/link_forum2.gif";
	kuva7a = new Image(); kuva7a.src = "pics/link_radio.gif";
	kuva7b = new Image(); kuva7b.src = "pics/link_radio2.gif";
	kuva8a = new Image(); kuva8a.src = "pics/link_rekisteri.gif";
	kuva8b = new Image(); kuva8b.src = "pics/link_rekisteri2.gif";
	kuva9a = new Image(); kuva9a.src = "pics/link_liikkeet.gif";
	kuva9b = new Image(); kuva9b.src = "pics/link_liikkeet2.gif";
	kuva10a = new Image(); kuva10a.src = "pics/link_chat.gif";
	kuva10b = new Image(); kuva10b.src = "pics/link_chat2.gif";
	kuva11a = new Image(); kuva11a.src = "pics/link_top10.gif";
	kuva11b = new Image(); kuva11b.src = "pics/link_top102.gif";
	kuva12a = new Image(); kuva12a.src = "pics/link_asetukset.gif";
	kuva12b = new Image(); kuva12b.src = "pics/link_asetukset2.gif";
}

var vanha = 0;
function kuva(num) {
	if (document.images) { 
		if (num == 0) {
			eval("document.images['i" + vanha + "'].src = kuva" + vanha + "a.src");
		} else {
			eval("document.images['i" + num + "'].src = kuva" + num + "b.src");
			vanha = num;
		}
	}
}
//-->
</script>
<base target="main">
</head>

<body bgcolor="#000000" text="#FFFFFF">

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td width="150" align="left" valign="middle" rowspan="3"><img src="pics/klubitus_logo.gif" width="71" height="79" alt="K L U B I T U S" border="0" hspace="30"></td>
	<td height="30" align="left" valign="middle">
<a href="bileet.asp" onmouseover="kuva(1); return true" onmouseout="kuva(0)"><img src="pics/link_bileet.gif" width="60" height="11" alt="bileet" border="0" name="i1"></a>
	</td>
	<td align="left" valign="middle">
<a href="uutiset.asp" onmouseover="kuva(2); return true" onmouseout="kuva(0)"><img src="pics/link_uutiset.gif" width="60" height="11" alt="uutiset" border="0" name="i2"></a>
	</td>
	<td align="left" valign="middle">
<a href="kuvat.asp" onmouseover="kuva(3); return true" onmouseout="kuva(0)"><img src="pics/link_kuvat.gif" width="60" height="11" alt="kuvat" border="0" name="i3"></a>
	</td>
	<td align="left" valign="middle">
<a href="palaute.asp" onmouseover="kuva(4); return true" onmouseout="kuva(0)"><img src="pics/link_palaute.gif" width="90" height="11" alt="palaute" border="0" name="i4"></a>
	</td>
	<td width="350" rowspan="3"><OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,0,0" ID=logo WIDTH=350 HEIGHT=90>
<%
randomize()
logo = int(rnd() * 23 + 1)
%>
<PARAM NAME=movie VALUE="flash/logo<%=logo %>.swf"><PARAM NAME=quality VALUE=high><PARAM NAME=bgcolor VALUE=#000000>
<EMBED src="flash/logo<%=logo %>.swf" quality=high bgcolor=#000000 swLiveConnect=FALSE WIDTH=350 HEIGHT=90 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash">
</EMBED>
</OBJECT>
</td>
</tr>
<tr>
	<td height="30" align="left" valign="middle">
<a href="paikat2.asp" onmouseover="kuva(5); return true" onmouseout="kuva(0)"><img src="pics/link_paikat.gif" width="60" height="11" alt="paikat" border="0" name="i5"></a>
	</td>
	<td align="left" valign="middle">
<a href="forum.asp" onmouseover="kuva(6); return true" onmouseout="kuva(0)"><img src="pics/link_forum.gif" width="60" height="11" alt="forum" border="0" name="i6"></a>
	</td>
	<td align="left" valign="middle">
<a href="radio.asp" onmouseover="kuva(7); return true" onmouseout="kuva(0)"><img src="pics/link_radio.gif" width="60" height="11" alt="radio" border="0" name="i7"></a>
	</td>
	<td>
<% If session("admin") = "true" Then %>
<a href="rekisteri.asp" onmouseover="kuva(8); return true" onmouseout="kuva(0)"><img src="pics/link_rekisteri.gif" width="90" height="11" alt="rekisteri" border="0" name="i8"></a>
<% end if %>
	</td>
</tr>
<tr>
	<td height="30" align="left" valign="middle">
<a href="liikkeet.asp" onmouseover="kuva(9); return true" onmouseout="kuva(0)"><img src="pics/link_liikkeet.gif" width="60" height="11" alt="liikkeet" border="0" name="i9"></a>
	</td>
	<td>
<a href="chat.asp?uusi=true" onmouseover="kuva(10); return true" onmouseout="kuva(0)"><img src="pics/link_chat.gif" width="60" height="11" alt="chat" border="0" name="i10"></a>
	</td>
	<td>
<a href="top10.asp" onmouseover="kuva(11); return true" onmouseout="kuva(0)"><img src="pics/link_top10.gif" width="60" height="11" alt="top 10" border="0" name="i11"></a>
	</td>
	<td align="left" valign="middle">
<% If Session("nikki") <> "" Then %>
<a href="asetukset.asp" onmouseover="kuva(12); return true" onmouseout="kuva(0)"><img src="pics/link_asetukset.gif" width="90" height="11" alt="asetukset" border="0" name="i12"></a>
<% end if %>
	</td>
</tr>
</table>

<%
'<!-- Start of TheCounter.com Code -->
'<script type="text/javascript" language="javascript">
's="na";c="na";j="na";f=""+escape(document.referrer)
'</script>
'<script type="text/javascript" language="javascript1.2">
's=screen.width;v=navigator.appName
'if (v != "Netscape") {c=screen.colorDepth}
'else {c=screen.pixelDepth}
'j=navigator.javaEnabled()
'</script>
'<script type="text/javascript" language="javascript">
'function pr(n) {document.write(n,"\n");}
'NS2Ch=0
'if (navigator.appName == "Netscape" &&
'navigator.appVersion.charAt(0) == "2") {NS2Ch=1}
'if (NS2Ch == 0) {
'r="&size="+s+"&colors="+c+"&referer="+f+"&java="+j+""
'pr("<A HREF=\"http://www.TheCounter.com\" TARGET=\"_top\"><IMG"+
'" BORDER=0 SRC=\"http://c2.thecounter.com/id=1435694"+r+"\"><\/A>")}
'</script>
'<noscript><a href="http://www.TheCounter.com" target="_top"><img
'src="http://c2.thecounter.com/id=1435694" alt="TC" border=0></a>
'</noscript>
'<!-- End of TheCounter.com Code -->
%>
</body>
</html>
