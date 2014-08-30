<!-- #INCLUDE FILE="dbfuncs.asp" -->
<html>
<head>
	<title>chat . viestit</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
<script language="javascript">
<!--
var ladattu = false;
var skrolli = true;

function muuta_skrolli() {
	if (skrolli) {
		skrolli = false;
	} else {
		skrolli = true;
	}
}

function kirjoita_viesti(viesti) {
	if (ladattu) {
		viesti = viesti.replace(/«/gi, "&lt;");
		viesti = viesti.replace(/»/gi, "&gt;");
		viestit.insertAdjacentHTML("BeforeEnd", viesti);
		if (skrolli) self.scroll(0, 71234);
	}
}

function shake() {
	for (i = 10; i > 0; i--) {
		for (z = 10; z > 0; z--) {
			top.moveBy(0, i); 
			top.moveBy(i, 0); 
			top.moveBy(0, -i); 
			top.moveBy(-i, 0); 
		}  
	} 
} 

//-->
</script>
</head>

<body bgcolor="#000000" background="pics/tausta_vihrea.jpg" bgproperties="fixed" onload="ladattu=true">
<table width="100%" border="0" cellpadding="0" cellspacing="5">
<tr>
	<td>
<div id="viestit">
</div>
	</td>
</tr>
</table>

</body>
</html>