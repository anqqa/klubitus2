<!-- #INCLUDE FILE="dbfuncs.asp" -->
<html>
<head>
	<title>chat . viestit</title>
	<meta http-equiv="refresh" content="4">
</head>
<%
if session("chat_nikki") <> "" then

application.lock
viestit = application("chat_viestit")
aika = application("chat_aika")
ip = application("chat_ip")
nikki = application("chat_nikki")
selain = application("chat_selain")
porukka = application("chat_porukka")
porukka_ip = application("chat_porukka_ip")
porukka_aika = application("chat_porukka_aika")
porukka_selain = application("chat_porukka_selain")
alias = application("chat_alias")
aliasto = application("chat_aliasto")
uusin = session("uusin")
nimet_vanha = session("chat_nimet")
block = application("chat_block")
application.unlock
nimet = ""

'onko paikalla?
poistui = ""
loyty = false
for i = 0 to 49
	if not loyty and porukka(i) = session("chat_nikki") then
		loyty = true
		porukka_aika(i) = now()
		porukka_selain(i) = session("selain")
	end if
next
if not loyty then
	for i = 0 to 49
		if not loyty and porukka(i) = "" then
			porukka_ip(i) = session("ip")
			porukka(i) = session("chat_nikki")
			porukka_aika(i) = now()
			porukka_selain(i) = session("selain")
			loyty = true
		end if
	next
end if
for i = 0 to 49
	if porukka(i) <> "" then
		'15s poissa päivityksistä -> lähteny
		if datediff("s", porukka_aika(i), now()) > 20 then
			for j = 0 to 48
				viestit(j) = viestit(j + 1)
				aika(j) = aika(j + 1)
				ip(j) = ip(j + 1)
				nikki(j) = nikki(j + 1)
				selain(j) = selain(j + 1)
			next
			aika(49) = now()
			ip(49) = porukka_ip(i)
			nikki(49) = porukka(i)
			selain(49) = porukka_selain(i)
			viestit(49) = "/me meni.."
			porukka(i) = ""
			porukka_aika(i) = ""
			porukka_ip(i) = ""
			porukka_selain(i) = ""
		end if
	end if
	'löyty, lisätään listaan
	if porukka(i) <> "" then
		if nimet <> "" then nimet = nimet & ", "
		nimet = nimet & porukka(i)
	end if
next
application.lock
application("chat_viestit") = viestit
application("chat_aika") = aika
application("chat_ip") = ip
application("chat_nikki") = nikki
application("chat_selain") = selain
application("chat_porukka") = porukka
application("chat_porukka_ip") = porukka_ip
application("chat_porukka_aika") = porukka_aika
application("chat_porukka_selain") = porukka_selain
application.unlock
extra = ""
if aika(49) <> uusin then
	viesti = ""
	for i = (49 - session("riveja")) to 49
		if aika(i) <> "" then 
			if datediff("s", uusin, aika(i)) > 0 then
				blocked = false
				for b = 0 to 19
					if block(b) = lcase(session("chat_nikki")) then blocked = true
				next
				uus_viesti = chat_parser(i, aika, nikki, viestit, ip, alias, aliasto, selain, block)
				if session("chat_seeall") = "true" or left(uus_viesti, 1) <> "*" or (left(uus_viesti, 1) = "*" and (nikki(i) = session("chat_nikki") or instr(lcase(uus_viesti), "(<b>" & lcase(session("chat_nikki")) & "</b>)") > 0)) then 
					viesti = viesti & uus_viesti
				end if
				if instr(lcase(uus_viesti), " <i>zzzap " & lcase(session("chat_nikki")) & "</i>") > 0 then extra = "parent.viestit.shake();"
'				if instr(lcase(uus_viesti), " <i>kvaak " & lcase(session("chat_nikki")) & "</i>") > 0 then extra = "parent.viestit.duckon();"
				if poistui <> "" then viesti = viesti & poistui
			end if
		end if
	next
	for a = 0 to 19
		if alias(a) <> "" then viesti = replace(viesti, alias(a), aliasto(a))
	next
	viesti = replace(viesti, "[img]", "<img align=absmiddle src=")
	viesti = replace(viesti, "[/img]", ">")
	viesti = replace(viesti, "[IMG]", "<img align=absmiddle src=")
	viesti = replace(viesti, "[/IMG]", ">")
	viesti = replace(viesti, "[link]", "<a href=")
	viesti = replace(viesti, "[/link]", " target=_new>linkki</a>")
	viesti = replace(viesti, "http://", "http")
	viesti = replace(viesti, ":/", "<img align=absmiddle src=smiley/indiff.gif>")
	viesti = replace(viesti, "http", "http://")
	viesti = replace(viesti, ":x-", "<img align=absmiddle src=smiley/AR15firing.gif>")
	viesti = replace(viesti, ":X-", "<img align=absmiddle src=smiley/AR15firing.gif>")
	viesti = replace(viesti, "B)", "<img align=absmiddle src=smiley/cool.gif>")
	viesti = replace(viesti, ";)D", "<img align=absmiddle src=smiley/deal.gif>")
	viesti = replace(viesti, "8(", "<img align=absmiddle src=smiley/puppy_dog_eyes.gif>")
	viesti = replace(viesti, ";(", "<img align=absmiddle src=smiley/sad2.gif>")
	viesti = replace(viesti, ":D", "<img align=absmiddle src=smiley/sweat.gif>")
	viesti = replace(viesti, "=D", "<img align=absmiddle src=smiley/lol.gif>")
	viesti = replace(viesti, "d:)", "<img align=absmiddle src=smiley/ymca.gif>")
	viesti = replace(viesti, "q:)", "<img align=absmiddle src=smiley/ymca.gif>")
	viesti = replace(viesti, ":o)", "<img align=absmiddle src=smiley/clown.gif>")
	viesti = replace(viesti, ":I", "<img align=absmiddle src=smiley/indiff.gif>")
	viesti = replace(viesti, ":666", "<img align=absmiddle src=smiley/devil.gif>")
	viesti = replace(viesti, "X(", "<img align=absmiddle src=smiley/dead.gif>")
	viesti = replace(viesti, "x(", "<img align=absmiddle src=smiley/dead.gif>")
	viesti = replace(viesti, ":B", "<img align=absmiddle src=smiley/bucked.gif>")
	viesti = replace(viesti, ":ss", "§ss")
	viesti = replace(viesti, ":s", "<img align=absmiddle src=smiley/badteeth.gif>")
	viesti = replace(viesti, ":S", "<img align=absmiddle src=smiley/badteeth.gif>")
	viesti = replace(viesti, "§ss", ":ss")
	viesti = replace(viesti, ":x)", "<img align=absmiddle src=smiley/nibble6.gif>")
	viesti = replace(viesti, "x)", "<img align=absmiddle src=smiley/loveeyes.gif>")
	viesti = replace(viesti, "X)", "<img align=absmiddle src=smiley/loveeyes.gif>")
	viesti = replace(viesti, "o:)", "<img align=absmiddle src=smiley/mpangel.gif>")
	viesti = replace(viesti, "O:)", "<img align=absmiddle src=smiley/mpangel.gif>")
	viesti = replace(viesti, ":oo", "<img align=absmiddle src=smiley/fruit.gif>")
	viesti = replace(viesti, ":Q", "<img align=absmiddle src=smiley/rasta.gif>")
	viesti = replace(viesti, ":)B", "<img align=absmiddle src=smiley/moon.gif>")
	viesti = replace(viesti, "x:(", "<img align=absmiddle src=smiley/newburn.gif>")
	viesti = replace(viesti, ":.(", "<img align=absmiddle src=smiley/mecry.gif>")
	viesti = replace(viesti, ":(", "<img align=absmiddle src=smiley/sad.gif>")
	viesti = replace(viesti, ":x", "<img align=absmiddle src=smiley/smiley_kiss.gif>")
	viesti = replace(viesti, ":)", "<img align=absmiddle src=smiley/smile.gif>")
	viesti = replace(viesti, "s:o", "<img align=absmiddle src=smiley/gorgeous.gif>")
	viesti = replace(viesti, ":o.", "<img align=absmiddle src=smiley/drooling3.gif>")
	viesti = replace(viesti, ":pb", "<img align=absmiddle src=smiley/readpb.gif>")
	viesti = replace(viesti, ":o", "<img align=absmiddle src=smiley/oogle.gif>")
	viesti = replace(viesti, ":O", "<img align=absmiddle src=smiley/oogle.gif>")
	viesti = replace(viesti, ":P", "<img align=absmiddle src=smiley/buck.gif>")
	viesti = replace(viesti, ":p", "<img align=absmiddle src=smiley/buck.gif>")
	viesti = replace(viesti, ";P", "<img align=absmiddle src=smiley/hehehmn.gif>")
'	viesti = replace(viesti, ";p", "<img align=absmiddle src=smiley/hehehmn.gif>")
	viesti = replace(viesti, ";D", "<img align=absmiddle src=smiley/evil_laughterpurple.gif>")
	viesti = replace(viesti, "8o", "<img align=absmiddle src=smiley/shocked.gif>")
	viesti = replace(viesti, "=)", "<img align=absmiddle src=smiley/tongue.gif>")
	viesti = replace(viesti, ":9", "<img align=absmiddle src=smiley/tongue.gif>")
	viesti = replace(viesti, ";)", "<img align=absmiddle src=smiley/wink2.gif>")
	viesti = replace(viesti, ":x", "<img align=absmiddle src=smiley/wink.gif>")
	viesti = replace(viesti, "3-o", "<img align=absmiddle src=smiley/sex.gif>")
	viesti = replace(viesti, "3-O", "<img align=absmiddle src=smiley/sex.gif>")
	viesti = replace(viesti, "=/", "<img align=absmiddle src=smiley/smiley_blush.gif>")
	viesti = replace(viesti, "-o/", "<img align=absmiddle src=smiley/duckie.gif>")
	header = "<body bgcolor=""#102908"" onload=""" & extra & "parent.viestit.kirjoita_viesti('" & viesti & "')"">"
	session("uusin") = aika(49)
'	response.flush
	application.lock
	application("chat_alias") = alias
	application("chat_aliasto") = aliasto
	application("chat_block") = block
	application.unlock	
else
	header = "<body bgcolor=""#102908"">"
end if	
if nimet <> nimet_vanha then
	session("chat_nimet") = nimet
end if
response.write header
response.write "</body>"

else
	response.write "<body bgcolor=""#102908""></body>"
end if
%>
</html>