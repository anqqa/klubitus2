<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<!-- #INCLUDE FILE="../dbfuncs.asp" -->
<%
pvm = date()
p = datepart("d", pvm)
kk = datepart("m", pvm)
v = datepart("v", pvm)
kaupunki = "helsinki"
%>
<wml>
  <card id="logo" title="K L U B I T U S" ontimer="bileet.asp">
    <timer value="50"/>
    <p align="center">
<img id="klubitus" src="klubitus.wbmp" alt="Klubitus" width="50" height="56"/>
    </p>
    <do type="accept" name="etusivu" label="T&#xE4;n&#xE4;&#xE4;n">
	  <go href="bileet.asp?p=<%=p %>&kk=<%kk %>&v=<%v %>&kaupunki=<%=kaupunki %>">
	    <setvar name="p" value="<%=p %>"/>
		<setvar name="kk" value="<%=kk %>"/>
		<setvar name="v" value="<%=v %>"/>
		<setvar name="kaupunki" value="<%=kaupunki %>"/>
	  </go>
    </do>
  </card>

</wml>

