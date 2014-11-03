<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
SQLvacances="SELECT * from [vacances] order by id_plage"
Set rsvacances=server.Createobject("adodb.recordset")
rsvacances.open SQLvacances,conn,3,3
nbre_vacances=rsvacances.recordcount

Zm=request.QueryString("m")
if Zm=1 then
couleur="#FF0000"
message="Les deux dates doivent etre remplies"
elseif zm=2 then
couleur="#FF0000"
message="La deuxi&egrave;me date doit etre post&eacute;rieure &agrave; la seconde"
elseif Zm=3 then
couleur="#00CC33"
message="Plage de dates bien ajout&eacute;e"
elseif Zm=4 then
couleur="#00CC33"
message="Plage de dates bien supprim&eacute;e"
else
message=""
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">
  <p><strong>Vacances judiciaires :</strong></p>
<%if message<>"" then%>
<p style="color:<%=couleur%>"><strong><%=message%></strong></p>
<%end if%>
  <form id="form1" name="form1" method="post" action="ajout_plage.asp">
    <input name="debut" type="text" id="debut" size="15" />
   au 
   <input name="fin" type="text" id="fin" size="15" />
   <input type="submit" name="button" id="button" value="Ajouter cette plage" />
  </form>
<%
if nbre_vacances>0 then
rsvacances.movefirst
%>
<p>
<table width="280" border="0" cellpadding="2" cellspacing="0" bordercolor="#000000">
<%do while not rsvacances.eof%>
  <tr>
    <td width="20">Du</td>
    <td width="100"><strong>
<%
Zdebut=month(rsvacances("debut"))&"/"&day(rsvacances("debut"))&"/"&year(rsvacances("debut"))
Zdebut=datevalue(Zdebut)
response.write(Zdebut)
%>
</strong></td>
    <td width="20">au</td>
    <td width="100"><strong>
<%
Zfin=month(rsvacances("fin"))&"/"&day(rsvacances("fin"))&"/"&year(rsvacances("fin"))
Zfin=datevalue(Zfin)
response.write(Zfin)
%>
</strong></td>
    <td width="40" align="center"><a href="suppr_plage.asp?id=<%=rsvacances("id_plage")%>"><img src="imgs/supprimer.png" width="15" height="15" border="0" /></a></td>
  </tr>
<%
rsvacances.movenext  
loop
%>
</table>
</p>
<%end if%>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
