<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zcoord=request.Form("coord")

SQLlistedispos="SELECT * from [dispos_coord_gav] where avo_code="&Zcoord&" order by annee_dispo,mois_dispo"
Set rslistedispos=server.Createobject("adodb.recordset")
rslistedispos.open SQLlistedispos,conn2,3,3
nbre_listedispos=rslistedispos.recordcount

SQLchoixcoord="SELECT * from [Coordinateurs_GAV] where avo_code="&Zcoord
Set rschoixcoord=server.Createobject("adodb.recordset")
rschoixcoord.open SQLchoixcoord,conn,3,3
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
<div id="menu">
<div id="users">
<ul>
<li><a href="modif_admin.asp">Mot de passe administrateur</a></li>
<li><a href="gestion_login.asp">Gestion des login / mot de passe</a></li>
<li><a href="maj_base.asp">Mise &agrave; jour de la base</a></li>
</ul>
</div>
<div id="infos">
<ul>
<li><a href="infos.asp">Infos accueil</a></li>
<li><a href="news.asp">News d&eacute;filantes accueil</a></li>
</ul>
</div>
<div id="docs">
<ul>
<li><a href="documents.asp">Gestion des documents</a></li>
<li><a href="upload/upload.asp">Vademecum</a></li>
</ul>
</div>
<div id="vacances">
<ul>
<li><a href="vacances_jud.asp">Vacances Judiciaires</a></li>
</ul>
</div>
<div id="plannings">
<ul>
<li><a href="infojour.asp">Infos jour - perm classique</a></li>
<li><a href="planning_gav.asp">Planning Garde A Vue</a></li>
<li><a href="gene_planning_sos.asp">Planning SOS Victimes</a></li>

<li><a href="compteurs.asp">Liste des compteurs</a></li>
<li><a href="raz.asp">RAZ des compteurs</a></li>
</ul>
</div>
<div id="forum">
<ul>
<li><a href="../forum/admin.asp" target="_blank'">Administration du forum</a></li>
</ul>
</div>
<div id="deconnect">
<ul>
<li><a href="logout.asp">Se d&eacute;connecter</a></li>
</ul>
</div>
</div>
<div id="main">
<p><strong>Supprimer une dispo :</strong></p>
<p><strong><%=rschoixcoord("avo_prenom")%>&nbsp;<%=rschoixcoord("avo_nom")%></strong></p>
<%
if nbre_listedispos=0 then
response.write("aucune dispo pour ce coordinateur")
else
%>
<table width="250" border="0" cellspacing="0" cellpadding="2">
<%
rslistedispos.movefirst
do while not rslistedispos.eof
%>
  <tr>
    <td width="200"><%=rslistedispos("mois_dispo")%>/<%=rslistedispos("annee_dispo")%></td>
    <td width="50" align="center"><a href="suppr_dispos_coord2.asp?id=<%=rslistedispos("id_dispo")%>"><img src="imgs/supprimer.png" width="15" height="15" border="0" /></a></td>
  </tr>
<%
rslistedispos.movenext
loop
%>  
</table>
<%end if%>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
