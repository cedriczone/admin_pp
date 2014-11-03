<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
'SI CERTAINES DISPOS CORRESPONDENT A DES AVOCATS QUI NE SONT NI INTERVENANTS NI COORDINATEURS ON LES ENLEVE
SQLvide_dispos="SELECT * FROM [dispos_gav] order by num_avo"
Set rsvide_dispos=server.Createobject("adodb.recordset")
rsvide_dispos.open SQLvide_dispos,conn2,3,3
nb_vide_dispos_gav=rsvide_dispos.recordcount
if nb_vide_dispos_gav>0 then
	rsvide_dispos.movefirst
	do while not rsvide_dispos.eof
		SQLcompare="SELECT * FROM [Intervenants_GAV] WHERE avo_code="&rsvide_dispos("num_avo")
		Set rscompare=server.Createobject("adodb.recordset")
		rscompare.open SQLcompare,conn,3,3
		nb_compare=rscompare.recordcount
		if nb_compare=0 then
			SQLcompare2="SELECT * FROM [Coordinateurs_GAV] WHERE avo_code="&rsvide_dispos("num_avo")
			Set rscompare2=server.Createobject("adodb.recordset")
			rscompare2.open SQLcompare2,conn,3,3
			nb_compare2=rscompare2.recordcount
			if nb_compare2=0 then
				SQLsuppr="DELETE * FROM [dispos_gav] WHERE num_avo="&rsvide_dispos("num_avo")
				Set suppr= Server.CreateObject("ADODB.RecordSet")
				suppr.open SQLsuppr,conn2
			end if
		end if
	rsvide_dispos.movenext
	loop
end if

'///OPERATION INVERSE: SI UN AVOCAT N'EXISTE PAS DANS LES DISPOS ON LE CREE

SQLliste="SELECT * from [Intervenants_GAV] order by avo_code"
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3
rsliste.movefirst
do while not rsliste.eof
	SQLdispos_gav="SELECT * FROM [dispos_gav] WHERE num_avo="&rsliste("avo_code")
	Set rsdispos_gav=server.Createobject("adodb.recordset")
	rsdispos_gav.open SQLdispos_gav,conn2,3,3
	nb_dispos_gav=rsdispos_gav.recordcount
	if nb_dispos_gav=0 then
		if rsliste("AVO_GAVCHE")="Vrai" then chev=1
		if rsliste("AVO_GAVCHE")="Faux" then chev=0
		SQLadd="Insert Into [dispos_gav](num_avo,avo_nom,lun,mar,mer,jeu,ven,sam,dim,chevronne,cpt) Values("&rsliste("avo_code")&",'"&replace(rsliste("avo_libelle"),"'","''")&"',0,0,0,0,0,0,0,"&chev&",0)"
		Set add= Server.CreateObject("ADODB.RecordSet")
		add.open SQLadd,conn2	
	end if
rsliste.movenext
loop

SQLliste2="SELECT * from [Coordinateurs_GAV] order by avo_code"
Set rsliste2=server.Createobject("adodb.recordset")
rsliste2.open SQLliste2,conn,3,3
rsliste2.movefirst
do while not rsliste2.eof
	SQLdispos_gav2="SELECT * FROM [dispos_gav] WHERE num_avo="&rsliste2("avo_code")
	Set rsdispos_gav2=server.Createobject("adodb.recordset")
	rsdispos_gav2.open SQLdispos_gav2,conn2,3,3
	nb_dispos_gav2=rsdispos_gav2.recordcount
	if nb_dispos_gav2=0 then
		SQLadd2="Insert Into [dispos_gav](num_avo,avo_nom,lun,mar,mer,jeu,ven,sam,dim,chevronne,cpt) Values("&rsliste2("avo_code")&",'"&replace(rsliste2("avo_libelle"),"'","''")&"',0,0,0,0,0,0,0,1,0)"
		Set add2= Server.CreateObject("ADODB.RecordSet")
		add2.open SQLadd2,conn2		
	end if
rsliste2.movenext
loop
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Administration</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
input {
	margin: 0px;
	padding: 0px;
}
-->
</style>
<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.6.2/jquery.min.js"></script>
<script type="text/javascript">
	$(document).ready(function() {
		
		$("input[type=checkbox]").click(function(){
			var id, action;
			id=$(this).attr("id");
			if ($(this).is(':checked')){
				action = 1;
			}else{
				action = 0;
			}
			$.ajax({  url: "changebox.asp?valeur="+id+"&action="+action });
		});	
	});//fin jquery
</script>

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
</ul>
</div>
<div id="docs">
<ul>
<li><a href="documents.asp">Gestion des documents</a></li>
</ul>
</div>
<div id="plannings">
<ul>

<li><a href="planning_gav.asp">Planning Garde A Vue</a></li>
<li><a href="gene_planning_sos.asp">Planning SOS Victimes</a></li>
<li><a href="audiences.asp">Saisie des audiences</a></li>
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
&nbsp;
<p>Gestion des disponibilit&eacute;s GAV</p>
<form id="form1" name="form1" method="post" action="">
<table width="550" border="0" cellspacing="0" cellpadding="2" style="margin:0">
<tr>
  <td width="200" height="25" align="center"><strong>INTERVENANTS</strong></td>
  <td width="25" height="25" align="center"><strong>L</strong></td>
  <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>M</strong></td>
  <td width="25" height="25" align="center"><strong>M</strong></td>
  <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>J</strong></td>
  <td width="25" height="25" align="center"><strong>V</strong></td>
  <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>S</strong></td>
  <td width="25" height="25" align="center"><strong>D</strong></td>
</tr>
<%
SQLliste_dispos="SELECT * FROM [dispos_gav] order by avo_nom"
Set rsliste_dispos=server.Createobject("adodb.recordset")
rsliste_dispos.open SQLliste_dispos,conn2,3,3
nbre_liste_dispos=rsliste_dispos.recordcount
if nbre_liste_dispos>0 then
i=0
rsliste_dispos.movefirst
do while not rsliste_dispos.eof
i=i+1

SQLcoches="SELECT * from [dispos_gav] where num_avo="&rsliste_dispos("avo_code")
Set rscoches=server.Createobject("adodb.recordset")
rscoches.open SQLcoches,conn2,3,3
nbre_coches=rscoches.recordcount
if nbre_coches>0 then
%>
<tr>
	<td width="200" align="center"><%=ucase(rsliste_dispos("avo_nom"))%></td>
 	<td width="25" align="center"><input name="lun<%=rsliste_dispos("avo_code")%>" type="checkbox" id="lun<%=rsliste_dispos("avo_code")%>"<%if rscoches("lun")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="mar<%=rsliste_dispos("avo_code")%>" type="checkbox" id="mar<%=rsliste_dispos("avo_code")%>"<%if rscoches("mar")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center"><input name="mer<%=rsliste_dispos("avo_code")%>" type="checkbox" id="mer<%=rsliste_dispos("avo_code")%>"<%if rscoches("mer")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="jeu<%=rsliste_dispos("avo_code")%>" type="checkbox" id="jeu<%=rsliste_dispos("avo_code")%>"<%if rscoches("jeu")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center"><input name="ven<%=rsliste_dispos("avo_code")%>" type="checkbox" id="ven<%=rsliste_dispos("avo_code")%>"<%if rscoches("ven")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="sam<%=rsliste_dispos("avo_code")%>" type="checkbox" id="sam<%=rsliste_dispos("avo_code")%>"<%if rscoches("sam")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center"><input name="dim<%=rsliste_dispos("avo_code")%>" type="checkbox" id="dim<%=rsliste_dispos("avo_code")%>"<%if rscoches("dim")=1 then response.write(" checked='checked'")%>/></td>
      </tr>
<%else%>      
<tr>
  	<td width="200" align="center"><%=ucase(rsliste_dispos("avo_nom"))%></td>
 	<td width="25" align="center"><input name="lun<%=rsliste_dispos("avo_code")%>" type="checkbox" id="lun<%=rsliste_dispos("avo_code")%>"/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="mar<%=rsliste_dispos("avo_code")%>" type="checkbox" id="mar<%=rsliste_dispos("avo_code")%>"/></td>
        <td width="25" align="center"><input name="mer<%=rsliste_dispos("avo_code")%>" type="checkbox" id="mer<%=rsliste_dispos("avo_code")%>"/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="jeu<%=rsliste_dispos("avo_code")%>" type="checkbox" id="jeu<%=rsliste_dispos("avo_code")%>"/></td>
        <td width="25" align="center"><input name="ven<%=rsliste_dispos("avo_code")%>" type="checkbox" id="ven<%=rsliste_dispos("avo_code")%>"/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="sam<%=rsliste_dispos("avo_code")%>" type="checkbox" id="sam<%=rsliste_dispos("avo_code")%>"/></td>
        <td width="25" align="center"><input name="dim<%=rsliste_dispos("avo_code")%>" type="checkbox" id="dim<%=rsliste_dispos("avo_code")%>"/></td>
      </tr>
<%
end if
rsliste_dispos.movenext
loop
end if
%>      
    </table>
  </form>
  <p>&nbsp;</p>
</div>
<div id="footer">
    &nbsp;
</div>
</body>
</html>
