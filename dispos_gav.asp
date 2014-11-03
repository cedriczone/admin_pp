<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLliste="SELECT * from [Intervenants_GAV] order by avo_nom"
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3
nbre_liste=rsliste.recordcount
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
if nbre_liste>0 then
i=0
rsliste.movefirst
do while not rsliste.eof
i=i+1

SQLcoches="SELECT * from [dispos_gav] where num_avo="&rsliste("avo_code")
Set rscoches=server.Createobject("adodb.recordset")
rscoches.open SQLcoches,conn2,3,3
nbre_coches=rscoches.recordcount
if nbre_coches>0 then
%>
<tr>
	<td width="200" align="center"><%=ucase(rsliste("avo_nom"))%></td>
 	<td width="25" align="center"><input name="lun<%=rsliste("avo_code")%>" type="checkbox" id="lun<%=rsliste("avo_code")%>"<%if rscoches("lun")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="mar<%=rsliste("avo_code")%>" type="checkbox" id="mar<%=rsliste("avo_code")%>"<%if rscoches("mar")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center"><input name="mer<%=rsliste("avo_code")%>" type="checkbox" id="mer<%=rsliste("avo_code")%>"<%if rscoches("mer")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="jeu<%=rsliste("avo_code")%>" type="checkbox" id="jeu<%=rsliste("avo_code")%>"<%if rscoches("jeu")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center"><input name="ven<%=rsliste("avo_code")%>" type="checkbox" id="ven<%=rsliste("avo_code")%>"<%if rscoches("ven")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="sam<%=rsliste("avo_code")%>" type="checkbox" id="sam<%=rsliste("avo_code")%>"<%if rscoches("sam")=1 then response.write(" checked='checked'")%>/></td>
        <td width="25" align="center"><input name="dim<%=rsliste("avo_code")%>" type="checkbox" id="dim<%=rsliste("avo_code")%>"<%if rscoches("dim")=1 then response.write(" checked='checked'")%>/></td>
      </tr>
<%else%>      
<tr>
  	<td width="200" align="center"><%=ucase(rsliste("avo_nom"))%></td>
 	<td width="25" align="center"><input name="lun<%=rsliste("avo_code")%>" type="checkbox" id="lun<%=rsliste("avo_code")%>"/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="mar<%=rsliste("avo_code")%>" type="checkbox" id="mar<%=rsliste("avo_code")%>"/></td>
        <td width="25" align="center"><input name="mer<%=rsliste("avo_code")%>" type="checkbox" id="mer<%=rsliste("avo_code")%>"/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="jeu<%=rsliste("avo_code")%>" type="checkbox" id="jeu<%=rsliste("avo_code")%>"/></td>
        <td width="25" align="center"><input name="ven<%=rsliste("avo_code")%>" type="checkbox" id="ven<%=rsliste("avo_code")%>"/></td>
        <td width="25" align="center" bgcolor="#CCCCCC"><input name="sam<%=rsliste("avo_code")%>" type="checkbox" id="sam<%=rsliste("avo_code")%>"/></td>
        <td width="25" align="center"><input name="dim<%=rsliste("avo_code")%>" type="checkbox" id="dim<%=rsliste("avo_code")%>"/></td>
      </tr>
<%
end if
rsliste.movenext
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
