<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
SQLinfos="SELECT * from [messagerie] order by id_message DESC"
Set rsinfos=server.Createobject("adodb.recordset")
rsinfos.open SQLinfos,conn,3,3
nbre_infos=rsinfos.recordcount

Zm=request.QueryString("m")
if Zm=1 then message="Message bien ajout&eacute;"
if Zm=2 then message="Message bien modifi&eacute;"
if Zm=3 then message="Message bien supprim&eacute;"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Messagerie</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.6.2/jquery.min.js"></script>
<script type="text/javascript">
	$(document).ready(function() {
		$('.active_add').click(function(){
			$('div#form_ajout').slideToggle();
		});
	});//fin jquery
</script>
</script>
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">
	<div><a href="#" class="active_add">Ajouter un message</a></div>
	<div class="space30"><!-- --></div>
	<div id="form_ajout" class="cache">
<!--FORMULAIRE AJOUT infos -->
	<form action="ajout_infos.asp" method="post" name="form_ajout_infos" id="form_ajout_infos">          
		<fieldset>
		  <p>
			<label for="titre">Titre</label>
			<input type="text" name="titre" id="titre" tabindex="30" />
			</p>
			<p>
			<label for="message">Message</label>
			<textarea name="message" id="message" cols="45" rows="5" tabindex="40"></textarea>
			</p>
		</fieldset>		
		<p><input type="submit" name="btn_envoi" id="btn_envoi" value="Envoyer" tabindex="150" /></p>
	</form>
</div>
<div class="space30"></div>
<%
if nbre_infos>0 then
rsinfos.movefirst
%>
<table width="600" border="0" cellpadding="0" cellspacing="0">
  <caption>Liste des infos</caption>
  
  <thead> <!-- En-tete du tableau -->
       <tr>
           <th width="100" valign="middle">Date</th>
           <th width="150" valign="middle">Titre</th>
           <th valign="middle">Message</th>
      </tr>
   </thead>
   
	<tbody> <!-- Corps du tableau -->
<%do while not rsinfos.eof%>
       <tr>
       	 <td width="100"><%=rsinfos("date_message")%></td>
   		 <td width="150"><%=rsinfos("titre")%></td>
   		 <td><%=left(rsinfos("texte"),120)%>&hellip;</td>
   		 <td width="40" align="center"><a href="modif_infos.asp?id=<%=rsinfos("id_message")%>"><img src="imgs/edit.png" width="24" height="24" alt="Editer" /></a></td>
   		 <td width="40" align="center"><a href="suppr_infos.asp?id=<%=rsinfos("id_message")%>" class="suppression"><img src="imgs/delete.png" width="24" height="24" alt="Supprimer" /></a></td>
      </tr>
<%
rsinfos.movenext
loop
%>      
</table>
<%end if%>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
