<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
SQLinfos="SELECT * from [news] order by id_news DESC"
Set rsinfos=server.Createobject("adodb.recordset")
rsinfos.open SQLinfos,conn,3,3
nbre_infos=rsinfos.recordcount

Zm=request.QueryString("m")
if Zm=1 then message="Info bien ajout&eacute;e"
if Zm=2 then message="Document bien ajouté"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" language="javascript">
<!--
function maj()
{
new Ajax.Updater('select1','maj_liste.asp', {method: 'get'});
new Ajax.Updater('select2','maj_liste2.asp', {method: 'get'});
$('titre').clear();
$('texte').clear();
$('titre').focus();
}

function supprimer()
{	
	var params='id='+$F('select2')
	new Ajax.Request('suppr_info.asp', {onSuccess:maj, asynchronous:true, method:'post', postBody:params});
}	

//-->
</script>
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">
<div>
<table width="550" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td width="275" height="50">
<form id="form_modif" name="form_modif" method="post" action="modif_info.asp">   
<strong>Modifier une news :</strong><br /> 
      <select name="select1" id="select1">
<%
if nbre_infos>0 then
rsinfos.movefirst
do while not rsinfos.eof
texte1=left(rsinfos("titre"),24)
%>      
        <option value="<%=rsinfos("id_news")%>"><%=texte1%></option>
<%
rsinfos.movenext
loop
end if
%>        
      </select>
      <br />
        <input name="button" type="submit" class="modif" id="button" value="MODIFIER" />
</form>        
    </td>
    <td width="275" height="50">
<strong>Supprimer une news:</strong><br />
      <select name="select2" id="select2">
<%
if nbre_infos>0 then
rsinfos.movefirst
do while not rsinfos.eof
texte2=left(rsinfos("titre"),24)
%>      
        <option value="<%=rsinfos("id_news")%>"><%=texte2%></option>
<%
rsinfos.movenext
loop
end if
%>  
      </select>
      <br />
      <input name="button2" type="submit" class="suppr" id="button2" value="SUPPRIMER" onclick="supprimer();" />
    </td>
  </tr>
</table>
</div>
<hr />
<div>
<div>
<strong>Ajouter des documents pour les infos de l'accueil</strong>
  <form action="ajout_doc_infos.asp" method="post" enctype="multipart/form-data" name="form2" id="form2">
    <input name="fileField" type="file" id="fileField" size="45" />
      <input type="submit" name="button4" id="button4" value="Ajouter" />
  </form>
</div>
<%
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
dir = Server.MapPath("../../upload/docs_infos/")
set foldPt = FSO.GetFolder(dir)
set fc = foldPt.Files
if fc.count>0 then
%>
<div>
 <select name="liste_docs" id="liste_docs">
 <option value="">Liste des documents disponibles</option>
<%for each f in fc%>
    <option value="<%=f.name%>"><%=f.name%></option>
<%next%>    
  </select>
<SCRIPT LANGUAGE="JavaScript">
function copyclipboard(intext) {
window.clipboardData.setData('Text', intext);
} 
</SCRIPT>

  <input type="submit" name="button5" id="button5" value="Copier le nom" onClick="copyclipboard('$$'+$F(liste_docs)+'##')">
</div>
<%end if%>  
<hr />
 <div>
<form action="add_info.asp" method="post" enctype="multipart/form-data" name="form1" id="form1">
<table width="550" border="0" cellspacing="0" cellpadding="2">
<caption align="left" style="padding-bottom:10px"><br />
<%if message<>"" then%>
<p style="color:#00CC33"><strong><%=message%></strong></p>
<%end if%>
<strong>Ajouter une info sur la page d'accueil :</strong>
</caption>
  <tr>
    <td width="100">Titre :</td>
    <td width="450"><input name="titre" type="text" id="titre" size="53" /></td>
  </tr>
  <tr>
    <td width="100">Texte :</td>
    <td width="450"><textarea name="texte" id="texte" cols="50" rows="5"></textarea></td>
  </tr>
  <tr>
    <td width="100">Image :</td>
    <td width="450"><input name="image" type="file" id="image" size="40" /></td>
  </tr>
  <tr>
    <td>Position :</td>
    <td>
      <select name="position" id="position">
        <option value="left">gauche</option>
        <option value="right">droite</option>
      </select>
    </td>
  </tr>
  <tr>
    <td width="100">&nbsp;</td>
    <td width="450"><input type="submit" name="button3" id="button3" value="Ajouter" /></td>
  </tr>
</table>
</form>
 </div>
</div>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
