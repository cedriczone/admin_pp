<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zupl=request.QueryString("upl")

SQLcat="SELECT * FROM [cat] where sur_cat=0 order by nom_cat"
Set rscat=server.Createobject("adodb.recordset")
rscat.open SQLcat,conn,3,3
nbre_cat=rscat.recordcount

SQLdocs="SELECT * from [docs] order by nom_doc"
Set rsdocs=server.Createobject("adodb.recordset")
rsdocs.open SQLdocs,conn,3,3
nbre_docs=rsdocs.recordcount
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script language="Javascript">
// ==================
//	Activations - Désactivations
// ==================
function GereControle(Controleur, Controle, Masquer) {
var objControleur = document.getElementById(Controleur);
var objControle = document.getElementById(Controle);
	if (Masquer=='1')
		objControle.style.visibility=(objControleur.checked==true)?'visible':'hidden';
	else
		objControle.disabled=(objControleur.checked==true)?false:true;
	return true;
}
</script>
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">&nbsp;
<p>Gestion des documents</p>
<%if Zupl=1 then%>
<p style="color:#16851C"><strong>Document ajout&eacute;</strong></p>
<%end if%>
<div id="ajout_doc">
<FORM ACTION="ajout_doc.asp" ENCTYPE="MULTIPART/FORM-DATA" METHOD="POST">
<TABLE width="550" class="bords_pointilles">
<TR>
  <TD width="190" ALIGN="RIGHT" VALIGN="TOP">Cat&eacute;gorie de documents :</TD>
  <TD width="360" ALIGN="LEFT">
  <select name="cat" id="cat">
<%
if nbre_cat>0 then
rscat.movefirst
do while not rscat.eof
%>  
    <option value="<%=rscat("id_cat")%>" style="color:#990000"><%=rscat("nom_cat")%></option>
<%
SQLcat2="SELECT * from [cat] where sur_cat="&rscat("id_cat")&" order by nom_cat"
Set rscat2=server.Createobject("adodb.recordset")
rscat2.open SQLcat2,conn,3,3
nbre_cat2=rscat2.recordcount
if nbre_cat2>0 then
rscat2.movefirst
do while not rscat2.eof
%>
<option value="<%=rscat2("id_cat")%>" style="color:#799247">&nbsp;&nbsp;&nbsp;&nbsp;<%=rscat2("nom_cat")%></option>
<%
SQLcat3="SELECT * from [cat] where sur_cat="&rscat2("id_cat")&" order by nom_cat"
Set rscat3=server.Createobject("adodb.recordset")
rscat3.open SQLcat3,conn,3,3
nbre_cat3=rscat3.recordcount
if nbre_cat3>0 then
rscat3.movefirst
do while not rscat3.eof
%>
<option value="<%=rscat3("id_cat")%>" style="color:#333333">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rscat3("nom_cat")%></option>
<%
rscat3.movenext
loop
end if

rscat2.movenext
loop
end if

rscat.movenext
loop
end if
%>
</select> 
</TD>
</TR>
<TR>
  <TD ALIGN="RIGHT" VALIGN="TOP">ou nouvelle cat&eacute;gorie :</TD>
  <TD ALIGN="LEFT"><input name="newcat" type="text" id="newcat" size="30" /></TD>
</TR>
<tr>
          <td width="200" height="20" align="right" class="titresmenu">Est-ce une sous-cat&eacute;gorie ? </td>
          <td width="336" height="20"><input name="verifsscat" type="checkbox" class="champs" id="verifsscat" onClick="GereControle('verifsscat', 'sur_cat', '1');"></td>
</tr>
<tr>
          <td height="20" align="right" class="titresmenu">Si oui, laquelle : </td>
          <td height="20"><select name="sur_cat" class="champs" id="sur_cat" style="visibility:hidden">
              <option value="" selected="selected">Cat&eacute;gories</option>
<%
if nbre_cat>0 then
rscat.movefirst
do while not rscat.eof
%>  
    <option value="<%=rscat("id_cat")%>" style="color:#990000"><%=rscat("nom_cat")%></option>
<%
SQLcat2="SELECT * from [cat] where sur_cat="&rscat("id_cat")&" order by nom_cat"
Set rscat2=server.Createobject("adodb.recordset")
rscat2.open SQLcat2,conn,3,3
nbre_cat2=rscat2.recordcount
if nbre_cat2>0 then
rscat2.movefirst
do while not rscat2.eof
%>
<option value="<%=rscat2("id_cat")%>" style="color:#799247">&nbsp;&nbsp;&nbsp;&nbsp;<%=rscat2("nom_cat")%></option>
<%
SQLcat3="SELECT * from [cat] where sur_cat="&rscat2("id_cat")&" order by nom_cat"
Set rscat3=server.Createobject("adodb.recordset")
rscat3.open SQLcat3,conn,3,3
nbre_cat3=rscat3.recordcount
if nbre_cat3>0 then
rscat3.movefirst
do while not rscat3.eof
%>
<option value="<%=rscat3("id_cat")%>" style="color:#333333">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rscat3("nom_cat")%></option>
<%
rscat3.movenext
loop
end if

rscat2.movenext
loop
end if

rscat.movenext
loop
end if
%>            </select>          </td>
</tr>
<TR>
  <TD ALIGN="RIGHT" VALIGN="TOP">&nbsp;</TD>
  <TD ALIGN="LEFT">&nbsp;</TD>
</TR>
<TR>
  <TD ALIGN="RIGHT" VALIGN="TOP">Description courte du document :</TD>
  <TD ALIGN="LEFT"><input name="description" type="text" id="description" size="55" /></TD>
</TR>
<TR>
  <TD ALIGN="RIGHT" VALIGN="TOP">&nbsp;</TD>
  <TD ALIGN="LEFT">&nbsp;</TD>
</TR>
<TR>
	<TD width="190" ALIGN="RIGHT" VALIGN="TOP">Choisir un fichier:</TD>
	<TD width="360" ALIGN="LEFT"><INPUT NAME="myFile" TYPE="FILE" size="40">
	  <BR>	</TD>
</TR>
<TR>
  <TD width="190" ALIGN="RIGHT">&nbsp;</TD>
  <TD width="360" ALIGN="LEFT">&nbsp;</TD>
</TR>
<TR>
	<TD width="190" ALIGN="RIGHT">&nbsp;</TD>
	<TD width="360" ALIGN="LEFT"><INPUT TYPE="SUBMIT" NAME="SUB1" VALUE="Envoyer le fichier"></TD>
</TR>
</TABLE>
</FORM>
</div>
<div id="modifcat">
<p>Modifier une cat&eacute;gorie</p>
<form id="form4" name="form4" method="post" action="modif_cat.asp">
  <p>
    <select name="cat" id="cat">
      <%
if nbre_cat>0 then
rscat.movefirst
do while not rscat.eof
%>  
      <option value="<%=rscat("id_cat")%>" style="color:#990000"><%=rscat("nom_cat")%></option>
        <%
SQLcat2="SELECT * from [cat] where sur_cat="&rscat("id_cat")&" order by nom_cat"
Set rscat2=server.Createobject("adodb.recordset")
rscat2.open SQLcat2,conn,3,3
nbre_cat2=rscat2.recordcount
if nbre_cat2>0 then
rscat2.movefirst
do while not rscat2.eof
%>
        <option value="<%=rscat2("id_cat")%>" style="color:#799247">&nbsp;&nbsp;&nbsp;&nbsp;<%=rscat2("nom_cat")%></option>
        <%
SQLcat3="SELECT * from [cat] where sur_cat="&rscat2("id_cat")&" order by nom_cat"
Set rscat3=server.Createobject("adodb.recordset")
rscat3.open SQLcat3,conn,3,3
nbre_cat3=rscat3.recordcount
if nbre_cat3>0 then
rscat3.movefirst
do while not rscat3.eof
%>
        <option value="<%=rscat3("id_cat")%>" style="color:#333333">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rscat3("nom_cat")%></option>
        <%
rscat3.movenext
loop
end if

rscat2.movenext
loop
end if

rscat.movenext
loop
end if
%>
    </select> 
  </p>
  <p>
    <input name="nveau_nom_cat" type="text" id="nveau_nom_cat" size="50" />
    <input type="submit" name="button" id="button" value="Modifier" />
  </p>
</form>
</div>
<div id="supprcat">
<p>Supprimer une cat&eacute;gorie</p>
<form id="form5" name="form5" method="post" action="suppr_cat.asp">
  <p>
    <select name="cat" id="cat">
      <%
if nbre_cat>0 then
rscat.movefirst
do while not rscat.eof
%>  
      <option value="<%=rscat("id_cat")%>" style="color:#990000"><%=rscat("nom_cat")%></option>
        <%
SQLcat2="SELECT * from [cat] where sur_cat="&rscat("id_cat")&" order by nom_cat"
Set rscat2=server.Createobject("adodb.recordset")
rscat2.open SQLcat2,conn,3,3
nbre_cat2=rscat2.recordcount
if nbre_cat2>0 then
rscat2.movefirst
do while not rscat2.eof
%>
        <option value="<%=rscat2("id_cat")%>" style="color:#799247">&nbsp;&nbsp;&nbsp;&nbsp;<%=rscat2("nom_cat")%></option>
        <%
SQLcat3="SELECT * from [cat] where sur_cat="&rscat2("id_cat")&" order by nom_cat"
Set rscat3=server.Createobject("adodb.recordset")
rscat3.open SQLcat3,conn,3,3
nbre_cat3=rscat3.recordcount
if nbre_cat3>0 then
rscat3.movefirst
do while not rscat3.eof
%>
        <option value="<%=rscat3("id_cat")%>" style="color:#333333">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rscat3("nom_cat")%></option>
        <%
rscat3.movenext
loop
end if

rscat2.movenext
loop
end if

rscat.movenext
loop
end if
%>
    </select>
    <input type="submit" name="button3" id="button3" value="Supprimer" />
  </p>
  </form>
</div>
<div id="liste_docs">
<p>Liste des documents sur le site</p>
  <form id="form_suppr" name="form_suppr" method="post" action="suppr_docs.asp">
    <table width="550" border="0" cellpadding="2" cellspacing="0" class="bords_pointilles">
      <tr>
        <td width="30" align="center">&nbsp;</td>
        <td width="60" height="30" align="center"><strong>Type</strong></td>
        <td width="280" height="30" align="center"><strong>Fichier</strong></td>
        <td width="180" height="30" align="center"><strong>Catégorie</strong></td>
      </tr>
      <%
if nbre_docs>0 then
rsdocs.movefirst
i=0
do while not rsdocs.eof
i=i+1
if i mod 2 <> 0 then bgc="#F3F2FF" else bgc="#FDFEEB"
ext=right(rsdocs("nom_doc"),4)

SQLcatdoc="SELECT * from [cat] where id_cat="&rsdocs("cat")
Set rscatdoc=server.Createobject("adodb.recordset")
rscatdoc.open SQLcatdoc,conn,3,3
nbredecat=rscatdoc.recordcount
if nbredecat>0 then
Zlacat=rscatdoc("nom_cat")
else
Zlacat="AUCUNE"
end if
%>
      <tr bgcolor="<%=bgc%>">
        <td width="30" align="center"><input name="checkbox<%=i%>" type="checkbox" id="checkbox<%=i%>" value="<%=rsdocs("id_doc")%>" /></td>
        <td width="60" align="center"><%=ext%></td>
        <td width="280" align="center"><a href="../../upload/docs_pp/<%=rsdocs("nom_doc")%>" target="_blank" class="liens_docs"><%=rsdocs("nom_doc")%></a></td>
        <td width="180" align="center"><%=Zlacat%></td>
      </tr>
<%
rsdocs.movenext
loop
end if
%>
    </table>
    
    <input name="nbre_box" type="hidden" id="nbre_box" value="<%=i%>" />
    <input type="submit" name="button2" id="button2" value="Supprimer la s&eacute;lection" />
  </form>
  </div>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
