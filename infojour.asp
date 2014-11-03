<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
SQLinfojour="SELECT * from [infojour] order by date_info"
Set rsinfojour=server.Createobject("adodb.recordset")
rsinfojour.open SQLinfojour,conn,3,3
nbre_infojour=rsinfojour.recordcount
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
  <p><strong>Ajouter une info pour une journ&eacute;e du planning perm classique</strong></p><form id="form2" name="form2" method="post" action="infojour2.asp">
  <p>Date (jj/mm/aaaa) : 
    <input name="date_info" type="text" id="date_info" size="12" />
  </p>
  <p>Info :
    <textarea name="info" id="info" cols="60" rows="5"></textarea>
    <br />
    <br />
    <input type="submit" name="button" id="button" value="Ajouter" />
  </p>
  </form>
<%
if nbre_infojour>0 then
rsinfojour.movefirst
%>  
  <p><strong>Modifier une info :</strong></p>
  <form id="form1" name="form1" method="post" action="modif_infojour.asp">
    <select name="id" id="id">
<%
do while not rsinfojour.eof  
Ztexte=left(rsinfojour("texte_info"),25)
%>
      <option value="<%=rsinfojour("id_infojour")%>"><%=rsinfojour("date_info")%> - <%=Ztexte%></option>
<%
rsinfojour.movenext
loop
%>
    </select>
    <input type="submit" name="button2" id="button2" value="Modifier" />
  </form>
<%end if%>  
  <p>&nbsp;</p>
  <%
if nbre_infojour>0 then
rsinfojour.movefirst
%>  
  <p><strong>Supprimer une info :</strong></p>
  <form id="form1" name="form3" method="post" action="suppr_infojour.asp">
    <select name="id" id="id">
<%
do while not rsinfojour.eof  
Ztexte=left(rsinfojour("texte_info"),25)
%>
      <option value="<%=rsinfojour("id_infojour")%>"><%=rsinfojour("date_info")%> - <%=Ztexte%></option>
<%
rsinfojour.movenext
loop
%>
    </select>
    <input type="submit" name="button2" id="button2" value="Supprimer" />
  </form>
<%end if%>  
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
