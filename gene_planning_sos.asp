<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<%
SQLliste="SELECT * from [intervenants_sos] order by avo_nom"
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3

SQLliste2="SELECT * from [avocats_pp] order by avo_nom"
Set rsliste2=server.Createobject("adodb.recordset")
rsliste2.open SQLliste2,conn,3,3
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
<%
Zm=request.QueryString("m")
if Zm=1 then response.write("<p style='color:#00CC33'><strong>planning SOS g&eacute;n&eacute;r&eacute;</strong></p>")
%>
<p>Génération du planning SOS Victimes :</p>
<form id="form1" name="form1" method="post" action="gene_planning_sos2.asp">
  <p>Date de d&eacute;part : 
    <input type="text" name="jour" id="jour" value="<%=date()%>" />
  </p>
  <p>Premier avocat : 
    <select name="num" id="num">
<%
rsliste.movefirst
do while not rsliste.eof
%>    
      <option value="<%=rsliste("avo_code")%>"><%=rsliste("avo_nom")%>&nbsp;<%=rsliste("avo_prenom")%></option>
<%
rsliste.movenext
loop
%>
    </select>
  </p>
  <p>
    <input type="submit" name="button" id="button" value="Générer le planning SOS victimes" />
  </p>
</form>
<hr />
<form id="form2" name="form2" method="post" action="echange_sos.asp">
  <p>Echanger le planning de 2 avocats :</p>
  <p>Date (jj/mm/aaaa) : 
    <input name="date_echange" type="text" id="date_echange" size="12" />
    <input type="submit" name="button2" id="button2" value="Visualiser" />
  </p>
</form>
<hr />
<form id="form3" name="form3" method="post" action="remplacement_global_sos.asp">
  <p>Remplacer un avocat de fa&ccedil;on globale :</p>
  <p>
    <select name="avocat" id="avocat">
      <%
rsliste2.movefirst
do while not rsliste2.eof
%>    
      <option value="<%=rsliste2("avo_code")%>"><%=rsliste2("avo_nom")%>&nbsp;<%=rsliste2("avo_prenom")%></option>
        <%
rsliste2.movenext
loop
%>
    </select>
<br />par<br /><select name="avocat2" id="avocat2">
      <%
rsliste.movefirst
do while not rsliste.eof
%>    
      <option value="<%=rsliste("avo_code")%>"><%=rsliste("avo_nom")%>&nbsp;<%=rsliste("avo_prenom")%></option>
        <%
rsliste.movenext
loop
%>
    </select></p>
  <p>Laisser les dates vides pour remplacer toutes les dates &agrave; partir d'aujourd'hui.</p>
  <p>Du (jj/mm/aaaa) : 
    <input name="date_debut" type="text" id="date_debut" size="12" />
    au : 
    <input name="date_fin" type="text" id="date_fin" size="12" />
    <input type="submit" name="button2" id="button2" value="Remplacer" />
  </p>
</form>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
