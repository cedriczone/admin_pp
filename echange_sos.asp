<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zdate=request.form("date_echange")
Zzdate=Zdate
jour_d=DatePart("d", Zdate)
mois_d=DatePart("m", Zdate)
annee_d=DatePart("yyyy", Zdate)
Zdate=mois_d&"/"&jour_d&"/"&annee_d
SQLdatesos="SELECT * from [planning_sos] where jour=#"&Zdate&"#"
Set rsdatesos=server.Createobject("adodb.recordset")
rsdatesos.open SQLdatesos,conn2,3,3
nbredatesos=rsdatesos.recordcount

SQLliste="SELECT * from [intervenants_sos] order by avo_nom"
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3
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
<p>Planning SOS : Echanger : <%=Zzdate%></p>
<%if nbredatesos=0 then%>
<p>Aucune date ne correspond dans le planning</p>
<%
else

SQLtitulaire="SELECT * from [avocats_pp] where avo_code="&rsdatesos("titulaire")
Set rstitulaire=server.Createobject("adodb.recordset")
rstitulaire.open SQLtitulaire,conn,3,3

SQLsuppleant="SELECT * from [avocats_pp] where avo_code="&rsdatesos("suppleant")
Set rssuppleant=server.Createobject("adodb.recordset")
rssuppleant.open SQLsuppleant,conn,3,3
%>
<p></p>
<form id="form1" name="form1" method="post" action="echange_sos2.asp?ech=tit">
<p>titulaire : <br />
  <input name="date_ech" type="hidden" id="date_ech" value="<%=Zdate%>" />
    <input name="ancien_tit" type="hidden" id="ancien_tit" value="<%=rstitulaire("avo_code")%>" />
    <input name="ancien_tit2" type="text" disabled="disabled" id="ancien_tit2" value="<%=rstitulaire("avo_prenom")%>&nbsp;<%=rstitulaire("avo_nom")%>" size="37" />
<span style="font-size:9px">remplac&eacute; par :</span><br />
<select name="nveau_tit" id="nveau_tit">
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
  <input type="submit" name="button" id="button" value="Valider" />
</p>
</form>
<form id="form2" name="form2" method="post" action="echange_sos2.asp?ech=sup">
  <p>
    suppl&eacute;ant :<br />
      <input name="date_ech" type="hidden" id="date_ech" value="<%=Zdate%>" />
      <input name="ancien_sup" type="hidden" id="ancien_sup" value="<%=rssuppleant("avo_code")%>" />
    <input name="ancien_sup2" type="text" disabled="disabled" id="ancien_sup2" value="<%=rssuppleant("avo_prenom")%>&nbsp;<%=rssuppleant("avo_nom")%>" size="37" />
<span style="font-size:9px">remplac&eacute; par :</span><br />
    <select name="nveau_sup" id="nveau_sup">
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
    <input type="submit" name="button2" id="button2" value="Valider" />
  </p>
</form>
<p> </p>
<%end if%>
</div>
<div id="footer">
  &nbsp;
</div>
</body>
</html>
