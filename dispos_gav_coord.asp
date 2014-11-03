<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLliste="SELECT * from [Coordinateurs_GAV] order by avo_nom"
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3
nbre_liste=rsliste.recordcount
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../js/prototype.js"></script>
<script language="javascript">
<!--
function changebox(id)
{	
	if ($(id).checked == true) {
	$(id).checked=true;
	new Ajax.Request('changebox2.asp?valeur='+id+'&action=1');
	}
	else {
	$(id).checked=false;
	new Ajax.Request('changebox2.asp?valeur='+id+'&action=0');
	}
}	
//-->
</script>
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">
<p><strong><%=year(date())%></strong></p>
<table width="550" border="0" cellspacing="0" cellpadding="2" style="margin:0">
  <tr>
    <td width="200" height="25" align="center"><strong>COORDINATEURS</strong></td>
    <td width="25" height="25" align="center"><strong>J</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>F</strong></td>
    <td width="25" height="25" align="center"><strong>M</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>A</strong></td>
    <td width="25" height="25" align="center"><strong>M</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>J</strong></td>
    <td width="25" height="25" align="center"><strong>J</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>A</strong></td>
    <td width="25" height="25" align="center"><strong>S</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>O</strong></td>
    <td width="25" height="25" align="center"><strong>N</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>D</strong></td>
    </tr>
  <%
if nbre_liste>0 then
rsliste.movefirst
do while not rsliste.eof
%>
  <tr>
  <td width="200" align="center"><%=ucase(rsliste("avo_nom"))%></td>
<%
for i=1 to 12
SQLcoches="SELECT * from [dispos_coord_gav] where avo_code="&rsliste("avo_code")&" and mois_dispo="&i&" and annee_dispo="&year(date())
Set rscoches=server.Createobject("adodb.recordset")
rscoches.open SQLcoches,conn2,3,3
nbre_coches=rscoches.recordcount
j=i
if j<10 then j="0"&i
%>
 <td width="25" align="center"><input name="<%=j%><%=year(date())%><%=rsliste("avo_code")%>" type="checkbox" id="<%=j%><%=year(date())%><%=rsliste("avo_code")%>" onclick="changebox('<%=j%><%=year(date())%><%=rsliste("avo_code")%>');"<%if nbre_coches>0 then response.write(" checked='checked'")%>/></td>
<%next%>
  </tr>
<%
rsliste.movenext
loop
end if
%>
</table>
<p><strong><%=year(date())+1%></strong></p>
<table width="550" border="0" cellspacing="0" cellpadding="2" style="margin:0">
  <tr>
    <td width="200" height="25" align="center"><strong>COORDINATEURS</strong></td>
    <td width="25" height="25" align="center"><strong>J</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>F</strong></td>
    <td width="25" height="25" align="center"><strong>M</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>A</strong></td>
    <td width="25" height="25" align="center"><strong>M</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>J</strong></td>
    <td width="25" height="25" align="center"><strong>J</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>A</strong></td>
    <td width="25" height="25" align="center"><strong>S</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>O</strong></td>
    <td width="25" height="25" align="center"><strong>N</strong></td>
    <td width="25" height="25" align="center" bgcolor="#CCCCCC"><strong>D</strong></td>
    </tr>
  <%
if nbre_liste>0 then
rsliste.movefirst
do while not rsliste.eof
%>
  <tr>
  <td width="200" align="center"><%=ucase(rsliste("avo_nom"))%></td>
<%
for i=1 to 12
SQLcoches="SELECT * from [dispos_coord_gav] where avo_code="&rsliste("avo_code")&" and mois_dispo="&i&" and annee_dispo="&year(date())+1
Set rscoches=server.Createobject("adodb.recordset")
rscoches.open SQLcoches,conn2,3,3
nbre_coches=rscoches.recordcount
j=i
if j<10 then j="0"&i
%>
 <td width="25" align="center"><input name="<%=j%><%=year(date())+1%><%=rsliste("avo_code")%>" type="checkbox" id="<%=j%><%=year(date())+1%><%=rsliste("avo_code")%>" onclick="changebox('<%=j%><%=year(date())+1%><%=rsliste("avo_code")%>');"<%if nbre_coches>0 then response.write(" checked='checked'")%>/></td>
<%next%>
  </tr>
<%
rsliste.movenext
loop
end if
%>
</table>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
