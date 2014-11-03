<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<!--#include file="connexion.asp"-->
<%
SQLremplacement="SELECT * FROM [remplacement] where validee=0 order by id_remplace"
Set rsremplacement=server.Createobject("adodb.recordset")
rsremplacement.open SQLremplacement,conn2,3,3
nbre_remplacement=rsremplacement.recordcount
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
<div id="main" style="width:800px">&nbsp;
<p><strong>Remplacements en attente de validation</strong></p>
<div>
  <table width="800" border="0" cellspacing="2" cellpadding="0">
    <tr>
      <td width="50" align="center"><strong>type</strong></td>
      <td width="170" height="30" align="center"><strong>remplac&eacute;</strong></td>
      <td width="170" height="30" align="center"><strong>rempla&ccedil;ant</strong></td>
      <td width="70" height="30" align="center"><strong>date</strong></td>
      <td width="70" height="30" align="center"><strong>tel</strong></td>
      <td width="170" height="30" align="center"><strong>coordinateur</strong></td>
      <td width="50" align="center">&nbsp;</td>
      <td align="center">&nbsp;</td>
    </tr>
  </table>
<%
if nbre_remplacement>0 then
i=0
rsremplacement.movefirst
do while not rsremplacement.eof
i=i+1

SQLremplace="SELECT * FROM [Avocats_PP] where avo_code="&rsremplacement("remplace")
Set rsremplace=server.Createobject("adodb.recordset")
rsremplace.open SQLremplace,conn,3,3
rsremplace.movefirst

SQLremplacant="SELECT * FROM [Avocats_PP] where avo_code="&rsremplacement("remplacant")
Set rsremplacant=server.Createobject("adodb.recordset")
rsremplacant.open SQLremplacant,conn,3,3
rsremplacant.movefirst

SQLcoord="SELECT * FROM [Avocats_PP] where avo_code="&rsremplacement("coord")
Set rscoord=server.Createobject("adodb.recordset")
rscoord.open SQLcoord,conn,3,3
nb_coord = rscoord.recordcount
if nb_coord>0 then rscoord.movefirst
%>
  <form id="form<%=i%>" name="form<%=i%>" method="post" action="valid_remplacement.asp?id=<%=rsremplacement("id_remplace")%>">
<table width="800" border="0" cellspacing="2" cellpadding="0" style="font-size:9px">
  <tr>
    <td width="50" align="center" style="border-right:#666 solid 1px"><%=rsremplacement("type_planning")%></td>
      <td width="170" height="25" align="center" style="border-right:#666 solid 1px"><%if rsremplacement("demande")=rsremplacement("remplace") then response.write("<strong>")%><%=rsremplace("avo_libelle")%><%if rsremplacement("demande")=rsremplacement("remplace") then response.write("</strong>")%></td>
      <td width="170" height="25" align="center" style="border-right:#666 solid 1px"><%if rsremplacement("demande")=rsremplacement("remplacant") then response.write("<strong>")%><%=rsremplacant("avo_libelle")%><%if rsremplacement("demande")=rsremplacement("remplacant") then response.write("</strong>")%></td>
      <td width="70" height="25" align="center" style="border-right:#666 solid 1px"><%=rsremplacement("date_remplace")%></td>
      <td width="70" height="25" align="center" style="border-right:#666 solid 1px"><%=rsremplacement("tel")%></td>
      <td width="170" height="25" align="center" style="border-right:#666 solid 1px"><% if nb_coord>0 then response.write(rscoord("avo_libelle"))%></td>
      <td width="50"><input type="submit" name="button" id="button" value="OK" /></td>
      <td><a href="suppr_remplacement.asp?id=<%=rsremplacement("id_remplace")%>">Supprimer</a></td>
   </tr>
    </table>
  </form>
<%
rsremplacement.movenext
loop
end if
%>  
</div>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
