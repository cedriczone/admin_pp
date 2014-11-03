<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<!--#include file="connexion.asp"-->
<%
SQLpermutation="SELECT * FROM [permutation] where validee=0 order by id_permutation"
Set rspermutation=server.Createobject("adodb.recordset")
rspermutation.open SQLpermutation,conn2,3,3
nbre_permutation=rspermutation.recordcount
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
<p><strong>Permutations en attente de validation</strong></p>
<div>
  <table width="800" border="0" cellspacing="2" cellpadding="0">
    <tr>
      <td width="40" align="center"><strong>type</strong></td>
      <td width="120" height="30" align="center"><strong>remplac&eacute;</strong></td>
      <td width="70" align="center"><strong>date</strong></td>
      <td width="120" align="center"><strong>coord</strong></td>
      <td width="120" height="30" align="center"><strong>rempla&ccedil;ant</strong></td>
      <td width="70" height="30" align="center"><strong>date</strong></td>
      <td width="120" align="center"><strong>coord</strong></td>
      <td width="80" height="30" align="center"><strong>tel</strong></td>
      <td width="50" height="30" align="center">&nbsp;</td>
      <td height="30" align="center">&nbsp;</td>
    </tr>
  </table>
<%
if nbre_permutation>0 then
i=0
rspermutation.movefirst
do while not rspermutation.eof
i=i+1

SQLremplace="SELECT * FROM [Avocats_PP] where avo_code="&rspermutation("num_avocat")
Set rsremplace=server.Createobject("adodb.recordset")
rsremplace.open SQLremplace,conn,3,3
rsremplace.movefirst

SQLremplacant="SELECT * FROM [Avocats_PP] where avo_code="&rspermutation("num_permute")
Set rsremplacant=server.Createobject("adodb.recordset")
rsremplacant.open SQLremplacant,conn,3,3
rsremplacant.movefirst

SQLcoord1="SELECT * FROM [Avocats_PP] where avo_code="&rspermutation("coord1")
Set rscoord1=server.Createobject("adodb.recordset")
rscoord1.open SQLcoord1,conn,3,3
rscoord1.movefirst

SQLcoord2="SELECT * FROM [Avocats_PP] where avo_code="&rspermutation("coord2")
Set rscoord2=server.Createobject("adodb.recordset")
rscoord2.open SQLcoord2,conn,3,3
rscoord2.movefirst
%>  
  <form id="form<%=i%>" name="form<%=i%>" method="post" action="valid_permutation.asp?id=<%=rspermutation("id_permutation")%>">
  <table width="800" border="0" cellspacing="2" cellpadding="0" style="font-size:9px">
  <tr>
    <td width="40" align="center" style="border-right:#666 solid 1px"><%=rspermutation("type_planning")%></td>
      <td width="120" height="25" align="center" style="border-right:#666 solid 1px"><%if rspermutation("demande")=rspermutation("num_avocat") then response.write("<strong>")%><%=rsremplace("avo_libelle")%><%if rspermutation("demande")=rspermutation("num_avocat") then response.write("</strong>")%></td>
      <td width="70" height="25" align="center" style="border-right:#666 solid 1px"><%=rspermutation("date_perm")%></td>
      <td width="120" height="25" align="center" style="border-right:#666 solid 1px"><%=rscoord1("avo_libelle")%></td>
      <td width="120" height="25" align="center" style="border-right:#666 solid 1px"><%if rspermutation("demande")=rspermutation("num_permute") then response.write("<strong>")%><%=rsremplacant("avo_libelle")%><%if rspermutation("demande")=rspermutation("num_permute") then response.write("</strong>")%></td>
      <td width="70" height="25" align="center" style="border-right:#666 solid 1px"><%=rspermutation("date_permute")%></td>
      <td width="120" align="center" style="border-right:#666 solid 1px"><%=rscoord2("avo_libelle")%></td>
      <td width="80" align="center" style="border-right:#666 solid 1px"><%=rspermutation("portable")%></td>
      <td width="50" align="center"><input type="submit" name="button" id="button" value="OK" /></td>
      <td><a href="suppr_permutation.asp?id=<%=rspermutation("id_permutation")%>">Supprimer</a></td>
    </tr>
    </table>
  </form>
<%
rspermutation.movenext
loop
end if
%>   
</div>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
