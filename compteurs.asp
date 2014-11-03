<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLcptinter="SELECT * FROM [repartition_gav] where archive=0 order by nbre_gav DESC"
Set rscptinter=server.Createobject("adodb.recordset")
rscptinter.open SQLcptinter,conn2,3,3
nbre_cptinter=rscptinter.recordcount

SQLcptcoord="SELECT * FROM [repartition_coord_gav] where archive=0 order by nbre_coord_gav DESC"
Set rscptcoord=server.Createobject("adodb.recordset")
rscptcoord.open SQLcptcoord,conn2,3,3
nbre_cptcoord=rscptcoord.recordcount
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
<div id="main">&nbsp;
  <table width="550" border="0" cellspacing="0" cellpadding="2">
    <tr>
      <td valign="top">
      <table width="270" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td height="30" bgcolor="#F0F0F0"><strong>Intervenants</strong></td>
          <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
        </tr>
<%
if nbre_cptinter>0 then
rscptinter.movefirst
do while not rscptinter.eof
SQLinter="SELECT * FROM [Intervenants_GAV] where avo_code="&rscptinter("avo_code")
Set rsinter=server.Createobject("adodb.recordset")
rsinter.open SQLinter,conn,3,3
nb_inter=rsinter.recordcount
if nb_inter>0 then
%>        
        <tr>
          <td width="230" height="20"><%=rsinter("avo_nom")%>&nbsp;<%=rsinter("avo_prenom")%></td>
          <td width="40" height="20"><%=rscptinter("nbre_gav")%></td>
        </tr>
<%
end if
rscptinter.movenext
loop
end if
%>        
      </table></td>
      <td valign="top">
      <table width="270" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td height="30" bgcolor="#F0F0F0"><strong>Coordinateurs</strong><br /></td>
          <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
        </tr>
<%
if nbre_cptcoord>0 then
rscptcoord.movefirst
do while not rscptcoord.eof
SQLcoord="SELECT * FROM [Coordinateurs_GAV] where avo_code="&rscptcoord("avo_code")
Set rscoord=server.Createobject("adodb.recordset")
rscoord.open SQLcoord,conn,3,3
%>             
        <tr>
          <td width="230" height="20"><%=rscoord("avo_nom")%>&nbsp;<%=rscoord("avo_prenom")%></td>
          <td width="40" height="20"><%=rscptcoord("nbre_coord_gav")%></td>
        </tr>
<%
rscptcoord.movenext
loop
end if
%>         
      </table></td>
    </tr>
  </table>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
