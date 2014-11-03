<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLdispo="SELECT * from [dispos_gav] order by avo_nom"
Set rsdispo=server.Createobject("adodb.recordset")
rsdispo.open SQLdispo,conn2,3,3
rsdispo.movefirst
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
  <table style="width: 300px;">
  <%do while not rsdispo.eof%>
    <tr style="width: 80%;">
      <td style=" border-bottom: 1px solid #CCC; padding: 5px 2px;"><%=rsdispo("avo_nom")%></td>
      <td><%=rsdispo("cpt")%></td>
    </tr>
  <%
  rsdispo.movenext
  loop
  %>
  </table>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
