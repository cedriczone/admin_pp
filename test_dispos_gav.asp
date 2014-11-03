<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="test_connexion2.asp"-->
<%
SQLliste="SELECT * from [Intervenants_GAV] order by avo_nom"
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
<style type="text/css">
<!--
input {
	margin: 0px;
	padding: 0px;
}
-->
</style>
<script type="text/javascript" src="../js/prototype.js"></script>
<script language="javascript">
<!--
function changebox(id)
{	
	if ($(id).checked == true) {
	$(id).checked=true;
	new Ajax.Request('test_changebox.asp?valeur='+id+'&action=1');
	}
	else {
	$(id).checked=false;
	new Ajax.Request('test_changebox.asp?valeur='+id+'&action=0');
	}
}	
//-->
</script>
</head>

<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">
&nbsp;
<p>Gestion des disponibilit&eacute;s GAV</p>
<form id="form1" name="form1" method="post" action="">
<table width="550" border="0" cellspacing="0" cellpadding="2" style="margin:0">
<tr>
  <td width="200" height="25" align="center"><strong>INTERVENANTS</strong></td>
  <td width="200" height="25" align="center"><strong>Semaine</strong></td>
  <td width="200" height="25" align="center" bgcolor="#CCCCCC">&nbsp;</td>
  <td width="200" height="25" align="center"><strong>Week end</strong></td>
  <td width="200" height="25" align="center" bgcolor="#CCCCCC">&nbsp;</td>
</tr>
<%
if nbre_liste>0 then
i=0
rsliste.movefirst
do while not rsliste.eof
i=i+1

SQLcoches="SELECT * from [dispos_gav] where num_avo="&rsliste("avo_code")
Set rscoches=server.Createobject("adodb.recordset")
rscoches.open SQLcoches,conn2,3,3
nbre_coches=rscoches.recordcount
if nbre_coches>0 then
%>
<tr>
  		<td width="200" align="center"><%=ucase(rsliste("avo_nom"))%></td>
 		<td width="200" align="center"><input name="sej<%=rsliste("avo_code")%>" type="checkbox" id="sej<%=rsliste("avo_code")%>" onclick="changebox('sej<%=rsliste("avo_code")%>');"<%if rscoches("sej")=1 then response.write(" checked='checked'")%>/></td>
        <td width="200" align="center" bgcolor="#CCCCCC"><input name="sen<%=rsliste("avo_code")%>" type="checkbox" id="sen<%=rsliste("avo_code")%>" onclick="changebox('sen<%=rsliste("avo_code")%>');"<%if rscoches("sen")=1 then response.write(" checked='checked'")%>/></td>
        <td width="200" align="center"><input name="wej<%=rsliste("avo_code")%>" type="checkbox" id="wej<%=rsliste("avo_code")%>" onclick="changebox('wej<%=rsliste("avo_code")%>');"<%if rscoches("wej")=1 then response.write(" checked='checked'")%>/></td>
        <td width="200" align="center" bgcolor="#CCCCCC"><input name="wen<%=rsliste("avo_code")%>" type="checkbox" id="wen<%=rsliste("avo_code")%>" onclick="changebox('wen<%=rsliste("avo_code")%>');"<%if rscoches("wen")=1 then response.write(" checked='checked'")%>/></td>
      </tr>
<%else%>      
<tr>
  		<td width="200" align="center"><%=ucase(rsliste("avo_nom"))%></td>
 		<td width="200" align="center"><input name="sej<%=rsliste("avo_code")%>" type="checkbox" id="sej<%=rsliste("avo_code")%>" onclick="changebox('sej<%=rsliste("avo_code")%>');"/></td>
        <td width="200" align="center" bgcolor="#CCCCCC"><input name="sen<%=rsliste("avo_code")%>" type="checkbox" id="sen<%=rsliste("avo_code")%>" onclick="changebox('sen<%=rsliste("avo_code")%>');"/></td>
        <td width="200" align="center"><input name="wej<%=rsliste("avo_code")%>" type="checkbox" id="wej<%=rsliste("avo_code")%>" onclick="changebox('wej<%=rsliste("avo_code")%>');"/></td>
        <td width="200" align="center" bgcolor="#CCCCCC"><input name="wen<%=rsliste("avo_code")%>" type="checkbox" id="wen<%=rsliste("avo_code")%>" onclick="changebox('wen<%=rsliste("avo_code")%>');"/></td>
      </tr>
<%
end if
rsliste.movenext
loop
end if
%>      
    </table>
  </form>
  <p>&nbsp;</p>
</div>
<div id="footer">
    &nbsp;
</div>
</body>
</html>
