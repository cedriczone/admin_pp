<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLderniermois="SELECT * from [planning_gav] order by date_gav DESC"
Set rsderniermois=server.Createobject("adodb.recordset")
rsderniermois.open SQLderniermois,conn2,3,3
nbre_derniermois=rsderniermois.recordcount
if nbre_derniermois>0 then
rsderniermois.movefirst
Zderniermois=rsderniermois("date_gav")
derniermois=month(Zderniermois)
end if
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
  <p>&nbsp;</p>
  <p><strong>Remise A Zero des compteurs :</strong></p>
  <form id="form1" name="form1" method="post" action="raz_coord.asp">
    <input type="submit" name="button" id="button" value="Remise A Zero des compteurs Coordinateurs" />
  </form>
  <form id="form1" name="form1" method="post" action="raz_inter.asp">
    <input type="submit" name="button2" id="button2" value="Remise A zero des compteurs Intervenants" />
    
  </form>
  <p>&nbsp;</p>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
