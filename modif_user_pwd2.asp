<!--#include file="connexion_perm.asp"-->
<%
Zid=request.querystring("id")
Znewpwd=request.form("newpwd1")

SQLmodif="UPDATE [login] set password='"&Znewpwd&"' WHERE id_login="&Zid
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-2" />
<title>Untitled Document</title>
<script language="Javascript">
function fermer(url)
{ 
eval("window.opener.parent.document.location.href='gestion_login.asp'");
window.close();
}
</script>
</head>
<body>
<p><strong>Mot de passe modifi&eacute; correctement</strong></p>
<p>&nbsp;</p>
<p><strong><A onClick=fermer(); style="cursor:hand">Fermer cette fenetre</A></strong></p>
</body>
</html>
