<!--#include file="verif_ident.asp"-->
<%
Zm=request.querystring("m")
if Zm=1 then message="Le mot de passe a bien &eacute;t&eacute; modifi&eacute;"
if Zm=2 then message="Les mots de passe ne sont pas identiques"
if Zm=3 then message="Le mot de passe actuel n'est pas le bon"
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
<%
if Zm<>"" then
Select Case Zm
   case 1
couleur="#00CC33"   
   case 2,3
couleur="#FF0000"
end Select
%>
<p style="color:<%=couleur%>"><strong><%=message%></strong></p>
<%end if%>  
  <form id="form1" name="form1" method="post" action="modif_admin2.asp">
    <table width="500" border="0" cellspacing="0" cellpadding="2">
      <tr>
        <td width="180">Ancien mot de passe : </td>
        <td width="320"><input type="password" name="old_pwd" id="old_pwd" /></td>
      </tr>
      <tr>
        <td width="180">Nouveau mot de passe :</td>
        <td width="320"><input type="password" name="new_pwd" id="new_pwd" /></td>
      </tr>
      <tr>
        <td width="180">Confirmation :</td>
        <td width="320"><input type="password" name="new_pwd2" id="new_pwd2" /></td>
      </tr>
      <tr>
        <td width="180">&nbsp;</td>
        <td width="320"><input type="submit" name="button" id="button" value="Modifier" /></td>
      </tr>
    </table>
  </form>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
