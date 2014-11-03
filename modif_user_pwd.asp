<!--#include file="connexion_perm.asp"-->
<%
Zid=request.querystring("id")

SQLuser="SELECT * from [login] where id_login="&Zid
Set rsuser=server.Createobject("adodb.recordset")
rsuser.open SQLuser,conn,3,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-2" />
<title>Untitled Document</title>
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/livevalidation.js"></script>
<style type="text/css">
<!--
.LV_validation_message{
    font-weight:bold;
    margin:0 0 0 5px;
}

.LV_valid {
    color:#00CC00;
}
	
.LV_invalid {
    color:#CC0000;
}
    
.LV_valid_field,
input.LV_valid_field:hover, 
input.LV_valid_field:active,
textarea.LV_valid_field:hover, 
textarea.LV_valid_field:active {
    border: 1px solid #00CC00;
}
    
.LV_invalid_field, 
input.LV_invalid_field:hover, 
input.LV_invalid_field:active,
textarea.LV_invalid_field:hover, 
textarea.LV_invalid_field:active {
    border: 1px solid #CC0000;
}
-->
</style>
</head>
<body>
<p><strong>Modifier le mot de passe de l'utilisateur : <%=rsuser("login")%></strong></p>
<p>Nouvel mot de passe :</p>
<form id="form1" name="form1" method="post" action="modif_user_pwd2.asp?id=<%=Zid%>">
  <table width="440" border="0" cellspacing="0" cellpadding="2">
    
    <tr>
      <td><input name="newpwd1" type="password" id="newpwd1" size="30" /></td>
    </tr>
    <tr>
      <td>confirmer :</td>
    </tr>
    <tr>
      <td><input name="newpwd2" type="password" id="newpwd2" size="30" /></td>
    </tr>
    <tr>
      <td><input type="submit" name="button" id="button" value="Valider" /></td>
    </tr>
  </table>
</form>
<script language="javascript">
var f5 = new LiveValidation('newpwd1');
f5.add( Validate.Presence );
f5.add( Validate.Length, { minimum: 4 } );
var f6 = new LiveValidation('newpwd2');
f6.add( Validate.Confirmation, { match: 'newpwd1', failureMessage: "pas identiques!", validMessage: "OK" } );
</script>
</body>
</html>
