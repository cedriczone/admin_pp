<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<!--#include file="connexion.asp"-->
<%
SQLusers="SELECT * from [login] order by login"
Set rsusers=server.Createobject("adodb.recordset")
rsusers.open SQLusers,conn2,3,3
nbre_users=rsusers.recordcount

SQLavocats="SELECT * from [Avocats_PP] order by avo_libelle"
Set rsavocats=server.Createobject("adodb.recordset")
rsavocats.open SQLavocats,conn,3,3

Zm=request.QueryString("m")
if Zm=1 then message="Utilisateur bien ajout&eacute;"
if Zm=2 then message="Identifiants envoy&eacute;s avec succ&egrave;s"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="js/prototype.js"></script>
<script type="text/javascript" src="js/livevalidation.js"></script>
<script type="text/javascript" language="javascript">
<!--
function verif_login()
{	
	var params='login='+encodeURIComponent($F('login'))
	new Ajax.Updater('verifuser','verif_user.asp', {asynchronous:true, method:'post', postBody:params});
}	
//-->

<!--
function Choixpage(numpage){
if(numpage==1){document.form_suppr.action="suppr_users.asp";}
if(numpage==2){document.form_suppr.action="envoi_passe.asp";}
document.form_suppr.submit();
}
//-->
</script>
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
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">&nbsp;
<p>Gestion des login/mot de passe pour l'acc&egrave;s sur le site public</p>
<p>Cr&eacute;er un nouvel utilisateur :</p>
<%if message<>"" then%>
<p style="color:#00CC33"><strong><%=message%></strong></p>
<%end if%>
<div id="ajout_doc">
<FORM ACTION="ajout_user.asp" METHOD="POST">
<TABLE width="550" class="bords_pointilles">

<TR>
  <TD ALIGN="RIGHT" VALIGN="TOP">Avocat :</TD>
  <TD ALIGN="LEFT">
  <select name="avo" id="avo">
    <option value="0" selected="selected">Aucun</option>
<%
rsavocats.movefirst
do while not rsavocats.eof
%>    
    <option value="<%=rsavocats("avo_code")%>"><%=rsavocats("avo_libelle")%></option>
<%
rsavocats.movenext
loop
%>
  </select>
  </TD>
</TR>
<TR>
  <TD ALIGN="RIGHT" VALIGN="TOP">login :</TD>
  <TD ALIGN="LEFT"><input name="login" type="text" id="login" size="30" onblur="verif_login(); return false;" /><span id="verifuser"></span></TD>
</TR>

<TR>
  <TD width="190" ALIGN="RIGHT">mot de passe :</TD>
  <TD width="360" ALIGN="LEFT"><input name="pwd1" type="password" id="pwd1" size="30" /></TD>
</TR>
<TR>
  <TD ALIGN="RIGHT">confirmation :</TD>
  <TD ALIGN="LEFT"><input name="pwd2" type="password" id="pwd2" size="30" /></TD>
</TR>
<TR>
  <TD ALIGN="RIGHT">email :</TD>
  <TD ALIGN="LEFT"><input name="email" type="text" id="email" size="30" /></TD>
</TR>
<TR>
	<TD width="190" ALIGN="RIGHT">&nbsp;</TD>
	<TD width="360" ALIGN="LEFT"><INPUT TYPE="SUBMIT" NAME="SUB1" VALUE="Cr&eacute;er l'utilisateur"></TD>
</TR>
</TABLE>
</FORM>
<script language="javascript">
var f0 = new LiveValidation('pwd1');
f0.add( Validate.Presence );
f0.add( Validate.Length, { minimum: 4 } );
var f1 = new LiveValidation('pwd2');
f1.add( Validate.Presence );
f1.add( Validate.Confirmation, { match: 'pwd1', failureMessage: "pas identiques!", validMessage: "OK" } );
var f2 = new LiveValidation('email', { failureMessage: "format incorrect", validMessage: "OK" } );
f2.add( Validate.Email );
var f3 = new LiveValidation('login');
f3.add( Validate.Presence );
</script>
</div>
<div id="liste_docs">
<p>Liste des utilisateurs :</p>
  <form id="form_suppr" name="form_suppr" method="post" action="suppr_users.asp">
    <table width="758" border="0" cellpadding="2" cellspacing="0" class="bords_pointilles">
      <tr>
        <td width="30" align="center">&nbsp;</td>
        <td width="180" align="center"><strong>nom</strong></td>
        <td width="180" height="30" align="center"><strong>login</strong></td>
        <td width="100" align="center"><strong>passe</strong></td>
        <td width="180" height="30" align="center"><strong>email</strong></td>
        <td width="30" align="center">&nbsp;</td>
        <td width="30" align="center">&nbsp;</td>
      </tr>
<%
if nbre_users>0 then
i=0
rsusers.movefirst
do while not rsusers.eof
i=i+1

SQLnompre="SELECT * from [Avocats_PP] where avo_code="&rsusers("avo")
Set rsnompre=server.Createobject("adodb.recordset")
rsnompre.open SQLnompre,conn,3,3
nbre_nompre=rsnompre.recordcount
if nbre_nompre>0 then
rsnompre.movefirst
libelle=rsnompre("avo_libelle")
else
libelle=""
end if
%>      
      <tr>
        <td width="30" align="center"><input name="checkbox<%=i%>" type="checkbox" id="checkbox<%=i%>" value="<%=rsusers("id_login")%>" /></td>
        <td width="180" align="center"><%=libelle%></td>
        <td width="180" align="center"><%=rsusers("login")%></td>
        <td width="100" align="center"><%=rsusers("password")%></td>
        <td width="180" align="center" style="font-size:9px"><%=rsusers("email")%></td>
        <td width="30" align="center"><A HREF="#" onClick="window.open('modif_user_mail.asp?id=<%=rsusers("id_login")%>','_blank','toolbar=0, location=0, directories=0, status=0, scrollbars=0, resizable=0, copyhistory=0, menuBar=0, width=450, height=200');return(false)"><img src="imgs/mail.jpg" alt="modifier l'email" width="20" height="20" border="0" /></a></td>
        <td width="30" align="center"><A HREF="#" onClick="window.open('modif_user_pwd.asp?id=<%=rsusers("id_login")%>','_blank','toolbar=0, location=0, directories=0, status=0, scrollbars=0, resizable=0, copyhistory=0, menuBar=0, width=400, height=450');return(false)"><img src="imgs/cadenas.jpg" alt="modifier le mot de passe" width="20" height="20" border="0" /></a></td>
      </tr>
<%
rsusers.movenext
loop
end if
%>      
    </table>
    
    <input name="nbre_box" type="hidden" id="nbre_box" value="<%=i%>" />
    <input name="button1" type="button" onClick ="javascript:Choixpage(2)" value="Envoyer les identifiants a la selection" />
    <input name="button2" type="button" onClick ="javascript:Choixpage(1)" value="Supprimer la s&eacute;lection" />
  </form>
  </div>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
