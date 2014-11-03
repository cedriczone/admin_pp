<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zm=request.QueryString("m")
if Zm=1 then
message="MISE A JOUR REUSSIE"
couleur_message="#00CC33"
elseif Zm=2 then message="LE FORMAT DE LA BASE N'EST PAS BON"
couleur_message="#FF0000"
end if

SQLcat="SELECT DISTINCT cat from [docs] order by cat"
Set rscat=server.Createobject("adodb.recordset")
rscat.open SQLcat,conn,3,3
nbre_cat=rscat.recordcount

SQLdocs="SELECT * from [docs] order by nom_doc"
Set rsdocs=server.Createobject("adodb.recordset")
rsdocs.open SQLdocs,conn,3,3
nbre_docs=rsdocs.recordcount
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
<p>Mise &agrave; jour de la base</p>
<div id="ajout_doc">
<%if Zm<>"" then%><h2 style="color:<%=couleur_message%>"><%=message%></h2><%end if%>
<FORM ACTION="maj_base2.asp" ENCTYPE="MULTIPART/FORM-DATA" METHOD="POST">
<TABLE width="550" class="bords_pointilles">

<TR>
	<TD width="190" ALIGN="RIGHT" VALIGN="TOP">S&eacute;lectionner la base :</TD>
	<TD width="360" ALIGN="LEFT"><INPUT NAME="myFile" TYPE="FILE" size="40">
	  <BR>	</TD>
</TR>
<TR>
  <TD width="190" ALIGN="RIGHT">&nbsp;</TD>
  <TD width="360" ALIGN="LEFT">&nbsp;</TD>
</TR>
<TR>
	<TD width="190" ALIGN="RIGHT">&nbsp;</TD>
	<TD width="360" ALIGN="LEFT"><INPUT TYPE="SUBMIT" NAME="SUB1" VALUE="Mettre &agrave; jour"></TD>
</TR>
</TABLE>
</FORM>
</div>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
