<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zid=request.form("select1")
Zid=cint(Zid)
SQLinfo="SELECT * from [defilantes] where id_defilante="&Zid
Set rsinfo=server.Createobject("adodb.recordset")
rsinfo.open SQLinfo,conn2,3,3
rsinfo.movefirst
function tarea(text)
         tarea=replace(text,"&amp;","&")
         tarea=replace(tarea,"&lt;","<")
		 tarea=replace(tarea,"&gt;",">")
		 tarea=replace(tarea,"<br>",VbCrLf)
		 tarea=replace(tarea,"''","'")
end function
texte=tarea(rsinfo("texte_defilante"))
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
<div>
<div>
 <div>
   <form id="form1" name="form1" method="post" action="modif_news2.asp?id=<%=Zid%>">
     <table width="550" border="0" cellspacing="0" cellpadding="2">
       <caption align="left" style="padding-bottom:10px">
         <br />
         Modifier une info sur la page d'accueil
         </caption>
       <tr>
         <td width="100">Texte :</td>
         <td width="450"><input name="texte" type="text" id="texte" value="<%=rsinfo("texte_defilante")%>" size="60" /></td>
       </tr>
       <tr>
         <td width="100">&nbsp;</td>
         <td width="450">&nbsp;</td>
       </tr>
       <tr>
         <td width="100">&nbsp;</td>
         <td width="450"><input type="submit" name="button3" id="button3" value="Modifier"/></td>
       </tr>
     </table>
      </form>
   </div>
</div>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
