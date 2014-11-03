<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zdate=request.form("date_echange")
Zzdate=Zdate
jour_d=DatePart("d", Zdate)
mois_d=DatePart("m", Zdate)
annee_d=DatePart("yyyy", Zdate)
Zdate=mois_d&"/"&jour_d&"/"&annee_d
jourencours=Weekday(Zzdate,2)

SQLdategav="SELECT * from [planning_gav] where date_gav=#"&Zdate&"#"
Set rsdategav=server.Createobject("adodb.recordset")
rsdategav.open SQLdategav,conn2,3,3
nbredategav=rsdategav.recordcount

SQLliste="SELECT * from [Intervenants_GAV] order by avo_nom"
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3

SQLliste2="SELECT * from [Coordinateurs_GAV] order by avo_nom"
Set rsliste2=server.Createobject("adodb.recordset")
rsliste2.open SQLliste2,conn,3,3
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
<p>Planning GAV : Echanger : <%=Zzdate%></p>
<%if nbredategav=0 then%>
<p>Aucune date ne correspond dans le planning</p>
<%
else
%>
<form id="form1" name="form1" method="post" action="remplacement2.asp">
<input name="date_ech" type="hidden" id="date_ech" value="<%=Zdate%>" />
<select name="remplace" id="remplace">
<%
rsdategav.movefirst
SQL1="SELECT * from [avocats_pp] where avo_code="&rsdategav("num_coordinateur")
Set rs1=server.Createobject("adodb.recordset")
rs1.open SQL1,conn,3,3
if rs1.recordcount>0 then rs1.movefirst
SQL2="SELECT * from [avocats_pp] where avo_code="&rsdategav("num_inter_chev")
Set rs2=server.Createobject("adodb.recordset")
rs2.open SQL2,conn,3,3
rs2.movefirst
SQL3="SELECT * from [avocats_pp] where avo_code="&rsdategav("num_inter1")
Set rs3=server.Createobject("adodb.recordset")
rs3.open SQL3,conn,3,3
rs3.movefirst
SQL4="SELECT * from [avocats_pp] where avo_code="&rsdategav("num_inter2")
Set rs4=server.Createobject("adodb.recordset")
rs4.open SQL4,conn,3,3
rs4.movefirst
SQL5="SELECT * from [avocats_pp] where avo_code="&rsdategav("num_inter3")
Set rs5=server.Createobject("adodb.recordset")
rs5.open SQL5,conn,3,3
rs5.movefirst
SQL6="SELECT * from [avocats_pp] where avo_code="&rsdategav("num_inter4")
Set rs6=server.Createobject("adodb.recordset")
rs6.open SQL6,conn,3,3
rs6.movefirst
%>    
<%if rs1.recordcount>0 then%><option value="1<%=rs1("avo_code")%>"><%=rs1("avo_libelle")%></option><%end if%>
<option value="2<%=rs2("avo_code")%>"><%=rs2("avo_libelle")%></option>
<option value="3<%=rs3("avo_code")%>"><%=rs3("avo_libelle")%></option>
<option value="4<%=rs4("avo_code")%>"><%=rs4("avo_libelle")%></option>
<option value="5<%=rs5("avo_code")%>"><%=rs5("avo_libelle")%></option>
<option value="6<%=rs6("avo_code")%>"><%=rs6("avo_libelle")%></option>
<%
Select Case jourencours
Case 1,2,3        
    SQL7="SELECT * from [avocats_pp] where avo_code="&rsdategav("num_inter5")
    Set rs7=server.Createobject("adodb.recordset")
    rs7.open SQL7,conn,3,3
    rs7.movefirst
    SQL8="SELECT * from [avocats_pp] where avo_code="&rsdategav("num_inter6")
    Set rs8=server.Createobject("adodb.recordset")
    rs8.open SQL8,conn,3,3
    rs8.movefirst
%>
<option value="7<%=rs7("avo_code")%>"><%=rs7("avo_libelle")%></option>
<option value="8<%=rs8("avo_code")%>"><%=rs8("avo_libelle")%></option>
<%end select%>
</select>
<span style="font-size:9px">remplac&eacute; par :</span><br />
<select name="remplacant" id="remplacant">
<%
rsliste.movefirst
do while not rsliste.eof
%>    
      <option value="<%=rsliste("avo_code")%>"><%=rsliste("avo_nom")%>&nbsp;<%=rsliste("avo_prenom")%></option>
<%
rsliste.movenext
loop
%>
<option>-----------------------------</option>
<%
rsliste2.movefirst
do while not rsliste2.eof
%>    
      <option value="<%=rsliste2("avo_code")%>"><%=rsliste2("avo_nom")%>&nbsp;<%=rsliste2("avo_prenom")%></option>
<%
rsliste2.movenext
loop
%>
    </select>
  <input type="submit" name="button" id="button" value="Valider" />
</p>
</form>
<p> </p>
<%end if%>
</div>
<div id="footer">
  &nbsp;
</div>
</body>
</html>
