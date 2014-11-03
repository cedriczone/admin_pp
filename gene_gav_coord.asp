<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLliste="SELECT * from [Coordinateurs_GAV] order by avo_nom"
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
<script type="text/javascript" src="js/prototype.js"></script>
<link rel="STYLESHEET" type="text/css" href="../dhtmlxcalendar.css">
<script>
      window.dhx_globalImgPath="../imgs/";
</script>
<script src="../dhtmlxcommon.js"></script>
<script src="../dhtmlxcalendar.js"></script>
<script>
	var cal1, cal2, mCal, mDCal, newStyleSheet;

	var dateFrom = null;
	var dateTo = null;
	
		function calendrier() {
		cal1 = new dhtmlxCalendarObject('date1');
		cal2 = new dhtmlxCalendarObject('date2');
		dhtmlxCalendarLangModules['fr'] = {
		langname:	'fr',
		dateformat:	'%d/%m/%Y',
		monthesFNames:	["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Decembre"],
		monthesSNames:	["Jan", "Fev", "Mar", "Avr", "Mai", "Jun", "Jul", "Aou", "Sep", "Oct", "Nov", "Dec"],
		daysFNames:	["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"],
		daysSNames:	["Di", "Lu", "Ma", "Me", "Je", "Ve", "Sa"],
		weekend:	[0, 6],
		weekstart:	1,
		msgClose:	 "Fermer",
		msgMinimize: "Réduire",
		msgToday:	 "Aujourd'hui"
	}

		cal1.loadUserLanguage('fr');
		cal2.loadUserLanguage('fr');
	}	
	
function changedate() {
	if (document.getElementById('mois').value<<%=month(now)%>)
	{
cal1.setDate('01/'+document.getElementById('mois').value+'/<%=year(now)+1%>');	
cal2.setDate('27/'+document.getElementById('mois').value+'/<%=year(now)+1%>');
	}
	else
	{
cal1.setDate('01/'+document.getElementById('mois').value+'/<%=year(now)%>');	
cal2.setDate('27/'+document.getElementById('mois').value+'/<%=year(now)%>');
	}
}
</script>
</head>
<body onLoad="calendrier();">
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">&nbsp;
  <p>Dernier mois g&eacute;n&eacute;r&eacute; :
</p>
  <p>&nbsp;</p>
  <p><strong>G&eacute;n&eacute;rer un planning mensuel :</strong></p>
  <form id="form1" name="form1" method="post" action="gene_planning_coord.asp">
    <br />
    <table width="400" border="0" cellspacing="6" cellpadding="0">
      <tr>
        <td width="120">Mois &agrave; g&eacute;n&eacute;rer :</td>
        <td width="280">
        <select name="mois" id="mois" onchange="changedate();">
          <option value="1"<%if month(date())=1 then response.write(" selected")%>>Janvier</option>
          <option value="2"<%if month(date())=2 then response.write(" selected")%>>F&eacute;vrier</option>
          <option value="3"<%if month(date())=3 then response.write(" selected")%>>Mars</option>
          <option value="4"<%if month(date())=4 then response.write(" selected")%>>Avril</option>
          <option value="5"<%if month(date())=5 then response.write(" selected")%>>Mai</option>
          <option value="6"<%if month(date())=6 then response.write(" selected")%>>Juin</option>
          <option value="7"<%if month(date())=7 then response.write(" selected")%>>Juillet</option>
          <option value="8"<%if month(date())=8 then response.write(" selected")%>>Aout</option>
          <option value="9"<%if month(date())=9 then response.write(" selected")%>>Septembre</option>
          <option value="10"<%if month(date())=10 then response.write(" selected")%>>Octobre</option>
          <option value="11"<%if month(date())=11 then response.write(" selected")%>>Novembre</option>
          <option value="12"<%if month(date())=12 then response.write(" selected")%>>D&eacute;cembre</option>
        </select></td>
      </tr>
<%
if nbre_liste>0 then
rsliste.movefirst
%>      
      <tr>
        <td>Coordinateur 1 :</td>
        <td>
        <select name="coord1" id="coord1"> 
          <option value="99991">en attente</option>
<%do while not rsliste.eof%>               
          <option value="<%=rsliste("avo_code")%>"><%=ucase(rsliste("avo_nom"))%></option>
<%
rsliste.movenext
loop
%>          
        </select></td>
      </tr>
      <tr>
        <td>Coordinateur 2 :</td>
        <td><select name="coord2" id="coord2">
        <option value="99992">en attente</option>
<%
rsliste.movefirst
do while not rsliste.eof
%>        
          <option value="<%=rsliste("avo_code")%>"><%=ucase(rsliste("avo_nom"))%></option>
<%
rsliste.movenext
loop
%>           
        </select></td>
      </tr>
      <tr>
        <td>Coordinateur 3 :</td>
        <td><select name="coord3" id="coord3">
        <option value="99993">en attente</option>
<%
rsliste.movefirst
do while not rsliste.eof
%>         
          <option value="<%=rsliste("avo_code")%>"><%=ucase(rsliste("avo_nom"))%></option>
<%
rsliste.movenext
loop
%>           
        </select></td>
      </tr>
      <tr>
        <td width="120">&nbsp;</td>
        <td width="280"><input type="submit" name="button2" id="button2" value="G&eacute;n&eacute;rer" /></td>
      </tr>
<%end if%>      
    </table>
  </form>
  <p>&nbsp;</p>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
