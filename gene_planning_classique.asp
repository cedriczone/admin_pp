<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
'//// Coordinateurs PP
SQLliste_coord="SELECT * from [Coordinateurs_PP] order by avo_libelle"
Set rsliste_coord=server.Createobject("adodb.recordset")
rsliste_coord.open SQLliste_coord,conn,3,3

'//// Observateurs PP
SQLliste_obspp="SELECT * from [Observateurs_PP] order by avo_libelle"
Set rsliste_obspp=server.Createobject("adodb.recordset")
rsliste_obspp.open SQLliste_obspp,conn,3,3

'//// Observateurs MIN
SQLliste_obsmin="SELECT * from [Observateurs_MIN] order by avo_libelle"
Set rsliste_obsmin=server.Createobject("adodb.recordset")
rsliste_obsmin.open SQLliste_obsmin,conn,3,3

'//// Observateurs ETR
SQLliste_obsetr="SELECT * from [Observateurs_ETR] order by avo_libelle"
Set rsliste_obsetr=server.Createobject("adodb.recordset")
rsliste_obsetr.open SQLliste_obsetr,conn,3,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<link type="text/css" href="css/datepicker/jquery-ui-1.8.13.custom.css" rel="stylesheet" />
<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.6.2/jquery.min.js"></script>
<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.8.13/jquery-ui.min.js"></script>
<script type="text/javascript" charset="utf-8">
  $(document).ready(function(){
    //DATE PICKER
    $('#date1').datepicker();
    $('#date2').datepicker();

    $.datepicker.regional['fr'] = {
            closeText: 'Fermer',
            prevText: '&#x3c;Pr�c',
            nextText: 'Suiv&#x3e;',
            currentText: 'Courant',
            monthNames: ['Janvier','F�vrier','Mars','Avril','Mai','Juin',
            'Juillet','Aout','Septembre','Octobre','Novembre','D�cembre'],
            monthNamesShort: ['Jan','F�v','Mar','Avr','Mai','Jun',
            'Jul','Aou','Sep','Oct','Nov','D�c'],
            dayNames: ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'],
            dayNamesShort: ['Dim','Lun','Mar','Mer','Jeu','Ven','Sam'],
            dayNamesMin: ['Di','Lu','Ma','Me','Je','Ve','Sa'],
            weekHeader: 'Sm',
            dateFormat: 'dd/mm/yy',
            firstDay: 1,
            isRTL: false,
            showMonthAfterYear: false,
            yearSuffix: ''
            };
            
            $.datepicker.setDefaults($.datepicker.regional['fr']);
  });//fin jquery
</script>
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">&nbsp;
  <p>Dernier mois g&eacute;n&eacute;r&eacute; :
</p>
  <p>&nbsp;</p>
  <p><strong>G&eacute;n&eacute;rer un planning mensuel :</strong></p>
  <form id="form1" name="form1" method="post" action="gene_planning_classique2.asp">
    <br />
    <table width="800" border="0" cellspacing="6" cellpadding="0">
      <tr>
        <td width="420">Mois &agrave; g&eacute;n&eacute;rer :</td>
        <td width="380">
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
      <tr>
        <td width="420">Date d&eacute;but: </td>
        <td width="380"><input name="date1" type="text" id="date1" size="30"  value="Cliquer pour choisir un jour" readonly="true" style="font-size:11px; color:#333"/></td>
      </tr>
      <tr>
        <td width="420">Date fin:</td>
        <td width="380"><input name="date2" type="text" id="date2" size="30" value="Cliquer pour choisir un jour" readonly="true" style="font-size:11px; color:#333" /></td>
      </tr>
            <tr><td><hr></td></tr>

<%
rsliste_coord.movefirst
%>      
      <tr>
        <td>Coordinateur 1 :</td>
        <td>
        <select name="coord1" id="coord1"> 
          <option value="99991">choisir dans la liste</option>
<%do while not rsliste_coord.eof%>               
          <option value="<%=rsliste_coord("avo_code")%>"><%=ucase(rsliste_coord("avo_libelle"))%></option>
<%
rsliste_coord.movenext
loop
%>          
        </select></td>
      </tr>
<tr>
        <td>Observateurs :</td>
      <td>
        <select name="obs0" id="obs0"> 
          <option value="99994">Majeur1</option>
<%
rsliste_obspp.movefirst
do while not rsliste_obspp.eof%>               
          <option value="<%=rsliste_obspp("avo_code")%>"><%=ucase(rsliste_obspp("avo_libelle"))%></option>
<%
rsliste_obspp.movenext
loop
%>          
        </select></td>
        <td>
        <select name="obs1" id="obs1"> 
          <option value="99994">Majeur2</option>
<%
rsliste_obspp.movefirst
do while not rsliste_obspp.eof%>               
          <option value="<%=rsliste_obspp("avo_code")%>"><%=ucase(rsliste_obspp("avo_libelle"))%></option>
<%
rsliste_obspp.movenext
loop
%>          
        </select></td>
        <td>
        <select name="obs2" id="obs2"> 
          <option value="99994">Mineurs</option>
<%
rsliste_obsmin.movefirst
do while not rsliste_obsmin.eof%>               
          <option value="<%=rsliste_obsmin("avo_code")%>"><%=ucase(rsliste_obsmin("avo_libelle"))%></option>
<%
rsliste_obsmin.movenext
loop
%>          
        </select></td>
        <td>
        <select name="obs3" id="obs3"> 
          <option value="99994">Etrangers</option>
<%
rsliste_obsetr.movefirst
do while not rsliste_obsetr.eof%>               
          <option value="<%=rsliste_obsetr("avo_code")%>"><%=ucase(rsliste_obsetr("avo_libelle"))%></option>
<%
rsliste_obsetr.movenext
loop
%>          
        </select></td>
      </tr>
      <tr><td><hr></td></tr>
      <tr>
        <td>Coordinateur 2 :</td>
        <td><select name="coord2" id="coord2">
        <option value="99992">en attente</option>
<%
rsliste_coord.movefirst
do while not rsliste_coord.eof
%>        
          <option value="<%=rsliste_coord("avo_code")%>"><%=ucase(rsliste_coord("avo_libelle"))%></option>
<%
rsliste_coord.movenext
loop
%>           
        </select></td>
      </tr>
<tr>
        <td>Observateurs :</td>
        <td>
        <select name="obs4" id="obs4"> 
          <option value="99994">Majeur1</option>
<%
rsliste_obspp.movefirst
do while not rsliste_obspp.eof%>               
          <option value="<%=rsliste_obspp("avo_code")%>"><%=ucase(rsliste_obspp("avo_libelle"))%></option>
<%
rsliste_obspp.movenext
loop
%>          
        </select></td>
                <td>
        <select name="obs5" id="obs5"> 
          <option value="99994">Majeur2</option>
<%
rsliste_obspp.movefirst
do while not rsliste_obspp.eof%>               
          <option value="<%=rsliste_obspp("avo_code")%>"><%=ucase(rsliste_obspp("avo_libelle"))%></option>
<%
rsliste_obspp.movenext
loop
%>          
        </select></td>
        <td>
        <select name="obs6" id="obs6"> 
          <option value="99994">Mineurs</option>
<%
rsliste_obsmin.movefirst
do while not rsliste_obsmin.eof%>               
          <option value="<%=rsliste_obsmin("avo_code")%>"><%=ucase(rsliste_obsmin("avo_libelle"))%></option>
<%
rsliste_obsmin.movenext
loop
%>          
        </select></td>
        <td>
        <select name="obs7" id="obs7"> 
          <option value="99994">Etrangers</option>
<%
rsliste_obsetr.movefirst
do while not rsliste_obsetr.eof%>               
          <option value="<%=rsliste_obsetr("avo_code")%>"><%=ucase(rsliste_obsetr("avo_libelle"))%></option>
<%
rsliste_obsetr.movenext
loop
%>          
        </select></td>
      </tr>
            <tr><td><hr></td></tr>

      <tr>
        <td>Coordinateur 3 :</td>
        <td><select name="coord3" id="coord3">
        <option value="99993">en attente</option>
<%
rsliste_coord.movefirst
do while not rsliste_coord.eof
%>         
          <option value="<%=rsliste_coord("avo_code")%>"><%=ucase(rsliste_coord("avo_libelle"))%></option>
<%
rsliste_coord.movenext
loop
%>           
        </select></td>
      </tr>
<tr>
        <td>Observateurs :</td>
        <td>
        <select name="obs8" id="obs8"> 
          <option value="99994">Majeur1</option>
<%
rsliste_obspp.movefirst
do while not rsliste_obspp.eof%>               
          <option value="<%=rsliste_obspp("avo_code")%>"><%=ucase(rsliste_obspp("avo_libelle"))%></option>
<%
rsliste_obspp.movenext
loop
%>          
        </select></td>
        <td>
        <select name="obs9" id="obs9"> 
          <option value="99994">Majeur2</option>
<%
rsliste_obspp.movefirst
do while not rsliste_obspp.eof%>               
          <option value="<%=rsliste_obspp("avo_code")%>"><%=ucase(rsliste_obspp("avo_libelle"))%></option>
<%
rsliste_obspp.movenext
loop
%>          
        </select></td>
        <td>
        <select name="obs10" id="obs10"> 
          <option value="99994">Mineurs</option>
<%
rsliste_obsmin.movefirst
do while not rsliste_obsmin.eof%>               
          <option value="<%=rsliste_obsmin("avo_code")%>"><%=ucase(rsliste_obsmin("avo_libelle"))%></option>
<%
rsliste_obsmin.movenext
loop
%>          
        </select></td>
        <td>
        <select name="obs11" id="obs11"> 
          <option value="99994">Etrangers</option>
<%
rsliste_obsetr.movefirst
do while not rsliste_obsetr.eof%>               
          <option value="<%=rsliste_obsetr("avo_code")%>"><%=ucase(rsliste_obsetr("avo_libelle"))%></option>
<%
rsliste_obsetr.movenext
loop
%>          
        </select></td>
      </tr>
      <tr>
        <td width="420">&nbsp;</td>
        <td width="380"><input type="submit" name="button2" id="button2" value="G&eacute;n&eacute;rer" /></td>
      </tr>
    </table>
  </form>
  <p>&nbsp;</p>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
