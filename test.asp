<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%

SQLverif1="SELECT * from [planning_gav] WHERE (date_gav>=#01/07/2010#) AND ((num_inter_jour1=1322) OR (num_inter_jour2=1322) OR (num_inter_nuit1=1322) OR (num_inter_nuit2=1322))"
Set rsverif1=server.Createobject("adodb.recordset")
rsverif1.open SQLverif1,conn2,3,3
nbre_verif1=rsverif1.recordcount

if nbre_verif1>0 then
rsverif1.movefirst
do while not rsverif1.eof

response.write(rsverif1("num_ligne")&" ---- "&rsverif1("date_gav")&"<br>")

rsverif1.movenext
loop

end if
%>