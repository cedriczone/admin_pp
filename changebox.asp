<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zzid=request.querystring("valeur")
Zaction=request.QueryString("action")
if Zzid<>"" then
lng=len(Zzid)-3
Zid=right(Zzid,lng)
periode=left(Zzid,3)

SQLrecherche="SELECT * from [dispos_gav] where num_avo="&Zid
Set rsrecherche=server.Createobject("adodb.recordset")
rsrecherche.open SQLrecherche,conn2,3,3
nbre_rech=rsrecherche.recordcount
if nbre_rech<1 then

    SQLavo="SELECT * FROM [Avocats_PP] WHERE avo_code="&Zid
    Set rsavo=server.Createobject("adodb.recordset")
    rsavo.open SQLavo,conn,3,3
    
    SQLadd="Insert Into [dispos_gav](num_avo,avo_nom) Values("&Zid&",'"&replace(rsavo("avo_libelle"),"'","''")&"')"
    Set saisie= Server.CreateObject("ADODB.RecordSet")
    saisie.open SQLadd,conn2

end if

SQLmodif="UPDATE [dispos_gav] set "&periode&"="&Zaction&" WHERE num_avo="&Zid
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2
end if
%>