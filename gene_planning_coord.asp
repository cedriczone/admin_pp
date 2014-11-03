<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
Response.Buffer=false
Server.ScriptTimeout=300

Zmois=request.form("mois")
if cint(Zmois)<month(date()) then
    Zannee=year(date())+1
else
    Zannee=year(date())
end if
if Zmois<month(date()) then Zannee=Zannee+1

'nombre de jours du mois
function DaysInMonth(mois,an)
   dim d1,d2
   d1 = dateserial(an,mois,1)
   d2 = dateserial(an,mois+1,1)
   DaysInMonth = datediff("d",d1,d2)
end function
                       
joursmois = DaysInMonth(Zmois,Zannee)
date_first=dateserial(Zannee,Zmois,1)
premier_jour = weekday(date_first,2)

dim coord(9)
coord(0)=request.form("coord1")
coord(1)=request.form("coord2")
coord(2)=request.form("coord3")

'//////////////////////

' SELECTION DES AVOCATS

j=-1 'NB DU COORDINATEUR
k=0 ' INCREMENTE LES JOURS AU FUR ET A MESURE EN AJOUTANT K A LA DATE AVEC DATEADD
l=0

'/////////////////////////////////////
'DEBUT DE LA GROSSE BOUCLE
'/////////////////////////////////////
                       
' on déclare un tableau d'inters
                       
for i=1 to joursmois
    datecours=dateadd("d",k,date_first)
    k=k+1
    l=l+1
    
    'le jour en cours dans la boucle
    dateencours=month(datecours)&"/"&day(datecours)&"/"&year(datecours)
    jourencours=Weekday(datecours,2)
    jourencours2=Weekday(datecours,2)
      
    if (i=1) AND (jourencours=1 OR jourencours=3 OR jourencours=5 OR jourencours=7) then
    response.write("<br>je vais chercher celui du jour d avant")
        SQLplanning_gav="SELECT TOP 1 * FROM [planning_gav] order by num_ligne DESC"
        Set rsplanning_gav=server.Createobject("adodb.recordset")
        rsplanning_gav.open SQLplanning_gav,conn2,3,3
        rsplanning_gav.movefirst
        Zcoord=rsplanning_gav("num_coordinateur")
        l=0
    
    else
    
        if (i=2) AND (jourencours=1) then
            response.write("<br>je vais chercher celui du jour d avant")
            SQLplanning_gav="SELECT TOP 1 * FROM [planning_gav] order by num_ligne DESC"
            Set rsplanning_gav=server.Createobject("adodb.recordset")
            rsplanning_gav.open SQLplanning_gav,conn2,3,3
            rsplanning_gav.movefirst
            Zcoord=rsplanning_gav("num_coordinateur")
            l=0
            
        else
        
            Select Case jourencours
                Case 2
                    jour_chev=1
                    j=j+1
                    if l>=7 then
                        response.write("<br>j ai saute un coord car l>=7")
                        j=j+1
                        if j=3 then j=0
                        if j=4 then j=1
                        l=0
                    end if
                Case 3
                    jour_chev=1
                Case 4
                    jour_chev=1
                    j=j+1
                    if j=3 then j=0
                    if j=4 then j=1
                    if l>=7 then
                        response.write("<br>j ai saute un coord car l>=7")
                        j=j+1
                        l=0
                    end if
                Case 6
                    jour="sam"
                    j=j+1
                    if j=3 then j=0
                    if j=4 then j=1
                    if l>=7 then
                        response.write("<br>j ai saute un coord car l>=7")
                        j=j+1
                        l=0
                    end if
            end select
        
            if j=3 then
                j=0
                response.write("<br>j=3 donc j=0")
            end if
            if j=-1 then j=0
            Zcoord=coord(j) 
            
        end if
    
    end if
    
    '/////////////////////////////////////////////////////////////
    '///////// INSERTION DANS LA BASE DE DONNEES//////////////////
    '/////////////////////////////////////////////////////////////
date_modif=month(datecours)&"/"&day(datecours)&"/"&year(datecours)

    SQLaddgav="UPDATE [planning_gav] set num_coordinateur="&Zcoord&" WHERE date_gav=#"&date_modif&"#"
    response.write("<br>"&SQLaddgav)
    Set saisiegav= Server.CreateObject("ADODB.RecordSet")
    saisiegav.open SQLaddgav,conn2
    
    '///////////////////
    'FIN DE BOUCLE
    '///////////////////
next

conn.close
Set conn=nothing 
conn2.close
Set conn2=nothing
%>