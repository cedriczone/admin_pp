<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
Response.Buffer=false
Server.ScriptTimeout=300

Function in_array(element, arr)
  in_array = False
  For g=0 To Ubound(arr)
     If Trim(arr(g)) = Trim(element) Then
        in_array = True
        Exit Function      
     End If
  Next
End Function

'//////////////////////MISE A JOUR DES DISPOS PAR RAPPORT A LA LISTE DES INTERVENANTS GAV

SQLmajdispo="SELECT * from [dispos_gav]"
Set rsmajdispo=server.Createobject("adodb.recordset")
rsmajdispo.open SQLmajdispo,conn2,3,3
do while not rsmajdispo.eof

    SQLlisteintervenants="SELECT * from [Intervenants_GAV] where avo_code="&rsmajdispo("num_avo")
    Set rslisteintervenants=server.Createobject("adodb.recordset")
    rslisteintervenants.open SQLlisteintervenants,conn,3,3
    nbre_correspondance=rslisteintervenants.recordcount
        
    if nbre_correspondance=0 then
        SQLsuppr="DELETE * from [dispos_gav] where num_avo="&rsmajdispo("num_avo")
        Set suppr= Server.CreateObject("ADODB.RecordSet")
        suppr.open SQLsuppr,conn2,3,3
    else
        'on en profite pour mettre à jour le chevronné
        if rslisteintervenants("AVO_GAVCHE")=False then
            val_chev=0
        else
            val_chev=1
        end if
        SQLmodif_chev="UPDATE [dispos_gav] set chevronne="&val_chev&" WHERE num_avo="&rsmajdispo("num_avo")
        response.write("<br>"&SQLmodif_chev)
        Set modif_chev= Server.CreateObject("ADODB.RecordSet")
        modif_chev.open SQLmodif_chev,conn2,3,3
    end if

rsmajdispo.movenext
loop

'////////////////////////////////////////////////////////////////////////////////////////////

Zmois=request.form("mois")
if cint(Zmois)<month(date()) then
    Zannee=year(date())+1
else
    Zannee=year(date())
end if
if Zmois<month(date()) then Zannee=Zannee+1

'nombre de jours du mois
d1 = request.form("date1")
d2 = request.form("date2")
datecours=d1
joursmois = datediff("d",d1,d2)
joursmois = joursmois+1
premier_jour = weekday(d1,2)
if premier_jour<>2 and premier_jour<>4 and premier_jour<>6 then
    response.write("<br /><strong>Ce n'est pas un jour de début de module</strong>")
    response.write("<br /><a href='gene_planning_gav.asp'>retour a l''admin</a>")
    Response.End
end if


'function DaysInMonth(mois,an)
'   dim d1,d2
'   d1 = dateserial(an,mois,1)
'   d2 = dateserial(an,mois+1,1)
'   DaysInMonth = datediff("d",d1,d2)
'end function
                       
'joursmois = DaysInMonth(Zmois,Zannee)
'date_first=dateserial(Zannee,Zmois,1)
'premier_jour = weekday(date_first,2)

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
dim tab_inters(9)
                       
for i=1 to joursmois
    
    if k>0 then datecours=dateadd("d",k,d1)

    k=k+1
    l=l+1

    'le jour en cours dans la boucle
    dateencours=month(datecours)&"/"&day(datecours)&"/"&year(datecours)
    jourencours=Weekday(datecours,2)
    jourencours2=Weekday(datecours,2)
    
    nb_boucle=4
    jour_chev=0
    
    Select Case jourencours
                Case 1
                    jour="lun"
                    nb_boucle=6
                Case 2
                    jour="mar"
                    nb_boucle=6
                Case 3
                    jour="mer"
                    nb_boucle=6
                Case 4
                    jour="jeu"
                Case 5
                    jour="ven"
                Case 6
                    jour="sam"
                Case 7
                    jour="dim"
            end select
    
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
                Case 1
                    jour_chev=1
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
    response.write("<br>j= "&j)
    
    SQLdispo="SELECT * from [dispos_gav] WHERE ("&jour&"=1) AND (num_avo<>"&Zcoord&") order by cpt, avo_nom"
    response.write("<br>"&SQLdispo)
    Set rsdispo=Server.Createobject("adodb.recordset")
    rsdispo.open SQLdispo,conn2,3,3
    nb_dispo=rsdispo.recordcount
    if nb_dispo>0 then
        rsdispo.movefirst
        
        response.write("<br>Pour ce jour là nous avons de dispo:")
        do while not rsdispo.eof
            response.write("<br>"&rsdispo("avo_nom"))
            rsdispo.movenext
        loop
        
        rsdispo.movefirst
    else
        response.write("personne n'est dispo ce jour là !")
        exit for
    end if
                           
    if jour_chev=1 then
        SQLdispo_chev="SELECT * from [dispos_gav] WHERE ("&jour&"=1) AND (chevronne=1) AND (num_avo<>"&Zcoord&") order by cpt, avo_nom"
        Set rsdispo_chev=Server.Createobject("adodb.recordset")
        rsdispo_chev.open SQLdispo_chev,conn2,3,3
        nb_dispo_chev=rsdispo_chev.recordcount
    end if

    '///////// SELECTION DES INTERVENANTS DU JOUR
    
    ' POUR LE CHEVRONNE ON PASSE JUSTE AU SUIVANT
    if jour_chev=1 then
    response.write("<br />youpi cest un jour avec un chevronne")
        if nb_dispo_chev>0 then
            if rsdispo_chev.eof then
                rsdispo_chev.movefirst
            else
                rsdispo_chev.movenext
                 if rsdispo_chev.eof then rsdispo_chev.movefirst
            end if
            tab_inters(0)=rsdispo_chev("num_avo")
            response.write("<br />le chevronne est: "&tab_inters(0))
        end if
    end if
    
    'ON INITIALISE LES DEUX DERNIERS (MARDI ET JEUDI UNIQUEMENT) A ZERO
    tab_inters(5)=0
    tab_inters(6)=0
    
    'POUR LES AUTRES INTERVENANTS ON FAIT UN BOUCLE DE 4 OU 6 POUR MARDI ET JEUDI
    if (nb_dispo_chev=0) OR (jour_chev=0) then
        deb=0
    else
        deb=1
    end if
    
    response.write("<br>début de boucle à "&deb&" et fin à "&nb_boucle)
    
    for m=deb to nb_boucle
    
    'on vérifie que l'intervenant n'est pas déjà dans le tableau
    If in_array(rsdispo("num_avo"), tab_inters) then
        response.write("<br>Le numéro "&rsdispo("num_avo")&" etait deja present")
        m=m-1
    else
        response.write("<br>Je choisis: "&rsdispo("num_avo"))
        tab_inters(m)=rsdispo("num_avo")
        
        ' MISE A JOUR DU COMPTEUR
        cpt=rsdispo("cpt")
        cpt=cpt+1
        SQLmodifcpt="UPDATE [dispos_gav] set cpt="&cpt&" WHERE num_avo="&rsdispo("num_avo")
        response.write("<br>"&SQLmodifcpt)
        Set modifcpt= Server.CreateObject("ADODB.RecordSet")
        modifcpt.open SQLmodifcpt,conn2
    end if
    
        rsdispo.movenext
        if rsdispo.eof then rsdispo.movefirst
    next
    
    if jour_chev=1 then

        cpt2=rsdispo_chev("cpt")
        cpt2=cpt2+1
        SQLmodifcpt_chev="UPDATE [dispos_gav] set cpt="&cpt2&" WHERE num_avo="&rsdispo_chev("num_avo")
        Set modifcpt_chev= Server.CreateObject("ADODB.RecordSet")
        modifcpt_chev.open SQLmodifcpt_chev,conn2
        '////////////// fin enregistrement DERNIER CHEVRONNE
    end if
    
    '/////////////////////////////////////////////////////////////
    '///////// INSERTION DANS LA BASE DE DONNEES//////////////////
    '/////////////////////////////////////////////////////////////
    SQLaddgav="Insert Into [planning_gav](date_gav,num_coordinateur,num_inter_chev,num_inter1,num_inter2,num_inter3,num_inter4,num_inter5,num_inter6) Values(#"&dateencours&"#,"&Zcoord&","&tab_inters(0)&","&tab_inters(1)&","&tab_inters(2)&","&tab_inters(3)&","&tab_inters(4)&","&tab_inters(5)&","&tab_inters(6)&")"
    response.write("<br>"&SQLaddgav)
    Set saisiegav= Server.CreateObject("ADODB.RecordSet")
    saisiegav.open SQLaddgav,conn2
    
    erase tab_inters
    
    '///////////////////
    'FIN DE BOUCLE
    '///////////////////

next

conn.close
Set conn=nothing 
conn2.close
Set conn2=nothing

response.write("<a href='gene_planning_gav.asp'>retour a l''admin</a>")
%>