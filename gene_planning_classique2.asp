<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
'Err.Clear
'On Error Resume Next

Response.Buffer=false
Server.ScriptTimeout=999

Function in_array(element, arr)
response.write("element:" &element)
  in_array = False
  For g=0 To Ubound(arr)
     If Trim(arr(g)) = Trim(element) Then
        in_array = True
        Exit Function      
     End If
  Next
End Function

auth_block=False

'////////////////////// RECEPTION DES DONNEES////////////////////////////
dim designes(30)

dim coord(9)
coord(0)=request.form("coord1")
coord(1)=request.form("coord2")
coord(2)=request.form("coord3")

Zobs0 = request.form("obs0")
Zobs1 = request.form("obs1")
Zobs2 = request.form("obs2")
Zobs3 = request.form("obs3")
Zobs4 = request.form("obs4")
Zobs5 = request.form("obs5")
Zobs6 = request.form("obs6")
Zobs7 = request.form("obs7")
Zobs8 = request.form("obs8")
Zobs9 = request.form("obs9")
Zobs10 = request.form("obs10")
Zobs11 = request.form("obs11")

for i=0 to 11
	designes(i)=Zobs&i
next

designes(12)=coord(0)
designes(13)=coord(1)
designes(14)=coord(2)


'////////////////////// CREATION DES LISTES /////////////////////////////
'////////////////////// INTERVENANTS ADULTES /////////////////////////

SQLlisteinter_pp="SELECT * from [Intervenants_PP] WHERE AVO_ADUCHE=0 ORDER BY AVO_LIBELLE"
Set rslisteinter_pp=server.Createobject("adodb.recordset")
rslisteinter_pp.open SQLlisteinter_pp,conn,3,3
rslisteinter_pp.movefirst

'////////////////////// INTERVENANTS ADULTES CHEVRONNES /////////////////////////

SQLlisteinter_pp_chev="SELECT * from [Intervenants_PP] WHERE AVO_ADUCHE=1 ORDER BY AVO_LIBELLE"
Set rslisteinter_pp_chev=server.Createobject("adodb.recordset")
rslisteinter_pp_chev.open SQLlisteinter_pp_chev,conn,3,3
rslisteinter_pp_chev.movefirst
        
'////////////////////// INTERVENANTS MINEURS /////////////////////////

SQLlisteinter_min="SELECT * from [Intervenants_MIN] WHERE AVO_MINCHE=0 ORDER BY AVO_LIBELLE"
Set rslisteinter_min=server.Createobject("adodb.recordset")
rslisteinter_min.open SQLlisteinter_min,conn,3,3
rslisteinter_min.movefirst

'////////////////////// INTERVENANTS MINEURS CHEVRONNES /////////////////////////

SQLlisteinter_min_chev="SELECT * from [Intervenants_min] WHERE AVO_MINCHE=1 ORDER BY AVO_LIBELLE"
Set rslisteinter_min_chev=server.Createobject("adodb.recordset")
rslisteinter_min_chev.open SQLlisteinter_min_chev,conn,3,3
rslisteinter_min_chev.movefirst

'////////////////////// INTERVENANTS ETRANGERS /////////////////////////

SQLlisteinter_etr="SELECT * from [Intervenants_ETR] WHERE AVO_ETRCHE=0 ORDER BY AVO_LIBELLE"
Set rslisteinter_etr=server.Createobject("adodb.recordset")
rslisteinter_etr.open SQLlisteinter_etr,conn,3,3
rslisteinter_etr.movefirst

'////////////////////// INTERVENANTS ETRANGERS CHEVRONNES /////////////////////////

SQLlisteinter_etr_chev="SELECT * from [Intervenants_etr] WHERE AVO_ETRCHE=1 ORDER BY AVO_LIBELLE"
Set rslisteinter_etr_chev=server.Createobject("adodb.recordset")
rslisteinter_etr_chev.open SQLlisteinter_etr_chev,conn,3,3
rslisteinter_etr_chev.movefirst


'////////////////////// INTERVENANTS HO /////////////////////////

SQLlisteinter_ho="SELECT * from [Intervenants_HOF] ORDER BY AVO_LIBELLE"
Set rslisteinter_ho=server.Createobject("adodb.recordset")
rslisteinter_ho.open SQLlisteinter_ho,conn,3,3
rslisteinter_ho.movefirst


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

' SI LA DATE DE DEBUT N'EST PAS UN DEBUT DE MODULE ON ARRETE LE TRAITEMENT
if (premier_jour<>2) and (premier_jour<>4) and (premier_jour<>6) then
    response.write("<br /><strong>Ce n'est pas un jour de début de module</strong>")
    response.write("<br /><a href='gene_planning_classique.asp'>retour a l''admin</a>")
    Response.End
end if

'//////////////////////

' SELECTION DES AVOCATS

j=0
k=0 ' INCREMENTE LES JOURS AU FUR ET A MESURE EN AJOUTANT K A LA DATE AVEC DATEADD
l=0 ' SUIVI DES MODULES

'/////////////////////////////////////
'DEBUT DE LA GROSSE BOUCLE
'/////////////////////////////////////

'//////////////////////////////////////////////////////////////////
' RECUPERATION DES DERNIERS SORTIS
'//////////////////////////////////////////////////////////////////
SQLlast="SELECT * FROM [params_pp] WHERE id_param=1"
Set rslast=server.Createobject("adodb.recordset")
rslast.open SQLlast,conn2,3,3
rslast.movefirst

'majeurs
dernier_majeur=rslast("dernier_maj")
if dernier_maj>0 then
	rslisteinter_pp.FindFirst("AVO_CODE="&dernier_maj)
	rslisteinter_pp.movenext
end if

'majeurs chevronnes
dernier_majeur_chev=rslast("dernier_maj_chev")
if dernier_maj>0 then
	rslisteinter_pp_chev.FindFirst("AVO_CODE="&dernier_majeur_chev)
	rslisteinter_pp_chev.movenext
end if

'mineurs
dernier_mineur=rslast("dernier_min")
if dernier_mineur>0 then
	rslisteinter_min.FindFirst("AVO_CODE="&dernier_mineur)
	rslisteinter_min.movenext
end if

'mineurs chevronnes
dernier_mineur_chev=rslast("dernier_min_chev")
if dernier_mineur_chev>0 then
	rslisteinter_min_chev.FindFirst("AVO_CODE="&dernier_mineur_chev)
	rslisteinter_min_chev.movenext
end if

'etr
dernier_etr=rslast("dernier_etr")
if dernier_etr>0 then
	rslisteinter_etr.FindFirst("AVO_CODE="&dernier_etr)
	rslisteinter_etr.movenext
end if

'etr chevronnes
dernier_etr_chev=rslast("dernier_etr_chev")
if dernier_etr_chev>0 then
	rslisteinter_etr_chev.FindFirst("AVO_CODE="&dernier_etr_chev)
	rslisteinter_etr_chev.movenext
end if

'ho
dernier_ho=rslast("dernier_ho")
if dernier_ho>0 then
	rslisteinter_ho.FindFirst("AVO_CODE="&dernier_ho)
	rslisteinter_ho.movenext
end if
                       
' on déclare un tableau d'inters
dim tab_inters(99)
dim tab_obs(99)

'///////////////////// DEBUT BOUCLE ////////////////////////////////////
for i=1 to joursmois ' début boucle !
    'response.write("<br>"&l)
    if k>0 then datecours=dateadd("d",k,d1)

    k=k+1 ' incrémente la date au fur et à mesure
    l=l+1 ' vérifie si on est à la fin d'un module

    'le jour en cours dans la boucle
    dateencours=month(datecours)&"/"&day(datecours)&"/"&year(datecours)
    jourencours=Weekday(datecours,2)
    jourencours2=Weekday(datecours,2)
	jourencours = int(jourencours)

    response.write("<br/>jour en cours: "&jourencours&"<br/>")
        
' Si le jour en cours n'est pas le 1er du module, il suffit de récupérer l'enregistrement d'avant et copier les données
	SQLlast_pp="SELECT TOP 1 * FROM [planning_pp] order by num_ligne DESC"
	Set rslast_pp=server.Createobject("adodb.recordset")
	rslast_pp.open SQLlast_pp,conn2,3,3
	nb_rslast = rslast_pp.recordcount

if (jourencours=1 OR jourencours=3 OR jourencours=5 OR jourencours=7) AND (nb_rslast>0) then
	response.write("<br/>je fais une copie de la ligne du dessus<br/>")

	rslast_pp.movefirst

	Zcoord = rslast_pp("num_coordinateur")
	tab_obs(0) = rslast_pp("majeur_obs1")
	tab_obs(1) = rslast_pp("majeur_obs2")
	tab_obs(2) = rslast_pp("mineur_obs")
	tab_obs(3) = rslast_pp("etr_obs")
	tab_inters(0) = rslast_pp("majeur1")
	tab_inters(1) = rslast_pp("majeur2")
	tab_inters(2) = rslast_pp("majeur3")
	tab_inters(3) = rslast_pp("majeur4")
	tab_inters(4) = rslast_pp("mineur1")
	tab_inters(5) = rslast_pp("mineur2")
	tab_inters(6) = rslast_pp("etr1")
	tab_inters(7) = rslast_pp("etr2")
	tab_inters(8) = rslast_pp("ho")
	
    if l>=7 then
		'response.write("<br>fin module")
		j=j+1
		if j=3 then j=0
		l=0
	end if

'SINON ON NE RECOPIE PAS LES DONNEES MAIS ON PASSE AUX SUIVANTS DANS LE TRAITEMENT
else

'/////////
'///////// Rotation des coordinateurs///////
'/////////
Select Case jourencours
	case 2,4,6
	j=j+1
	if l>=7 then
		'response.write("fin module")
		j=j+1
		if j>=3 then j=0
		l=0
	else
		if j>=3 then j=0
	end if
	if j>=3 then j=0
end select

if j=3 then j=0
Zcoord=coord(j)


'rotation des observateurs
Select Case j
	Case 0
		tab_obs(0)=Zobs0
		tab_obs(1)=Zobs1
		tab_obs(2)=Zobs2
		tab_obs(3)=Zobs3
	Case 1
		tab_obs(0)=Zobs4
		tab_obs(1)=Zobs5
		tab_obs(2)=Zobs6
		tab_obs(3)=Zobs7
	Case 2
		tab_obs(0)=Zobs8
		tab_obs(1)=Zobs9
		tab_obs(2)=Zobs10
		tab_obs(3)=Zobs11
End Select

' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////// SELECTION DES INTERVENANTS DU JOUR //////////////////////////////////////////////////////////
' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

' /////////////////////////////////////////////////////////////
' MAJEURS
' /////////////////////////////////////////////////////////////
for t=0 to 2
	if auth_block=TRUE then

		SQLpriorite_majeur = "SELECT * FROM [cpt_pp] WHERE type='majeur' ORDER BY cpt"
		Set rspriorite_majeur=server.Createobject("adodb.recordset")
		rspriorite_majeur.open SQLpriorite_majeur,conn2,3,3
		nbprio1 = rspriorite_majeur.recordcount

		if nbprio1>0 then

			rspriorite_majeur.movefirst

			' on compare si un avocat est à la bourre au niveau compteur
			Zcpt_majeur = 0
			SQLcpt_majeur = "SELECT * FROM [cpt_pp] WHERE num_avo = "&rslisteinter_pp("AVO_CODE")
			Set rscpt_majeur=server.Createobject("adodb.recordset")
			rscpt_majeur.open SQLcpt_majeur,conn2,3,3
			nb_cpt_majeur=rscpt_majeur.recordcount
			if nb_cpt_majeur>0 then Zcpt_majeur=int(rscpt_majeur("cpt"))

			if (Zcpt_majeur-int(rspriorite_majeur("cpt")))>2 then
				' s'il est à la bourre on prend celui qui a le moins de désignations
				tab_inters(t) = rspriorite_majeur("num_avo")
			else
				' sinon on continue la logique de la boucle et on prend celui en cours dans l'alphabétique
				if  in_array(rslisteinter_pp("AVO_CODE"),designes)=True then
					do while in_array(rslisteinter_pp("AVO_CODE"),designes)=True
						rslisteinter_pp.movenext
						if rslisteinter_pp.eof then rslisteinter_pp.movefirst
					loop
				end if

		tab_inters(t) = rslisteinter_pp("AVO_CODE")
		rslisteinter_pp.movenext
		if rslisteinter_pp.eof then rslisteinter_pp.movefirst
			end if

		else

			'on vérifie qu'il n'est pas déjà désigné
		if  in_array(rslisteinter_pp("AVO_CODE"),designes)=True then
			do while in_array(rslisteinter_pp("AVO_CODE"),designes)=True
				rslisteinter_pp.movenext
				if rslisteinter_pp.eof then rslisteinter_pp.movefirst
			loop
		end if

		tab_inters(t) = rslisteinter_pp("AVO_CODE")
		rslisteinter_pp.movenext
		if rslisteinter_pp.eof then rslisteinter_pp.movefirst

		end if

	'SI auth_block = FALSE
	else

		'on vérifie qu'il n'est pas déjà désigné
		if  in_array(rslisteinter_pp("AVO_CODE"),designes)=True then
			do while in_array(rslisteinter_pp("AVO_CODE"),designes)=True
				rslisteinter_pp.movenext
				if rslisteinter_pp.eof then rslisteinter_pp.movefirst
			loop
		end if

		tab_inters(t) = rslisteinter_pp("AVO_CODE")
		rslisteinter_pp.movenext
		if rslisteinter_pp.eof then rslisteinter_pp.movefirst

	END IF


next
designes(15)=tab_inters(0)
designes(16)=tab_inters(1)
designes(17)=tab_inters(2)

' /////////////////////////////////////////////////////////////
' MAJEURS CHEVRONNES
' /////////////////////////////////////////////////////////////

if auth_block=TRUE then

	SQLpriorite_majeur_chev = "SELECT * FROM [cpt_pp] WHERE type='majeur' ORDER BY cpt"
	Set rspriorite_majeur_chev=server.Createobject("adodb.recordset")
	rspriorite_majeur_chev.open SQLpriorite_majeur_chev,conn2,3,3
	nbprio2 = rspriorite_majeur_chev.recordcount

	if nbprio2>0 then

		rspriorite_majeur_chev.movefirst

		' on compare si un avocat est à la bourre au niveau compteur
		Zcpt_majeur_chev = 0
		SQLcpt_majeur_chev = "SELECT * FROM [cpt_pp] WHERE num_avo = "&rslisteinter_pp_chev("AVO_CODE")
		Set rscpt_majeur_chev=server.Createobject("adodb.recordset")
		rscpt_majeur_chev.open SQLcpt_majeur_chev,conn2,3,3
		
		if rscpt_majeur_chev.recordcount>0 then Zcpt_majeur_chev = int(rscpt_majeur_chev("cpt"))

		if (Zcpt_majeur_chev-int(rspriorite_majeur_chev("cpt")))>2 then
			' s'il est à la bourre on prend celui qui a le moins de désignations
			tab_inters(3) = rspriorite_majeur_chev("num_avo")
		else
			' sinon on continue la logique de la boucle et on prend celui en cours dans l'alphabétique
			if  in_array(rslisteinter_pp_chev("AVO_CODE"),designes)=True then
				do while in_array(rslisteinter_pp_chev("AVO_CODE"),designes)=True
					rslisteinter_pp_chev.movenext
					if rslisteinter_pp_chev.eof then rslisteinter_pp_chev.movefirst
				loop
			end if
			tab_inters(3) = rslisteinter_pp_chev("AVO_CODE")
			rslisteinter_pp_chev.movenext
			if rslisteinter_pp_chev.eof then rslisteinter_pp_chev.movefirst
		end if

	else

		if  in_array(rslisteinter_pp_chev("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_pp_chev("AVO_CODE"),designes)=True
			rslisteinter_pp_chev.movenext
			if rslisteinter_pp_chev.eof then rslisteinter_pp_chev.movefirst
		loop
	end if
	tab_inters(3) = rslisteinter_pp_chev("AVO_CODE")
	rslisteinter_pp_chev.movenext
	if rslisteinter_pp_chev.eof then rslisteinter_pp_chev.movefirst

	end if

else
	if  in_array(rslisteinter_pp_chev("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_pp_chev("AVO_CODE"),designes)=True
			rslisteinter_pp_chev.movenext
			if rslisteinter_pp_chev.eof then rslisteinter_pp_chev.movefirst
		loop
	end if
	tab_inters(3) = rslisteinter_pp_chev("AVO_CODE")
	rslisteinter_pp_chev.movenext
	if rslisteinter_pp_chev.eof then rslisteinter_pp_chev.movefirst

END IF
		
designes(18)=tab_inters(3)


' /////////////////////////////////////////////////////////////
' MINEURS
' /////////////////////////////////////////////////////////////

if auth_block=TRUE then

	SQLpriorite_mineur = "SELECT * FROM [cpt_pp] WHERE type='mineur' ORDER BY cpt"
	Set rspriorite_mineur=server.Createobject("adodb.recordset")
	rspriorite_mineur.open SQLpriorite_mineur,conn2,3,3
	nbprio3 = rspriorite_mineur.recordcount
	if nbprio3>0 then

		rspriorite_mineur.movefirst

		' on compare si un avocat est à la bourre au niveau compteur
		Zcpt_mineur = 0
		SQLcpt_mineur = "SELECT * FROM [cpt_pp] WHERE num_avo = "&rslisteinter_min("AVO_CODE")
		Set rscpt_mineur=server.Createobject("adodb.recordset")
		rscpt_mineur.open SQLcpt_mineur,conn2,3,3
		if rscpt_mineur.recordcount>0 then Zcpt_mineur = int(rscpt_mineur("cpt"))

		if (Zcpt_mineur-int(rspriorite_mineur("cpt")))>2 then
			' s'il est à la bourre on prend celui qui a le moins de désignations
			tab_inters(4) = rspriorite_mineur("num_avo")
		else
			' sinon on continue la logique de la boucle et on prend celui en cours dans l'alphabétique
			if  in_array(rslisteinter_min("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_min("AVO_CODE"),designes)=True
			rslisteinter_min.movenext
			if rslisteinter_min.eof then rslisteinter_min.movefirst
		loop
	end if
	tab_inters(4) = rslisteinter_min("AVO_CODE")
	rslisteinter_min.movenext
	if rslisteinter_min.eof then rslisteinter_min.movefirst
		end if

	else

		if  in_array(rslisteinter_min("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_min("AVO_CODE"),designes)=True
			rslisteinter_min.movenext
			if rslisteinter_min.eof then rslisteinter_min.movefirst
		loop
	end if
	tab_inters(4) = rslisteinter_min("AVO_CODE")
	rslisteinter_min.movenext
	if rslisteinter_min.eof then rslisteinter_min.movefirst

	end if

else

	if  in_array(rslisteinter_min("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_min("AVO_CODE"),designes)=True
			rslisteinter_min.movenext
			if rslisteinter_min.eof then rslisteinter_min.movefirst
		loop
	end if
	tab_inters(4) = rslisteinter_min("AVO_CODE")
	rslisteinter_min.movenext
	if rslisteinter_min.eof then rslisteinter_min.movefirst

END IF

designes(19)=tab_inters(4)

' /////////////////////////////////////////////////////////////
' MINEURS CHEVRONNES
' /////////////////////////////////////////////////////////////
  
IF auth_block=TRUE then

	SQLpriorite_mineur_chev = "SELECT * FROM [cpt_pp] WHERE type='mineur' ORDER BY cpt"
	Set rspriorite_mineur_chev=server.Createobject("adodb.recordset")
	rspriorite_mineur_chev.open SQLpriorite_mineur_chev,conn2,3,3
	nbprio4 = rspriorite_mineur_chev.recordcount
	if nbprio4>0 then

		rspriorite_mineur_chev.movefirst

		' on compare si un avocat est à la bourre au niveau compteur
		Zcpt_mineur_chev = 0
		SQLcpt_mineur_chev = "SELECT * FROM [cpt_pp] WHERE num_avo = "&rslisteinter_min_chev("AVO_CODE")
		Set rscpt_mineur_chev=server.Createobject("adodb.recordset")
		rscpt_mineur_chev.open SQLcpt_mineur_chev,conn2,3,3
		if rscpt_mineur_chev.recordcount>0 then Zcpt_mineur_chev = int(rscpt_mineur_chev("cpt"))

		if (Zcpt_mineur_chev-int(rspriorite_mineur_chev("cpt")))>2 then
			' s'il est à la bourre on prend celui qui a le moins de désignations
			tab_inters(5) = rspriorite_mineur_chev("num_avo")
		else
			' sinon on continue la logique de la boucle et on prend celui en cours dans l'alphabétique
			if  in_array(rslisteinter_min_chev("AVO_CODE"),designes)=True then
			do while in_array(rslisteinter_min_chev("AVO_CODE"),designes)=True
				rslisteinter_min_chev.movenext
				if rslisteinter_min_chev.eof then rslisteinter_min_chev.movefirst
			loop
		end if
		tab_inters(5) = rslisteinter_min_chev("AVO_CODE")
		rslisteinter_min_chev.movenext
		if rslisteinter_min_chev.eof then rslisteinter_min_chev.movefirst
		end if

	else

		if  in_array(rslisteinter_min_chev("AVO_CODE"),designes)=True then
			do while in_array(rslisteinter_min_chev("AVO_CODE"),designes)=True
				rslisteinter_min_chev.movenext
				if rslisteinter_min_chev.eof then rslisteinter_min_chev.movefirst
			loop
		end if
		tab_inters(5) = rslisteinter_min_chev("AVO_CODE")
		rslisteinter_min_chev.movenext
		if rslisteinter_min_chev.eof then rslisteinter_min_chev.movefirst

	end if

else
	if  in_array(rslisteinter_min_chev("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_min_chev("AVO_CODE"),designes)=True
			rslisteinter_min_chev.movenext
			if rslisteinter_min_chev.eof then rslisteinter_min_chev.movefirst
		loop
	end if
	tab_inters(5) = rslisteinter_min_chev("AVO_CODE")
	rslisteinter_min_chev.movenext
	if rslisteinter_min_chev.eof then rslisteinter_min_chev.movefirst

END IF
designes(20)=tab_inters(5)

	
' /////////////////////////////////////////////////////////////
' ETRANGERS
' /////////////////////////////////////////////////////////////

IF auth_block=TRUE then

	SQLpriorite_etr = "SELECT * FROM [cpt_pp] WHERE type='etr' ORDER BY cpt"
	Set rspriorite_etr=server.Createobject("adodb.recordset")
	rspriorite_etr.open SQLpriorite_etr,conn2,3,3
	nbprio5 = rspriorite_etr.recordcount
	if nbprio5>0 then

		rspriorite_etr.movefirst

		' on compare si un avocat est à la bourre au niveau compteur
		Zcpt_etr = 0
		SQLcpt_etr = "SELECT * FROM [cpt_pp] WHERE num_avo = "&rslisteinter_etr("AVO_CODE")
		Set rscpt_etr=server.Createobject("adodb.recordset")
		rscpt_etr.open SQLcpt_etr,conn2,3,3
		if rscpt_etr.recordcount>0 then Zcpt_etr = int(rscpt_etr("cpt"))

		if (Zcpt_etr-int(rspriorite_etr("cpt")))>2 then
			' s'il est à la bourre on prend celui qui a le moins de désignations
			tab_inters(6) = rspriorite_etr("num_avo")
		else
			' sinon on continue la logique de la boucle et on prend celui en cours dans l'alphabétique
			if  in_array(rslisteinter_etr("AVO_CODE"),designes)=True then
				do while in_array(rslisteinter_etr("AVO_CODE"),designes)=True
					rslisteinter_etr.movenext
					if rslisteinter_etr.eof then rslisteinter_etr.movefirst
				loop
			end if
			tab_inters(6) = rslisteinter_etr("AVO_CODE")
			rslisteinter_etr.movenext
			if rslisteinter_etr.eof then rslisteinter_etr.movefirst
		end if

	else

		if  in_array(rslisteinter_etr("AVO_CODE"),designes)=True then
			do while in_array(rslisteinter_etr("AVO_CODE"),designes)=True
				rslisteinter_etr.movenext
				if rslisteinter_etr.eof then rslisteinter_etr.movefirst
			loop
		end if
		tab_inters(6) = rslisteinter_etr("AVO_CODE")
		rslisteinter_etr.movenext
		if rslisteinter_etr.eof then rslisteinter_etr.movefirst

	end if

else

	if  in_array(rslisteinter_etr("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_etr("AVO_CODE"),designes)=True
			rslisteinter_etr.movenext
			if rslisteinter_etr.eof then rslisteinter_etr.movefirst
		loop
	end if
	tab_inters(6) = rslisteinter_etr("AVO_CODE")
	rslisteinter_etr.movenext
	if rslisteinter_etr.eof then rslisteinter_etr.movefirst

END IF

designes(21)=tab_inters(6)


' /////////////////////////////////////////////////////////////
' ETRANGERS CHEVRONNES
' /////////////////////////////////////////////////////////////

if auth_block=TRUE then

	SQLpriorite_etr_chev = "SELECT * FROM [cpt_pp] WHERE type='etr' ORDER BY cpt"
	Set rspriorite_etr_chev=server.Createobject("adodb.recordset")
	rspriorite_etr_chev.open SQLpriorite_etr_chev,conn2,3,3
	nbprio5 = rspriorite_etr_chev.recordcount
	if nbprio5>0 then

		rspriorite_etr_chev.movefirst

		' on compare si un avocat est à la bourre au niveau compteur
		Zcpt_etr_chev = 0
		SQLcpt_etr_chev = "SELECT * FROM [cpt_pp] WHERE num_avo = "&rslisteinter_etr_chev("AVO_CODE")
		Set rscpt_etr_chev=server.Createobject("adodb.recordset")
		rscpt_etr_chev.open SQLcpt_etr_chev,conn2,3,3
		if rscpt_etr_chev.recordcount>0 then Zcpt_etr_chev = int(rscpt_etr_chev("cpt"))

		if (Zcpt_etr_chev-int(rspriorite_etr_chev("cpt")))>2 then
			' s'il est à la bourre on prend celui qui a le moins de désignations
			tab_inters(7) = rspriorite_etr_chev("num_avo")
		else
			' sinon on continue la logique de la boucle et on prend celui en cours dans l'alphabétique
			if  in_array(rslisteinter_etr_chev("AVO_CODE"),designes)=True then
				do while in_array(rslisteinter_etr_chev("AVO_CODE"),designes)=True
					rslisteinter_etr_chev.movenext
					if rslisteinter_etr_chev.eof then rslisteinter_etr_chev.movefirst
				loop
			end if
			tab_inters(7) = rslisteinter_etr_chev("AVO_CODE")
			rslisteinter_etr_chev.movenext
			if rslisteinter_etr_chev.eof then rslisteinter_etr_chev.movefirst
		end if

	else

			if  in_array(rslisteinter_etr_chev("AVO_CODE"),designes)=True then
				do while in_array(rslisteinter_etr_chev("AVO_CODE"),designes)=True
					rslisteinter_etr_chev.movenext
					if rslisteinter_etr_chev.eof then rslisteinter_etr_chev.movefirst
				loop
			end if
			tab_inters(7) = rslisteinter_etr_chev("AVO_CODE")
			rslisteinter_etr_chev.movenext
			if rslisteinter_etr_chev.eof then rslisteinter_etr_chev.movefirst

end if

else

	if  in_array(rslisteinter_etr_chev("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_etr_chev("AVO_CODE"),designes)=True
			rslisteinter_etr_chev.movenext
			if rslisteinter_etr_chev.eof then rslisteinter_etr_chev.movefirst
		loop
	end if
	tab_inters(7) = rslisteinter_etr_chev("AVO_CODE")
	rslisteinter_etr_chev.movenext
	if rslisteinter_etr_chev.eof then rslisteinter_etr_chev.movefirst

END IF

designes(22)=tab_inters(7)

' /////////////////////////////////////////////////////////////
' HO
' /////////////////////////////////////////////////////////////

IF auth_block=TRUE then

	SQLpriorite_ho = "SELECT * FROM [cpt_pp] WHERE type='ho' ORDER BY cpt"
	Set rspriorite_ho=server.Createobject("adodb.recordset")
	rspriorite_ho.open SQLpriorite_ho,conn2,3,3
	nbprio6 = rspriorite_ho.recordcount
	if nbprio6>0 then

		rspriorite_etr.movefirst

		' on compare si un avocat est à la bourre au niveau compteur
		Zcpt_ho = 0
		SQLcpt_ho = "SELECT * FROM [cpt_pp] WHERE num_avo = "&rslisteinter_ho("AVO_CODE")
		Set rscpt_ho=server.Createobject("adodb.recordset")
		rscpt_ho.open SQLcpt_ho,conn2,3,3
		if rscpt_ho.recordcount>0 then Zcpt_ho = int(rscpt_ho("cpt"))

		if (Zcpt_etr-int(rspriorite_etr("cpt")))>2 then
			' s'il est à la bourre on prend celui qui a le moins de désignations
			tab_inters(8) = rspriorite_etr("num_avo")
		else
			' sinon on continue la logique de la boucle et on prend celui en cours dans l'alphabétique
			if  in_array(rslisteinter_ho("AVO_CODE"),designes)=True then
				do while in_array(rslisteinter_ho("AVO_CODE"),designes)=True
					rslisteinter_ho.movenext
					if rslisteinter_ho.eof then rslisteinter_ho.movefirst
				loop
			end if
			tab_inters(8) = rslisteinter_ho("AVO_CODE")
			rslisteinter_ho.movenext
			if rslisteinter_ho.eof then rslisteinter_ho.movefirst
		end if

	else

		if  in_array(rslisteinter_ho("AVO_CODE"),designes)=True then
			do while in_array(rslisteinter_ho("AVO_CODE"),designes)=True
				rslisteinter_ho.movenext
				if rslisteinter_ho.eof then rslisteinter_ho.movefirst
			loop
		end if
		tab_inters(8) = rslisteinter_ho("AVO_CODE")
		rslisteinter_ho.movenext
		if rslisteinter_ho.eof then rslisteinter_ho.movefirst

	end if

else
	if  in_array(rslisteinter_ho("AVO_CODE"),designes)=True then
		do while in_array(rslisteinter_ho("AVO_CODE"),designes)=True
			rslisteinter_ho.movenext
			if rslisteinter_ho.eof then rslisteinter_ho.movefirst
		loop
	end if
	tab_inters(8) = rslisteinter_ho("AVO_CODE")
	rslisteinter_ho.movenext
	if rslisteinter_ho.eof then rslisteinter_ho.movefirst

END IF

designes(23)=tab_inters(8)

End if 'Fin de la condition pour savoir si on recopie la ligne précédente ou pas
   
    '/////////////////////////////////////////////////////////////
    '///////// INSERTION DANS LA BASE DE DONNEES//////////////////
    '/////////////////////////////////////////////////////////////

'response.write("le coord est: "&Zcoord)

SQLaddpp="Insert Into [planning_pp](date_pp,num_coordinateur,majeur1,majeur2,majeur3,majeur4,majeur_obs1,majeur_obs2,mineur1,mineur2,mineur_obs,etr1,etr2,etr_obs,ho) Values(#"&dateencours&"#,"&Zcoord&","&tab_inters(0)&","&tab_inters(1)&","&tab_inters(2)&","&tab_inters(3)&","&tab_obs(0)&","&tab_obs(1)&","&tab_inters(4)&","&tab_inters(5)&","&tab_obs(2)&","&tab_inters(6)&","&tab_inters(7)&","&tab_obs(3)&","&tab_inters(8)&")"
response.write("<br>"&SQLaddpp)
Set saisiepp= Server.CreateObject("ADODB.RecordSet")
saisiepp.open SQLaddpp,conn2,3,3


    '/////////////////////////////////////////////////////////////
    '///////// MODIF COMPTEURS//////////////////
    '/////////////////////////////////////////////////////////////

SQLverifcptcoord = "SELECT * FROM [cpt_pp] WHERE num_avo="&Zcoord
Set rsverifcptcoord= Server.CreateObject("ADODB.RecordSet")
rsverifcptcoord.open SQLverifcptcoord,conn2,3,3
nb_trouve = rsverifcptcoord.recordcount
if nb_trouve=1 then
	nbcpt = rsverifcptcoord("cpt")+1
	SQLmodifcpt="UPDATE [cpt_pp] set cpt="&nbcpt&" WHERE num_avo="&Zcoord
	Set modifcpt= Server.CreateObject("ADODB.RecordSet")
	modifcpt.open SQLmodifcpt,conn2,3,3
else
	SQLaddcoord="Insert Into [cpt_pp](num_avo,cpt) Values("&Zcoord&",1)"
	Set addcoord= Server.CreateObject("ADODB.RecordSet")
	addcoord.open SQLaddcoord,conn2,3,3
end if

for s=0 to 8

	Select Case s
		Case 0,1,2,3
			Ztype="majeur"
		Case 4,5
			Ztype="mineur"
		Case 6,7
			Ztype="etr"
		Case 8
			Ztype="ho"
	End Select

	SQLverifcpt = "SELECT * FROM [cpt_pp] WHERE num_avo="&tab_inters(s)
	Set rsverifcpt= Server.CreateObject("ADODB.RecordSet")
	rsverifcpt.open SQLverifcpt,conn2,3,3
	nb_trouve = rsverifcpt.recordcount
	if nb_trouve=1 then
		nbcpt = rsverifcpt("cpt")+1
		SQLmodifcpt="UPDATE [cpt_pp] set cpt="&nbcpt&" WHERE num_avo="&tab_inters(s)
		Set modifcpt= Server.CreateObject("ADODB.RecordSet")
		modifcpt.open SQLmodifcpt,conn2,3,3
	else
		SQLadd1="Insert Into [cpt_pp](num_avo,cpt,type) Values("&tab_inters(s)&",1,'"&Ztype&"')"
		Set add1= Server.CreateObject("ADODB.RecordSet")
		add1.open SQLadd1,conn2,3,3
	end if
next

for t=0 to 2
	SQLverifcpt2 = "SELECT * FROM [cpt_pp] WHERE num_avo="&tab_obs(t)
	Set rsverifcpt2= Server.CreateObject("ADODB.RecordSet")
	rsverifcpt2.open SQLverifcpt2,conn2,3,3
	nb_trouve2 = rsverifcpt2.recordcount
	if nb_trouve2=1 then
		nbcpt = rsverifcpt2("cpt")+1
		SQLmodifcpt2="UPDATE [cpt_pp] set cpt="&nbcpt&" WHERE num_avo="&tab_obs(t)
		Set modifcpt2= Server.CreateObject("ADODB.RecordSet")
		modifcpt2.open SQLmodifcpt2,conn2,3,3
	else
		SQLadd2="Insert Into [cpt_pp](num_avo,cpt) Values("&tab_obs(t)&",1)"
		Set add2= Server.CreateObject("ADODB.RecordSet")
		add2.open SQLadd2,conn2,3,3
	end if
next


erase tab_inters
    
'///////////////////
'FIN DE BOUCLE
'///////////////////

next

conn.close
Set conn=nothing 
conn2.close
Set conn2=nothing

response.write("<a href='gene_planning_classique.asp'>retour a l''admin</a>")

If Err.Number <> 0 Then
  
  Response.Write ("<br><br>" & Err.Description& "<br><br>" & Err.Source & "<br><br>" & Err.Number)
   

  Response.End
 
End If
%>