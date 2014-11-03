<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
On Error resume Next

demain=dateAdd("d",1,date())
lendemain=month(demain)&"/"&day(demain)&"/"&year(demain)
strdate = strdate & weekdayname(weekday(demain))
strdate = strdate & " "
If day(demain) = 1 Then
   strdate = strdate & "1er"
Else
   strdate = strdate & day(demain)
End If
strdate = strdate & " "
strdate = strdate & monthname(month(demain))
strdate = strdate & " "
strdate = strdate & year(demain)

if day(demain)<10 then
jourdemain="0"&day(demain)
else
jourdemain=day(demain)
end if

if month(demain)<10 then
moisdemain="0"&month(demain)
else
moisdemain=month(demain)
end if

SQLGAV="SELECT * from [planning_gav] where date_gav=#"&lendemain&"#"
Set rsgav= Server.CreateObject("ADODB.RecordSet")
rsgav.open SQLGAV,conn2,3,3
nb_gav=rsgav.recordcount
if nb_gav>0 then
rsgav.movefirst

select case Weekday(date(),2)
case 1
njour="Mon"
case 2
njour="Tue"
case 3
njour="Wed"
case 4
njour="Thu"
case 5
njour="Fri"
case 6
njour="Sat"
case 7
njour="Sun"
end select

select case Month(date())
case 1
mois="Jan"
case 2
mois="Feb"
case 3
mois="Mar"
case 4
mois="Apr"
case 5
mois="May"
case 6
mois="Jun"
case 7
mois="Jul"
case 8
mois="Aug"
case 9
mois="Sep"
case 10
mois="Oct"
case 11
mois="Nov"
case 12
mois="Dec"
end select

datejour=njour&", "&day(date())&" "&mois&" "&year(date())&" "&hour(now())&":"&minute(now())&":"&second(now())&" +0200 (CEST)"

'//////////////////////coordinateur
'//////////////////////////////////
SQLavo0="SELECT * from [Avocats_PP] where avo_code="&rsgav("num_coordinateur")
Set rsavo0= Server.CreateObject("ADODB.RecordSet")
rsavo0.open SQLavo0,conn,3,3
nb_avo0=rsavo0.recordcount

letexte0="Bonjour "&rsavo0("avo_prenom")&" "&rsavo0("avo_nom")&" !"
letexte0=letexte0&vbCrlF&vbCrlF&"Ici le bon genie de la lampe qui te rappelle que tu es coordinateur GAV demain, le "&strdate&"."

SQLmail0="SELECT * from [login] where avo="&rsavo0("avo_code")
Set rsmail0= Server.CreateObject("ADODB.RecordSet")
rsmail0.open SQLmail0,conn2,3,3
nb_mail0=rsmail0.recordcount

if nb_mail0>0 then
Zemail0=rsmail0("email")

Set Mail = Server.CreateObject("Persits.MailSender")
Mail.Host = "localhost"
Mail.From = "permanence.penale@avocats-montpellier.com" ' Adresse de l'expéditeur
Mail.FromName = "Permanence Penale" ' Nom de l'expéditeur
Mail.AddAddress Zemail0, nom_complet
Mail.AddReplyTo "permanence.penale@avocats-montpellier.com"
Mail.Subject = "Rappel permanence GAV"
Mail.Body = letexte0

Mail.Send
If Err <> 0 Then
  Response.Write "Erreur: " & Err.Description
End If

response.write(rsavo0("avo_libelle")&" "&rsavo0("avo_mail")&"<br />")

'intervenant jour 1////////////////
'//////////////////////////////////
SQLavo1="SELECT * from [Avocats_PP] where avo_code="&rsgav("num_inter_chev")
Set rsavo1= Server.CreateObject("ADODB.RecordSet")
rsavo1.open SQLavo1,conn,3,3
nb_avo1=rsavo1.recordcount

letexte1="Bonjour "&rsavo1("avo_prenom")&" "&rsavo1("avo_nom")&" !"
letexte1=letexte1&vbCrlF&vbCrlF&"Ici le bon genie de la lampe qui te rappelle que tu es de permanence GAV demain, le "&strdate&"."&vbCrlF&"Ton coordinateur peut t appeler à tout moment durant cette garde. Je compte sur ton active participation et te souhaite une agreable journee !"

SQLmail1="SELECT * from [login] where avo="&rsavo1("avo_code")
Set rsmail1= Server.CreateObject("ADODB.RecordSet")
rsmail1.open SQLmail1,conn2,3,3
nb_mail1=rsmail1.recordcount

if nb_mail1>0 then
Zemail1=rsmail1("email")

Set Mail = Server.CreateObject("Persits.MailSender")
Mail.Host = "localhost"
Mail.From = "permanence.penale@avocats-montpellier.com" ' Adresse de l'expéditeur
Mail.FromName = "Permanence Penale" ' Nom de l'expéditeur
Mail.AddAddress Zemail1, nom_complet
Mail.AddReplyTo "permanence.penale@avocats-montpellier.com"
Mail.Subject = "Rappel permanence GAV"
Mail.Body = letexte1

Mail.Send
If Err <> 0 Then
  Response.Write "Erreur: " & Err.Description
End If

response.write(rsavo1("avo_libelle")&" "&rsavo1("avo_mail")&"<br />")

end if

'intervenant jour 2////////////////
'//////////////////////////////////
SQLavo2="SELECT * from [Avocats_PP] where avo_code="&rsgav("num_inter1")
Set rsavo2= Server.CreateObject("ADODB.RecordSet")
rsavo2.open SQLavo2,conn,3,3
nb_avo2=rsavo2.recordcount

letexte2="Bonjour "&rsavo2("avo_prenom")&" "&rsavo2("avo_nom")&" !"
letexte2=letexte2&vbCrlF&vbCrlF&"Ici le bon genie de la lampe qui te rappelle que tu es de permanence GAV demain, le "&strdate&"."&vbCrlF&"Ton coordinateur peut t appeler à tout moment durant cette garde. Je compte sur ton active participation et te souhaite une agreable journee !"

SQLmail2="SELECT * from [login] where avo="&rsavo2("avo_code")
Set rsmail2= Server.CreateObject("ADODB.RecordSet")
rsmail2.open SQLmail2,conn2,3,3
nb_mail2=rsmail2.recordcount

if nb_mail2>0 then
Zemail2=rsmail2("email")

Set Mail = Server.CreateObject("Persits.MailSender")
Mail.Host = "localhost"
Mail.From = "permanence.penale@avocats-montpellier.com" ' Adresse de l'expéditeur
Mail.FromName = "Permanence Penale" ' Nom de l'expéditeur
Mail.AddAddress Zemail2, nom_complet
Mail.AddReplyTo "permanence.penale@avocats-montpellier.com"
Mail.Subject = "Rappel permanence GAV"
Mail.Body = letexte2

Mail.Send
If Err <> 0 Then
  Response.Write "Erreur: " & Err.Description
End If

response.write(rsavo2("avo_libelle")&" "&rsavo2("avo_mail")&"<br />")

end if

'intervenant nuit 1////////////////
'//////////////////////////////////
SQLavo3="SELECT * from [Avocats_PP] where avo_code="&rsgav("num_inter2")
Set rsavo3= Server.CreateObject("ADODB.RecordSet")
rsavo3.open SQLavo3,conn,3,3
nb_avo3=rsavo3.recordcount

letexte3="Bonjour "&rsavo3("avo_prenom")&" "&rsavo3("avo_nom")&" !"
letexte3=letexte3&vbCrlF&vbCrlF&"Ici le bon genie de la lampe qui te rappelle que tu es de permanence GAV demain, le "&strdate&"."&vbCrlF&"Ton coordinateur peut t appeler à tout moment durant cette garde. Je compte sur ton active participation et te souhaite une agreable soiree !"

SQLmail3="SELECT * from [login] where avo="&rsavo3("avo_code")
Set rsmail3= Server.CreateObject("ADODB.RecordSet")
rsmail3.open SQLmail3,conn2,3,3
nb_mail3=rsmail3.recordcount

if nb_mail3>0 then
Zemail3=rsmail3("email")

Set Mail = Server.CreateObject("Persits.MailSender")
Mail.Host = "localhost"
Mail.From = "permanence.penale@avocats-montpellier.com" ' Adresse de l'expéditeur
Mail.FromName = "Permanence Penale" ' Nom de l'expéditeur
Mail.AddAddress Zemail3, nom_complet
Mail.AddReplyTo "permanence.penale@avocats-montpellier.com"
Mail.Subject = "Rappel permanence GAV"
Mail.Body = letexte3

Mail.Send
If Err <> 0 Then
  Response.Write "Erreur: " & Err.Description
End If

response.write(rsavo3("avo_libelle")&" "&rsavo3("avo_mail")&"<br />")

end if

'intervenant nuit 2////////////////
'//////////////////////////////////
SQLavo4="SELECT * from [Avocats_PP] where avo_code="&rsgav("num_inter3")
Set rsavo4= Server.CreateObject("ADODB.RecordSet")
rsavo4.open SQLavo4,conn,3,3
nb_avo4=rsavo4.recordcount

letexte4="Bonjour "&rsavo4("avo_prenom")&" "&rsavo4("avo_nom")&" !"
letexte4=letexte4&vbCrlF&vbCrlF&"Ici le bon genie de la lampe qui te rappelle que tu es de permanence GAV demain, le "&strdate&"."&vbCrlF&"Ton coordinateur peut t appeler à tout moment durant cette garde. Je compte sur ton active participation et te souhaite une agreable soiree !"

SQLmail4="SELECT * from [login] where avo="&rsavo4("avo_code")
Set rsmail4= Server.CreateObject("ADODB.RecordSet")
rsmail4.open SQLmail4,conn2,3,3
nb_mail4=rsmail4.recordcount

if nb_mail4>0 then
Zemail4=rsmail4("email")

Set Mail = Server.CreateObject("Persits.MailSender")
Mail.Host = "localhost"
Mail.From = "permanence.penale@avocats-montpellier.com" ' Adresse de l'expéditeur
Mail.FromName = "Permanence Penale" ' Nom de l'expéditeur
Mail.AddAddress Zemail4, nom_complet
Mail.AddReplyTo "permanence.penale@avocats-montpellier.com"
Mail.Subject = "Rappel permanence GAV"
Mail.Body = letexte4

Mail.Send
If Err <> 0 Then
  Response.Write "Erreur: " & Err.Description
End If

response.write(rsavo4("avo_libelle")&" "&rsavo4("avo_mail")&"<br />")

end if

end if
end if
%>