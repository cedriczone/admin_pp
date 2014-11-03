<!--#include file="includes/connexion_perm.asp"-->
<!--#include file="includes/connexion_planning.asp"-->
<!--#include file="includes/functions.asp"-->
<%
cpt_nodes=0
cpt_des=0

demain=dateAdd("d",1,date())
demain=CDate(demain)
Zdatechoisie=month(demain)&"/"&day(demain)&"/"&year(demain)

SQLpp="SELECT * FROM [planning_pp_jour] where L01C1=#"&Zdatechoisie&"# OR L01C2=#"&Zdatechoisie&"# OR L01C3=#"&Zdatechoisie&"# OR L01C4=#"&Zdatechoisie&"# OR L01C5=#"&Zdatechoisie&"# OR L01C6=#"&Zdatechoisie&"# OR L01C7=#"&Zdatechoisie&"#"
Set rspp=server.Createobject("adodb.recordset")
rspp.open SQLpp,conn,3,3
nb_pp=rspp.recordcount

if nb_pp>0 then

   if rspp("L01C1")=demain then colonne=1
   if rspp("L01C2")=demain then colonne=2
   if rspp("L01C3")=demain then colonne=3
   if rspp("L01C4")=demain then colonne=4
   if rspp("L01C5")=demain then colonne=5
   if rspp("L01C6")=demain then colonne=6
   if rspp("L01C7")=demain then colonne=7
   
   for i=2 to 13
      ligne=i
      if ligne<10 then ligne="0"&i
      
      if rspp("L"&ligne&"C"&colonne)>0 then
         SQLavo="SELECT * from [Avocats_PP] where avo_code="&rspp("L"&ligne&"C"&colonne)
         Set rsavo= Server.CreateObject("ADODB.RecordSet")
         rsavo.open SQLavo,conn,3,3
         
         strdate = weekdayname(weekday(demain))
         strdate = strdate & " "
         If day(demain) = 1 Then
            strdate = strdate & "1er"
         Else
            strdate = strdate & day(demain)
         End If
         strdate = strdate & " " & monthname(month(demain)) & " " & year(demain)

         letexte="Bonjour "&rsavo("avo_prenom")&" "&rsavo("avo_nom")&" !"
         if i=2 then
            letexte=letexte&vbCrlF&vbCrlF&"Ici le bon genie de la lampe qui te rappelle que tu es coordinateur GAV demain, le "&strdate&"."
         else
            letexte=letexte&vbCrlF&vbCrlF&"Ici le bon genie de la lampe qui te rappelle que tu es de permanence penale demain, le "&strdate&"."
         end if 'FIN VERIF SI COORDINATEUR
         
         liste_des=designation(demain,rsavo("avo_code"))
         
            if liste_des="" then
               if i>2 then
                  letexte=letexte&vbCrlF&"Tu n as pas encore ete designe pour intervenir, mais ton coordinateur peut t appeler a toute heure. Je compte sur ton active participation et te souhaite une agreable journee !"
               end if
            else
               liste_des=replace(liste_des,"&agrave;"," - ")
               liste_des=replace(liste_des,"<br />",vbCrlF)
               letexte=letexte&vbCrlF&"Je te rappelle que tu as ete designe pour intervenir ce jour :"&vbCrlF&liste_des
            end if ' SI LA LISTE DES DESIGNATIONS N'EST PAS VIDE
         
         SQLmail="SELECT * from [login] where avo="&rsavo("avo_code")
         Set rsmail= Server.CreateObject("ADODB.RecordSet")
         rsmail.open SQLmail,conn2,3,3
         nb_mail=rsmail.recordcount
         
         if nb_mail>0 then
            Zemail=rsmail("email")
            
            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
            Mailer.CharSet = 2
            Mailer.FromName   = "Permanence Penale"
            Mailer.FromAddress= "permanence.penale@avocats-montpellier.com"
            Mailer.RemoteHost = "127.0.0.1"
            Mailer.ContentType = "text/plain"
            Mailer.DateTime = datejour
            Mailer.Organization = "Ordre des avocats"
            Mailer.AddExtraHeader "X-MimeOLE: Produced by Ordre des avocats"
            Mailer.AddRecipient nom_complet,Zemail
            'Mailer.AddRecipient "cedric","cedric@solution34.fr"
            Mailer.Subject = "Rappel permanence penale"
            
            Mailer.BodyText = letexte
            Mailer.SendMail
            
            response.write(rsavo("avo_libelle")&" "&rsavo("avo_mail")&"<br />")
         
         end if ' SI IL Y A UN MAIL EXISTANT
      
      end if 'SI SUR LA CASE IL Y A BIEN UN NUM D'AVOCAT ET PAS 0
      
   next ' BOUCLE AVOCATS DU LENDEMAIN
   
end if ' SI CE JOUR EST PRESENT DANS LA BASE

%>