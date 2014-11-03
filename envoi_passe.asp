<!--#include file="connexion_perm.asp"-->
<%
Znbre_box=request.form("nbre_box")
if Znbre_box<1 then response.redirect("gestion_login.asp")
Dim Zbox(9999)
for i=1 to Znbre_box
Zbox(i)=request.form("checkbox"&i)
if Zbox(i)<>"" and Zbox(i)>0 then
Zid=cint(Zbox(i))

SQLselection="SELECT * from [login] where id_login="&Zid
Set rsselection= Server.CreateObject("ADODB.RecordSet")
rsselection.open SQLselection,conn
rsselection.movefirst

select case Weekday(date(),2)
case 1
jour="Mon"
case 2
jour="Tue"
case 3
jour="Wed"
case 4
jour="Thu"
case 5
jour="Fri"
case 6
jour="Sat"
case 7
jour="Sun"
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

datejour=jour&", "&day(date())&" "&mois&" "&year(date())&" "&hour(now())&":"&minute(now())&":"&second(now())&" +0200 (CEST)"

Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.CharSet = 2
Mailer.FromName   = "Permanence Penale"
Mailer.FromAddress= "permanence.penale@avocats-montpellier.com"
Mailer.RemoteHost = "127.0.0.1"
Mailer.ContentType = "text/plain"
Mailer.DateTime = datejour
Mailer.Organization = "Ordre des avocats"
Mailer.AddExtraHeader "X-MimeOLE: Produced by Ordre des avocats"
Mailer.AddRecipient rsselection("email"),rsselection("email")
Mailer.Subject = "Vos parametres d'acces au site de la Permanence Penale"
letexte="Adresse du site: http://www.avocats-montpellier.com/permanence"&vbCrLf&vbCrLf&"Login : "
letexte=letexte&rsselection("login")
letexte=letexte&vbCrLf&"Mot de passe : "
letexte=letexte&rsselection("password")

Mailer.BodyText = letexte
Mailer.SendMail

end if
next

response.redirect("gestion_login.asp?m=2")
%>