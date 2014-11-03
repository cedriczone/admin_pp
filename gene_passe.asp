<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLliste="SELECT * from [Avocats_PP] order by avo_code"
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3

rsliste.movefirst
do while not rsliste.eof

Zlogin=rsliste("avo_mail")

SQLverif="SELECT * from [login] where login='"&Zlogin&"'"
Set rsverif=server.Createobject("adodb.recordset")
rsverif.open SQLverif,conn2,3,3
nbre_verif=rsverif.recordcount

if nbre_verif=0 then

'/////////////////////////////////////////////////////
' Génération du mot de passe
'/////////////////////////////////////////////////////

cars="az0erty2ui3op4qs5df6gh7jk8lm9wxcvbn"
wlong=len(cars)
wpas=""
taille=6
randomize time
for i=1 to taille
' Tirage aléatoire d'une valeur entre 1 et wlong
      wpos=1+int((Rnd*wlong))
' On cumule le caractère dans le mot de passe
      wpas=wpas & mid(cars,wpos,1)
' On continue avec le caractère suivant à générer      
next

'//////////////////////////////////////////////////

SQLadd_user="Insert Into [login](login,password,email) Values('"&Zlogin&"','"&wpas&"','"&Zlogin&"')"
Set saisie= Server.CreateObject("ADODB.RecordSet")
saisie.open SQLadd_user,conn2

end if

rsliste.movenext
loop
%>
