<!--#include file="connexion2.asp"-->
<%
Set upl = Server.CreateObject("SoftArtisans.FileUp")
if upl.UserFilename<>"" then
nom_fichier=Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1)
nom_fichier=lcase(nom_fichier)
ext=right(nom_fichier,3)

upl.SaveInVirtual "../../upload/imgs_pp/tmp/"&nom_fichier

Set Image = Server.CreateObject("AspImage.Image")
Image.LoadImage("D:/WS/ORDRAVO6/upload/imgs_pp/tmp/"&nom_fichier)
IF Image.MaxX > 120 THEN
Coefficient=120/Image.MaxX
W=int(Image.MaxX*Coefficient)
H=int(Image.MaxY*Coefficient)
Image.Resize W,H
END IF
Image.JPEGQuality = 80
Image.FileName = "D:/WS/ORDRAVO6/upload/imgs_pp/"&nom_fichier
Image.SaveImage
Set Image = nothing

end if

Ztitre=upl.Formex("titre")
Ztexte=upl.Formex("texte")
Zposition=upl.Formex("position")

function tarea(text)
         tarea=replace(text,"&","&amp;")
         tarea=replace(tarea,"<","&lt;")
		 tarea=replace(tarea,">","&gt;")
		 tarea=replace(tarea,VbCrLf,"<br>")
		 tarea=replace(tarea,"'","''")
end function

Ztitre=tarea(Ztitre)
Ztexte=tarea(Ztexte)

SQLadd="Insert Into [news](titre,texte,image,position_img) Values('"&Ztitre&"','"&Ztexte&"','"&nom_fichier&"','"&Zposition&"')"
Set saisie= Server.CreateObject("ADODB.RecordSet")
saisie.open SQLadd,conn2

conn2.close
Set conn2=nothing

response.redirect("infos.asp?m=1")
%>