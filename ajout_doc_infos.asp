<%
Set upl = Server.CreateObject("SoftArtisans.FileUp")
nom_fichier=Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1)
nom_fichier=lcase(nom_fichier)
function tarea(text)
	Set regEx = New RegExp
   ' Casse ignor�e
   regEx.IgnoreCase = True
   ' Recherche sur toute la cha�ne
   regEx.Global = True
   regEx.Pattern = "[a�]"
   tarea = regEx.REPLACE(text,"a")
   regEx.Pattern = "[�ee�]"
   tarea = regEx.REPLACE(tarea,"e")
   tarea=replace(tarea,">","")
   tarea=replace(tarea,">","")
   tarea=replace(tarea,"'","")
   tarea=replace(tarea," ","")
   tarea=replace(tarea,"�","i")
   tarea=replace(tarea,"�","o")
   tarea=replace(tarea,"u","u")
   tarea=replace(tarea,"�","c")
   tarea=replace(tarea,"&","")
   tarea=replace(tarea,"~","")
   tarea=replace(tarea,"}","")
   tarea=replace(tarea,"#","")
   tarea=replace(tarea,"{","")
   tarea=replace(tarea,"(","")
   tarea=replace(tarea,"^","")
   tarea=replace(tarea,"@","")
   tarea=replace(tarea,"[","")
   tarea=replace(tarea,"]","")
end function

nom_fichier=tarea(nom_fichier)

upl.SaveInVirtual "../../upload/docs_infos/"&nom_fichier

response.redirect("infos.asp?m=2")
%>

