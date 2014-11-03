<!--#include file="connexion_perm.asp"-->
<%
Server.ScriptTimeout = 600

function tarea(text)
      Set regEx = New RegExp
      ' Casse ignorée
      regEx.IgnoreCase = True
      ' Recherche sur toute la chaîne
      regEx.Global = True
      regEx.Pattern = "[aâ]"
      tarea = regEx.REPLACE(text,"a")
      regEx.Pattern = "[éeeë]"
      tarea = regEx.REPLACE(tarea,"e")
      tarea=replace(tarea,">","")
      tarea=replace(tarea,">","")
      tarea=replace(tarea,"'","")
      tarea=replace(tarea," ","")
      tarea=replace(tarea,"î","i")
      tarea=replace(tarea,"ô","o")
      tarea=replace(tarea,"u","u")
      tarea=replace(tarea,"ç","c")
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


Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = True
Upload.Save

For Each File in Upload.Files
   nom_fichier=Mid(File.Filename, InstrRev(File.Filename, "\") + 1)
   nom_fichier=lcase(nom_fichier)

   nom_fichier=tarea(nom_fichier)

   File.SaveAsVirtual "./../upload/docs_pp/"&nom_fichier

Znewcat=Upload.Form("newcat")
Zdescription=Upload.Form("description")
Zdescription=replace(Zdescription,"'","''")

if Znewcat<>"" then
Zsur_cat=Upload.Form("sur_cat")
if Zsur_cat="" then Zsur_cat=0
Znewcat=replace(Znewcat,"'","''")
Znewcat=ucase(Znewcat)
SQLadd_cat="Insert Into [cat](nom_cat,sur_cat) Values('"&Znewcat&"',"&Zsur_cat&")"
Set saisie_cat= Server.CreateObject("ADODB.RecordSet")
saisie_cat.open SQLadd_cat,conn

SQLcatinv="SELECT * FROM [cat] order by id_cat DESC"
Set rscatinv=server.Createobject("adodb.recordset")
rscatinv.open SQLcatinv,conn,3,3
rscatinv.movefirst
Zid_cat=rscatinv("id_cat")

else
Zid_cat=Upload.Form("cat")
end if

if Zsur_cat>0 then

Zchemin=Zid_cat
SQLancetre1="SELECT * FROM [cat] where id_cat="&Zsur_cat
else

Zchemin=Zid_cat
SQLancetre0="SELECT * FROM [cat] where id_cat="&Zid_cat
Set rsancetre0=server.Createobject("adodb.recordset")
rsancetre0.open SQLancetre0,conn,3,3
rsancetre0.movefirst
Zzid_cat=rsancetre0("sur_cat")
SQLancetre1="SELECT * FROM [cat] where id_cat="&Zzid_cat
end if

Set rsancetre1=server.Createobject("adodb.recordset")
rsancetre1.open SQLancetre1,conn,3,3
if not rsancetre1.eof then
rsancetre1.movefirst
if rsancetre1("id_cat")>0 then
Zchemin=rsancetre1("id_cat")&","&Zchemin

SQLancetre2="SELECT * FROM [cat] where id_cat="&rsancetre1("sur_cat")
Set rsancetre2=server.Createobject("adodb.recordset")
rsancetre2.open SQLancetre2,conn,3,3
if not rsancetre2.eof then
rsancetre2.movefirst
if rsancetre2("id_cat")>0 then
Zchemin=rsancetre2("id_cat")&","&Zchemin
end if
end if
end if
end if

SQLadd="Insert Into [docs](cat,descr_doc,nom_doc,chemin) Values("&Zid_cat&",'"&Zdescription&"','"&nom_fichier&"','"&Zchemin&"')"
Set saisie= Server.CreateObject("ADODB.RecordSet")
saisie.open SQLadd,conn
Next


response.redirect("documents.asp?upl=1")
%>

