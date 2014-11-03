<!--#include file="connexion_perm.asp"-->
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.Charset="iso-8859-1"
Response.ContentType = "text/plain"
SQLinfos="SELECT * from [defilantes] order by id_defilante DESC"
Set rsinfos=server.Createobject("adodb.recordset")
rsinfos.open SQLinfos,conn,3,3
nbre_infos=rsinfos.recordcount
if nbre_infos>0 then
rsinfos.movefirst
do while not rsinfos.eof
texte2=left(rsinfos("texte_defilante"),24)
%>      
        <option value="<%=rsinfos("id_defilante")%>"><%=texte2%></option>
<%
rsinfos.movenext
loop
end if
conn.close
Set conn=nothing
%>