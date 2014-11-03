<%
Server.ScriptTimeout=3000

Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = True
Count = Upload.SaveVirtual ("../../data")

If Err <> 0 Then
%>

	<h3>Une erreur a eu lieu:</h3>

	<h2>"<% = Err.Description %>"</h2>

<%
Else
%>

<h2><% = Count %> fichier bien envoy&eacute;.</h2>

<%End if%>

