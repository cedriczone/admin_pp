<html>
  <body>

    <table>
    <%
      For Each var_http In Request.ServerVariables
        Response.Write "<tr><td>" & var_http & "</td><td>" _
                      & Request.ServerVariables(var_http) & "</td></tr>"
      Next
    %>
    </table>

  </body>
</html>