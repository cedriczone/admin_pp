<%connection = "DBQ=" & Server.MapPath("../../data/planning_pp.mdb")&";DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}"
Set conn = Server.CreateObject("ADODB.Connection")
conn.open connection
%>