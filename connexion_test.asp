<%connection = "DBQ=" & Server.MapPath("../../data/planning_pp2.mdb")&";DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}"
Set conn = Server.CreateObject("ADODB.Connection")
conn.open connection
%>