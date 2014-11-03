<%connection2 = "DBQ=" & Server.MapPath("../../data/base_perm_test.mdb")&";DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}"
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.open connection2
%>