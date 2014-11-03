<%connection2 = "DBQ=" & Server.MapPath("../../data/test_base_perm.mdb")&";DRIVER={Microsoft Access Driver (*.mdb)}"
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.open connection2
%>