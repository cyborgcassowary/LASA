<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%PID=Request.QueryString("PID")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Project where PID=" & PID & ";")
%>
<%do until rstemp.eof%>
<%Project=rstemp(1)%>
<%Developer=rstemp(8)%>
<%Engineer=rstemp(9)%>
<%
rstemp.movenext
loop
rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select D_Name,D_Contact from tbl_Developer where DID=" & Developer & ";")
%>
<%do until rstemp.eof%>
<%D_Name=rstemp(0)%>
<%D_Contact=rstemp(1)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<%If Int(Instr(D_Contact,",")) <= 0 then Dim D_Array(1)%>
<%If Int(Instr(D_Contact,",")) > 0 then D_Array=Split(D_Contact,",") else D_Array(0)=D_Contact%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select E_Name,E_Contact from tbl_Engineer where EID=" & Engineer & ";")
%>
<%do until rstemp.eof%>
<%E_Name=rstemp(0)%>
<%E_Contact=rstemp(1)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%If Int(Instr(E_Contact,",")) <= 0 then Dim E_Array(1)%>
<%If Int(Instr(E_Contact,",")) > 0 then E_Array=Split(E_Contact,",") else E_Array(0)=E_Contact%>
</body>

</html>
