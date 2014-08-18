<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Page 3 Queries</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.QueryString("PID") <> "" then Session("ProjectID")=Request.QueryString("PID")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Project where PID=" & Session("ProjectID") & ";")
%>
<%do until rstemp.eof %>
<%Project=rstemp(1)%>
<%Res_Type=rstemp(2)%>
<%Res_Amount=rstemp(3)%>
<%User_Type=rstemp(4)%>
<%T_Plant=rstemp(5)%>
<%MNID=rstemp(6)%>
<%Phase=rstemp(7)%>
<%DID=rstemp(8)%>
<%EID=rstemp(9)%>
<%P_Contact=rstemp(10)%>
<%PSID=rstemp(14)%>
<%UID=rstemp(15)%>
<%D_Contact=rstemp(16)%>
<%E_Contact=rstemp(17)%>
<%REQ=rstemp(18)%>
<%EFN=rstemp(20)%>
<%GISID=rstemp(23)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<%If User_Type="R" then C_Type="Residential"%>
<%If User_Type="I" then C_Type="Industrial"%>
<%If User_Type="C" then C_Type="Commercial"%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Developer where DID=" & DID & ";")
%>
<%do until rstemp.eof %>
<%D_Name=rstemp(1)%>
<%D_Street=rstemp(2)%>
<%D_City=rstemp(3)%>
<%D_State=rstemp(4)%>
<%D_Zip=rstemp(5)%>
<%D_Number=rstemp(6)%>
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
set rstemp=conntemp.execute("Select * from tbl_Engineer where EID=" & EID & ";")
%>
<%do until rstemp.eof %>
<%E_Name=rstemp(1)%>
<%E_Street=rstemp(2)%>
<%E_City=rstemp(3)%>
<%E_State=rstemp(4)%>
<%E_Zip=rstemp(5)%>
<%E_Number=rstemp(6)%>
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
set rstemp=conntemp.execute("Select Municipality from tbl_municipality where MID=" & MNID & ";")
%>
<%do until rstemp.eof %>
<%M_Name=rstemp(0)%>
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
set rstemp=conntemp.execute("Select PS_Name from tbl_PS where Station_Number='" & PSID & "';")
%>
<%do until rstemp.eof %>
<%PS_Name=rstemp(0)%>
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
set rstemp=conntemp.execute("Select Utility from tbl_Utility where UID=" & UID & ";")
%>
<%do until rstemp.eof %>
<%Utility=rstemp(0)%>
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
set rstemp=conntemp.execute("Select ID,Status from tbl_Assign where PID='" & Project & "';")
%>
<%do until rstemp.eof %>
<%AID=rstemp(0)%>
<%Status=rstemp(1)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

</body>

</html>
