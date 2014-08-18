<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Permits Closed Per Project</title>
<style type="text/css">
.style1 {
				text-align: left;
}
</style>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%PID=Request.QueryString("PID")%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project from tbl_Project where PID = " & PID & ";")
%>
<%P_Count=conntemp.execute("Select Count(PID) from tbl_permit where PID='" & PID & "';")%>
<%Closed=P_Count(0)%>
<%do until rstemp.eof %>
<%Project=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<div align="center">
<b><font size="3"><%=Closed%> Closed Permits for <%=Project%></font></b>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; width: 205px; height: 17px;">
			<b>Permit</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; width: 273px; height: 17px;">
			<b>Applicant Name</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; width: 331px; height: 17px;">
			<b>Owner Name</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; height: 17px;">
			<strong>Property Address</strong></td>
		</tr>


<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Permit,App_Name,Owner,ID,P_Number,P_Street from tbl_permit where PID = '" & PID & "' order by Permit ASC;")
%>

<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="text-align: center; width: 205px;"><a href="Permit_Edit.asp?X=<%=rstemp(3)%>"><%=rstemp(0)%></a>&nbsp;</td>
			<td style="text-align: center; width: 273px;"><%=rstemp(1)%>&nbsp;</td>
			<td style="text-align: center; width: 331px;"><%=rstemp(2)%>&nbsp;</td>
			<td style="width: 331px;" class="style1"><%=rstemp(4)&" "&rstemp(5)%>&nbsp;</td>
		</tr>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<tr>
<td colspan="4" style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px"><a href="Project_Add_3.asp?PID=<%=PID%>">Back to Project</a></td>
</tr>
</table>
</div>
</body>

</html>
