<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Closed Project Details</title>
<style type="text/css">
.style1 {
				text-align: left;
}
.style2 {
				text-align: center;
}
.style3 {
				border: 3px solid #0000FF;
}
</style>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%PID=Request.QueryString("PID")%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project,Municipality,Treatment_location,Developer,Engineer,Pumpstation,Utility,D_contact,E_contact,EscrowFile,Opendate,LastUpdate,Notes from tbl_Project where PID=" & PID & ";")
%>
<%do until rstemp.eof %>
<%Project=rstemp(0)%>
<%Municipality=rstemp(1)%>
<%Plant=rstemp(2)%>
<%DID=rstemp(3)%>
<%EID=rstemp(4)%>
<%Pumpstation=rstemp(5)%>
<%Utility=rstemp(6)%>
<%Developer=rstemp(7)%>
<%Engineer=rstemp(8)%>
<%EscrowFile=rstemp(9)%>
<%OpenDate=rstemp(10)%>
<%CloseDate=rstemp(11)%>
<%Notes=rstemp(12)%>
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
set rstemp=conntemp.execute("Select D_Name,D_Street,D_city,D_State,D_Zip,D_Number from tbl_Developer where DID=" & DID & ";")
%>
<%do until rstemp.eof %>
<%D_Name=rstemp(0)%>
<%D_Street=rstemp(1)%>
<%D_City=rstemp(2)%>
<%D_State=rstemp(3)%>
<%D_Zip=rstemp(4)%>
<%D_Number=rstemp(5)%>
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
set rstemp=conntemp.execute("Select E_Name,E_Street,E_City,E_State,E_Zip,E_Number from tbl_engineer where EID=" & EID & ";")
%>
<%do until rstemp.eof %>
<%E_Name=rstemp(0)%>
<%E_Street=rstemp(1)%>
<%E_City=rstemp(2)%>
<%E_State=rstemp(3)%>
<%E_Zip=rstemp(4)%>
<%E_Number=rstemp(5)%>
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
set rstemp=conntemp.execute("Select PS_Name from tbl_PS where ID=" & Pumpstation & ";")
%>
<%do until rstemp.eof %>
<%PumpStation=rstemp(0)%>
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
set rstemp=conntemp.execute("Select Utility from tbl_Utility where UID=" & Utility & ";")
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
set rstemp=conntemp.execute("Select Municipality from tbl_Municipality where MID=" & Municipality & ";")
%>
<%do until rstemp.eof %>
<%Municipality=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td style="font-size: 14pt; font-weight: bold"><%=Project%>&nbsp;</td>
						<td><b>Open Date:</b><%=OpenDate%></td>
						<td><b>Close Date:</b><%=CloseDate%></td>
					</tr>
					<tr>
						<td style="font-size: 14pt; font-weight: bold" colspan="3">
						&nbsp;</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td>
			&nbsp;</td>
			<td>
			<div align="center">
				<table cellpadding="0" cellspacing="0" width="100%" id="table3" style="padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
					<tr>
						<td><b>Developer:</b>&nbsp;&nbsp;<%=D_Name%></td>
					</tr>
					<tr>
						<td><%=D_Street%>&nbsp;</td>
					</tr>
					<tr>
						<td><%=D_City%>,&nbsp;<%=D_State%>&nbsp;&nbsp;<%=D_Zip%></td>
					</tr>
					<tr>
						<td><%=D_Number%>&nbsp;</td>
					</tr>
					<tr>
						<td><b>Primary Contact:</b>&nbsp;&nbsp;<%=Developer%>&nbsp;</td>
					</tr>
				</table>
			</div>
			</td>
			<td>
			<div align="center">
				<table cellpadding="0" cellspacing="0" width="100%" id="table4" style="padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
					<tr>
						<td><b>Engineer:</b>&nbsp;&nbsp;<%=E_Name%></td>
					</tr>
					<tr>
						<td><%=E_Street%>&nbsp;</td>
					</tr>
					<tr>
						<td><%=E_City%>,&nbsp;<%=E_State%>&nbsp;&nbsp;<%=E_Zip%></td>
					</tr>
					<tr>
						<td><%=E_Number%>&nbsp;</td>
					</tr>
					<tr>
						<td><b>Primary Contact:</b>&nbsp;&nbsp;<%=Engineer%>&nbsp;</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td width="50%">&nbsp;</td>
		</tr>
		<tr>
			<td style="text-align: center">&nbsp;</td>
			<td style="text-align: center"><b>Project Notes</b></td>
			<td width="50%">&nbsp;</td>
		</tr>
		<tr>
			<td rowspan="7" style="text-align: center">
			&nbsp;</td>
			<td rowspan="7" style="text-align: center">
			<textarea rows="10" name="S1" cols="60"><%=Notes%></textarea></td>
			<td width="50%"><b>Treatment Plant:</b>&nbsp;&nbsp;<%=Plant%></td>
		</tr>
		<tr>
			<td width="50%"><b>Water Utility:</b>&nbsp;&nbsp;<%=Utility%></td>
		</tr>
		<tr>
			<td width="50%"><b>Pump Station:</b>&nbsp;&nbsp;<%=PumpStation%></td>
		</tr>
		<tr>
			<td width="50%"><b>Municipality:</b>&nbsp;&nbsp;<%=Municipality%></td>
		</tr>
		<tr>
			<td width="50%"><b>Escrow File:</b>&nbsp;&nbsp;<%=EscrowFile%></td>
		</tr>
		<tr>
			<td width="50%">&nbsp;</td>
		</tr>
		<tr>
			<td width="50%">&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td width="50%">&nbsp;</td>
		</tr>
		<tr><td colspan="3" style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">
			<a href="Closed_Index.asp">Project Report</a>, 
			<a href="Project_Index.asp">Project Index</a>, 
			<a href="Permit_Detail.asp">Permit Browser</a>
			
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Issued Permits</title>
</head>

<body>
<script language="javascript" type="text/javascript">
<!--
function Closed_Details(id)
{
window.open('Closed_Details.asp?ID=' + id +'','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=300,left = 250,top = 200');
}
//-->
</script>
<table border="0" id="table1" cellspacing="0" cellpadding="0" style="padding: 1px 4px; width: 50%;" class="style3">
<caption><b><font size="4">Issued Permits</font></b></caption>
	<tr>
		<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; height: 19px;" class="style1"><b>Permit</b></td>
		<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; height: 19px;" class="style1">
		<strong>Street Address</strong></td>
		<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; height: 19px;" class="style1">
		<strong>Issue Date</strong></td>
	</tr>
	
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select ID,Permit,P_Number,P_Street,I_Date from tbl_Permit where PID='" & PID & "';")
%>
<%do until rstemp.eof %>

<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
		<td class="style2"><%=rstemp(1)%>&nbsp;</td>
		<td class="style1"><%=rstemp(2)&" "&rstemp(3)%>&nbsp;</td>
		<td class="style1"><%=rstemp(4)%>&nbsp;</td>
	</tr>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</table>
<table width="100%">
<tr>
<td align="right"><!--#include file="inc/print.inc"--></td>
</tr>
