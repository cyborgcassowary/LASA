<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Closed Project Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->
<%Dim M_Array(9)%>
<%M_Array(1)="East Petersburg Borough"%>
<%M_Array(2)="East Hempfield Township"%>
<%M_Array(3)="Lancaster Township"%>
<%M_Array(4)="Manheim Township"%>
<%M_Array(5)="Manor Township"%>
<%M_Array(6)="Mountville Borough"%>
<%M_Array(7)="West Hempfield Township"%>
<%M_Array(8)="Lancaster City"%>

<%D=request.QueryString("D")%>
<%F=Request.QueryString("F")%>
<%If D="" then D="ASC":F="Project"%>
<%If Request.QueryString("Status")="" then Status="Open" else Status=Request.QueryString("Status")%>
<div align="center">
<table width="100%">
	<tr>
		<td>&nbsp;</td>
		<td align="center"><b><font size="4">Project Report</font></b></td>
		<td><!--#include file="inc/print.inc"--></td>
	</tr>
</table>

	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
		<td colspan="4"><a href="Closed_Index.asp?Status=Open"><b>Open Projects</b></a></td>
		<td colspan="5" align="right"><a href="Closed_Index.asp?Status=Closed"><b>Closed Projects</b></a></td>
		</tr>
		<tr>
			<td style="text-align: center" colspan="2">
			<table width="100%"><tr>
			<td><a href="closed_index.asp?D=ASC&F=Project"><img border="0" src="../images/ASC.gif" width="18" height="15"></a></td>
			<td><b>Project</b></td>
			<td><a href="closed_index.asp?D=DESC&F=Project"><img border="0" src="../images/DESC.gif" width="18" height="15"></a></td>
			</tr>
			</table>
			</td>
			<td style="text-align: center" colspan="2">
			<table width="100%"><tr>
			<td><a href="closed_index.asp?D=ASC&F=Municipality"><img border="0" src="../images/ASC.gif" width="18" height="15"></a></td>
			<td><b>Township</b></td>
			<td><a href="closed_index.asp?D=DESC&F=Municipality"><img border="0" src="../images/DESC.gif" width="18" height="15"></a></td>
			</tr>
			</table>
			</td>
			<td style="text-align: center" width="20%"><table width="117%">
			<tr>
			<td><a href="closed_index.asp?D=ASC&F=Res_Amount"><img border="0" src="../images/ASC.gif" width="18" height="15"></a></td>
			<td align="center"><b>Reserve Type</b><br>(N/E=Not Entered)</td>
			<td><a href="closed_index.asp?D=DESC&F=Res_Amount"><img border="0" src="../images/DESC.gif" width="18" height="15"></a></td>
			</tr>
			</table>			
			</td>
			<td style="text-align: center" width="12%">
			<table width="107%"><tr>	
			<td><a href="closed_index.asp?D=ASC&F=EscrowFile"><img border="0" src="../images/ASC.gif" width="18" height="15"></a></td>
			<td><b>Escrow File</b></td>
			<td><a href="closed_index.asp?D=DESC&F=EscrowFile"><img border="0" src="../images/DESC.gif" width="18" height="15"></a></td>
			</tr>
			</table>
			</td>
			<td style="text-align: center" width="19%">
			<table width="100%"><tr>
			<td><a href="closed_index.asp?D=ASC&F=OpenDate"><img border="0" src="../images/ASC.gif" width="18" height="15"></a></td>
			<td><b>Open date</b></td>
			<td><a href="closed_index.asp?D=DESC&F=OpenDate"><img border="0" src="../images/DESC.gif" width="18" height="15"></a></td>	
			</tr>
			</table>
			</td>
			<td style="text-align: center" width="6%">
			<table width="100%"><tr>
			<td><a href="closed_index.asp?D=ASC&F=LastUpdate"><img border="0" src="../images/ASC.gif" width="18" height="15"></a></td>
			<td><%If Status="Closed" then TitleVal="Close Date" else TitleVal="Last Update"%>
			<b><%=TitleVal%></b></td>
			<td><a href="closed_index.asp?D=DESC&F=LastUpdate"><img border="0" src="../images/DESC.gif" width="18" height="15"></a></td>
			</tr>
			</table>
			</td>
			<td style="text-align: center" width="6%"><b>Permits</b>
			</td>
		</tr>
		<tr><td colspan="9"><hr width="100%"></td></tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID,Project,OpenDate,EscrowFile,LastUpdate,Phase,Res_amount,Res_type,Municipality from tbl_Project where Status='" & Status & "' order by " & F & " " & D & ";")
%>
<%IDUTotal=Conntemp.execute("Select Sum(Res_Amount) from tbl_Project where Status='" & Status & "' and Res_Type='IDU' and Res_Amount<>0;")%>
<%GPDTotal=Conntemp.execute("Select Sum(Res_Amount) from tbl_Project where Status='" & Status & "' and Res_Type='GPD' and Res_Amount<>0;")%>
<%PrjTotal=Conntemp.execute("Select Count(PID) from tbl_Project where Status='" & Status & "';")%>
<%Set TPermits = Conntemp.Execute("Select Count(ID) From tbl_Permit where PID<>'';")%>
<tr>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="11%">&nbsp;</td>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" colspan="2">Total <%=Status%> Projects=<%=PrjTotal(0)%></td>
<td align="center" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" colspan="2">Total IDU=<%=IDUTotal(0)%><br>Total GPD=<%=GPDTotal(0)%></td>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="12%">&nbsp;</td>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="19%">&nbsp;</td>
<td align="right" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" colspan="2"><%=TPermits(0)%>&nbsp;</td>
</tr>
<%do until rstemp.eof %>
<%Set Permits = Conntemp.Execute("Select Count(ID) From tbl_Permit where PID='" & rstemp(0) & "';")%>
<%TotalPermits=TotalPermits + Permits(0)%>
<%X=X+1%>
<%if X/2 = Int(X/2) then Bgcolor="#C0C0C0" else bgcolor=""%>
		<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			
			
			<td style="text-align: left" colspan="2" width="35%"><a href="Closed_Details.asp?PID=<%=rstemp(0)%>"><%=rstemp(1)%></a>&nbsp;</td>
			<td style="text-align: left" colspan="2" width="20%"><%=M_Array(Int(rstemp(8)))%></a>&nbsp;</td>
			<%if rstemp(6)="" then amt="<b>N/E</b>" else amt=rstemp(6)%>
			<td style="text-align: Center" width="20%"><%=amt%>&nbsp;<%=rstemp(7)%></td>
			<td style="text-align: center" width="12%"><%=rstemp(3)%>&nbsp;</td>
			<td style="text-align: center" width="19%"><%=FormatDateTime(rstemp(2),vbshortdate)%>&nbsp;</td>
			<td  style="text-align: center" width="6%"><%=FormatDateTime(rstemp(4),vbShortdate)%></td>
			<td  style="text-align: center" width="6%"><%=Permits(0)%></td>
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
</div>
</body>

</html>