<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Closed Project Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!-- include file="PDM_Check.inc"-->
<%D=request.QueryString("D")%>
<%F=Request.QueryString("F")%>
<%If D="" then D="ASC":F="Project"%>
<div align="center">
<b><font size="4">Completed Projects</font></b>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center" colspan="2">
			<a href="Open_Index.asp?D=ASC&F=Project"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Project</b>
			<a href="Open_Index.asp?D=DESC&F=Project"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
			<td style="text-align: center">
			<a href="Open_Index.asp?D=ASC&F=Res_Amount"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>RES</b>
			<a href="Open_Index.asp?D=DESC&F=Res_Amount"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
			<td style="text-align: center">	
			<a href="Open_Index.asp?D=ASC&F=EscrowFile"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Escrow File</b>
			<a href="Open_Index.asp?D=DESC&F=EscrowFile"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
			<td style="text-align: center">
			<a href="Open_Index.asp?D=ASC&F=OpenDate"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Open date</b>
			<a href="Open_Index.asp?D=DESC&F=OpenDate"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>	
			</td>
			<td style="text-align: center">
			<a href="Open_Index.asp?D=ASC&F=LastUpdate"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Close Date</b>
			<a href="Open_Index.asp?D=DESC&F=LastUpdate"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
			<td style="text-align: center"><b>Permits</b>
			</td>
		</tr>
		<tr><td colspan="7"><hr width="100%"></td></tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID,Project,OpenDate,EscrowFile,LastUpdate,Phase,Res_amount,Res_type from tbl_Project where Status='Open' order by " & F & " " & D & ";")
%>
<%do until rstemp.eof %>
<%Set Permits = Conntemp.Execute("Select Count(ID) From tbl_Permit where PID='" & rstemp(0) & "';")%>
<%X=X+1%>
<%if X/2 = Int(X/2) then Bgcolor="#C0C0C0" else bgcolor=""%>
		<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="text-align: center">&nbsp;</td>
			<td style="text-align: left"><a href="Closed_Details.asp?PID=<%=rstemp(0)%>"><%=rstemp(1)%></a>&nbsp;</td>
			<td style="text-align: Center"><%=rstemp(6)%>&nbsp;<%=rstemp(7)%></td>
			<td style="text-align: center"><%=rstemp(3)%>&nbsp;</td>
			<td style="text-align: center"><%=FormatDateTime(rstemp(2),vbshortdate)%>&nbsp;</td>
			<td  style="text-align: center"><%=rstemp(4)%></td>
			<td  style="text-align: center"><%=Permits(0)%></td>
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
