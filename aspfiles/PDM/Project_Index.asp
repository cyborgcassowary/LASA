<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Project Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%TDate=FormatDateTime(Now,VbShortDate)%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Count(ID) from tbl_Assign where Status <> 'Complete' and DueDate < '" & TDate & "';")
%>
<%DoDD=conntemp.execute("Select Count(PID) from ph_indemn where DateDiff(m,DODD,'" & TDate & "') >= 18 and DODD <> '1/1/1990';")%>
<%If rstemp(0) + DoDD(0) > 0 and Request.Cookies("PDM")("Assignment") < TDate then PopWindow%>
<%rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<!--#include file="PDM_Check.inc"-->
<%D=request.QueryString("D")%>
<%F=Request.QueryString("F")%>
<%If D="" then D="ASC":F="Project"%>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center" colspan="2">
			<a href="Project_index.asp?D=ASC&F=Project"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Project</b>
			<a href="Project_index.asp?D=DESC&F=Project"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
			<td style="text-align: center">
			<a href="Project_index.asp?D=ASC&F=Res_Amount"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>RES</b>
			<a href="Project_index.asp?D=DESC&F=Res_Amount"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
			<td style="text-align: center">	
			<a href="Project_index.asp?D=ASC&F=EscrowFile"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Escrow File</b>
			<a href="Project_index.asp?D=DESC&F=EscrowFile"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
			<td style="text-align: center">
			<a href="Project_index.asp?D=ASC&F=OpenDate"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Open date</b>
			<a href="Project_index.asp?D=DESC&F=OpenDate"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>	
			</td>
			<td style="text-align: center">
			<a href="Project_index.asp?D=ASC&F=LastUpdate"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Last Update</b>
			<a href="Project_index.asp?D=DESC&F=LastUpdate"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
			<td style="text-align: center">
			<a href="Project_index.asp?D=ASC&F=Phase"><img border="0" src="../images/ASC.gif" width="18" height="15"></a>
			<b>Phase</b>
			<a href="Project_index.asp?D=DESC&F=Phase"><img border="0" src="../images/DESC.gif" width="18" height="15"></a>
			</td>
		</tr>
		<tr><td colspan="7"><hr width="100%"></td></tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID,Project,OpenDate,EscrowFile,LastUpdate,Phase,Res_amount,Res_type from tbl_Project where Status='Open' order by " & F & " " & D & ";")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%if X/2 = Int(X/2) then Bgcolor="#C0C0C0" else bgcolor=""%>
		<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="text-align: center"><a href="Project_Comments.asp?PID=<%=rstemp(0)%>">
			<img border="0" src="../images/dir_txt.gif" width="16" height="16" title="View/Add Comments"></a></td>
			<td style="text-align: left"><a href="Project_Add_3.asp?PID=<%=rstemp(0)%>"><%=rstemp(1)%></a>&nbsp;</td>
			<td style="text-align: Center"><%=rstemp(6)%>&nbsp;<%=rstemp(7)%></td>
			<td style="text-align: center"><%=rstemp(3)%>&nbsp;</td>
			<td style="text-align: center"><%=FormatDateTime(rstemp(2),vbshortdate)%>&nbsp;</td>
			<td  style="text-align: center"><%=rstemp(4)%></td>
			<td  style="text-align: center"><%=rstemp(5)%></td>
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
<%sub PopWindow%>
<script language="javascript" type="text/javascript">
<!--
{
window.open('inc/Assignment_PopUp.asp','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=350,height=200,left = 250,top = 200');
}
//-->
</script>
<%End Sub%>
</body>

</html>