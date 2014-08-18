<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Chapter94 Yearly Report</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%Res_Type=Request.form("Res_Type")%>
<form method="POST" action="Chapter94.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
<table border="0" cellspacing="0" cellpadding="0" width="80%">
	<tr>
		<td width="50"><%If Res_Type="GPD" then chk="checked" else chk=""%>
		<input type="radio" value="GPD" <%=chk%> name="Res_Type">GPD</td>
		<td width="50"><%If Res_type="IDU" then chk="checked" else chk=""%>
		<input type="radio" name="Res_Type" <%=chk%> value="IDU">IDU</td>
		<td><input type="submit" value="Submit" name="B1">&nbsp;</td>
		<td><!--#include file="inc/print.inc"-->&nbsp;</td>
	</tr>
</table>
</div>
</form>


<%If Res_Type="" then Res_Type="GPD"%>

<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse; border: 2px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
	<tr>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>File#</b></td>	
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Project</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Municipality</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Res. Type</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Reserved</b></td>
		<%If Res_Type="IDU" then FieldName="Permits Issued" else FieldName="GPD Issued"%>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b><%=FieldName%></b>&nbsp;</td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b><%=Res_Type%>'s Remaining</b>&nbsp;</td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Customer Type</b></td>
	</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID,Project,Res_Type,Res_Amount,User_Type,EscrowFile,Municipality from tbl_Project where Res_Type='" & Res_Type & "' order by Municipality, Project ASC;")
%>
<%do until rstemp.eof %>
<%Set Remain=conntemp.execute("Select Count(PID) from tbl_Permit where PID='" & rstemp(0) & "' and Right(I_Date,4) = Year(Now)-1;")%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set GPD=conntemp1.execute("Select EST_GPD from tbl_Permit where PID='" & rstemp(0) & "' and C_Type='C' or C_Type='I' and Right(I_Date,4) = Year(NOw)-1;")
%>
<%do until GPD.eof %>
<%GPDTotal=GPDTotal+Int(GPD(0))%>
<%
GPD.movenext
loop

GPD.close
set GPD=nothing
conntemp1.close
set conntemp1=nothing
%>
<%If rstemp(3) <> "" then Reserved=rstemp(3) else Reserved="0"%>
<%If Res_Type="IDU" then RemainRES=Reserved-Remain(0) else RemainRES=Reserved-GPDTotal%>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<%Set Muni=Conntemp.execute("Select Municipality from tbl_Municipality where MID=" & rstemp(6) & ";")%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
		<td style="text-align: center"><%=rstemp(5)%>&nbsp;</td>
		<td><%=rstemp(1)%>&nbsp;</td>
		<td><%=Muni(0)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
		<td style="text-align: center"><%=Reserved%>&nbsp;</td>
		<td style="text-align: center"><%If Res_type="IDU" then Response.write Remain(0) else Response.write GPDTotal%>&nbsp;</td>
		<td style="text-align: center"><%=RemainRES%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(4)%>&nbsp;</td>
	</tr>
<%ResTotal=ResTotal+Int(Reserved)%>
<%RemainTotal=RemainTotal+Int(RemainRES)%>
<%GPDTotal=0%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<tr>
<td style="text-align: right"><b>Reserved/Issued:</b></td>
<td style="text-align: Left" colspan="2"><%=ResTotal%>/<%=RemainTotal%></td>
<td style="text-align: right"><b>Flow (350 GPD/Unit):</b></td>
<td style="text-align: Left" colspan="3"><%=Int(350*RemainTotal)%></td>
<td></td>
</tr>
</table>

</body>

</html>