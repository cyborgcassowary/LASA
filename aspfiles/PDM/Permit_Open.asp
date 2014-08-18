<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Open Permit Index</title>
<%PID=Request.QueryString("PID")%>
<%If PID="" then PID=Request.Form("PID")%>
<script language="javascript" type="text/javascript">
<!--
function opendelete(id)
{
window.open('Open_Delete.asp?OID=' + id +'&PID=<%=PID%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=200,left = 250,top = 200');
}
//-->
</script>
<style type="text/css">
.style1 {
				font-weight: bold;
				border-left-color: #808080;
				border-left-width: 1px;
				border-right-color: #808080;
				border-right-width: 1px;
				border-top-color: #808080;
				border-top-width: 1px;
				border-bottom: 1px solid #808080;
}
.style2 {
				text-align: center;
				font-weight: bold;
				font-size: small;
}
.style3 {
				text-align: center;
}
</style>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project from tbl_Project where PID=" & PID & ";")
%>
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
<p class="style3"><br>
</p>
<div align="center">
<form method="POST" action="Permit_Print.asp">
<b><font size="4">Open Permits for <%=Project%>
	</font></b>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="border: 2px outset #0000FF; font-size:8pt">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			&nbsp;</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			Update</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			Delete</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Municipality</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; width: 197px;">
			<b>Serving Water Utility</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; width: 215px;">
			<b>Downstream Pump Station</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; width: 224px;" class="style1">
			Property Address</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Lot #</b></td>
		</tr>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select ID,Municipality,Utility,DS_Pump,Lot,P_Number,P_Street from tbl_open where PID = '" & PID & "';")
%>
<%do until rstemp1.eof %>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PS_name from tbl_PS where ID = " & rstemp1(3) & ";")
%>
<%do until rstemp.eof %>


<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<%OID=rstemp1(0)%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="text-align: center"><%=X%></td>
			<td style="text-align: center"><a href="Open_Edit.asp?OID=<%=rstemp1(0)%>&PID=<%=PID%>" title="Update Permit Information"><img border="0" src="../images/dir_txt.gif" width="16" height="16"></a></td>
			<td style="text-align: center"><a href="javascript:opendelete(<%=OID%>)"><img border="0" src="../images/dir_delete.gif" width="16" height="16"></a></td>
			<td style="text-align: center"><%=rstemp1(1)%></td>
			<td style="text-align: center; width: 197px;"><%=rstemp1(2)%>&nbsp;</td>
			<td style="text-align: center; width: 215px;"><%=rstemp(0)%>&nbsp;</td>
			<td style="text-align: center; width: 224px;"><%=rstemp1(5)&" "&rstemp1(6)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp1(4)%>&nbsp;</td>
		</tr>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>	
	</table>
	<input type="submit" value="Cancel" name="B1">
	<input type="hidden" name="PID" value="<%=PID%>">
</form>
</div>

<p><br>
</p>
<p class="style2">&nbsp;</p>
<p class="style2">&nbsp;</p>
<p>&nbsp;</p>

</body>

</html>