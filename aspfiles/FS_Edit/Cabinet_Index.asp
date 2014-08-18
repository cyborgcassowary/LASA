<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Location Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%X=Request.QueryString("X")%>
<%If X <= 0 or X=999 then
		Response.Cookies("FS")("Location")=""
		Response.cookies("FS").expires=Date()-1
		X=0
		Location=""
End If%>
<%Location=Request.QueryString("Location")%>
<%If X => 1  then Response.cookies("FS")("Location")=Request.Cookies("FS")("Location") & "," & Location%>


<%Dim tbl(5)%>
<%tbl(1)="tbl_Building"%>
<%tbl(2)="tbl_Room"%>
<%tbl(3)="tbl_Cabinet"%>
<%tbl(4)="tbl_Drawer"%>

<%X=X+1%>
<%If X > 4 then
	 X = 1
	 Response.cookies("FS").Expires=Date()-1
	 Go=1
End If
%>
<div align="center">
<b><font size="4">File Index (Cabinets)</font></b><br><br>
		<table border="0" cellpadding="0" cellspacing="0" width="80%" id="table1">
			<tr>
				<td width="178"><b>Current Selected Location:</b>&nbsp;</td>
				<td align="left"><%=Trim(Mid(Request.Cookies("FS")("Location"),2,Len(Request.cookies("FS")("Location"))))%>&nbsp;</td>
				<td align="right"><a href="Cabinet_Index.asp?X=999">Clear Location</a></td>
			</tr>
		</table>
<br>
<table width="80%">
<tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Name from " & tbl(X) & " order by Name ASC;")
%>

<%do until rstemp.eof %>
<%If X < 4 then
	 Count=conntemp.execute("Select Count(ID) from tbl_Files where Location Like '%" & rstemp(0) & "%';")
	 	else
	  Count=conntemp.execute("Select Count(ID) from tbl_Files where Location Like '%" & rstemp(0) & "';")
End If%>
<%I=I+1%>
<td style="border-style: solid; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" align="center"><a href="Cabinet_index.asp?X=<%=X%>&Location=<%=rstemp(0)%>"><%=rstemp(0)%></a>
(<%=count(0)%>)
&nbsp;</td>


<%If I/5=Int(I/5) then response.write "</tr>"%>
<%
rstemp.movenext
loop

rstemp.close
set RecCount=Nothing
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</tr>
</table>
</div>
<%If Go=1 then GetFiles%>

<%Sub GetFiles%>
<%Location=Request.cookies("FS")("Location")%>
<%Location=Trim(Mid(Location,2,Len(Location)))%>
<br>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select ID from tbl_links where Location='" & Location & "';")
%>
<%do until rstemp.eof %>
<%ID=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%If ID="" then ID="Blank"%>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="80%" style="border: 2px outset #0000FF">
		<tr>
			<td>Files at <font color="red"><%=Location%></font>.</td>
			<td width="60"><%If ID<>"Blank" then Response.write "<a href=FS_Add.asp?Location=":Response.write ID:Response.write">Add New</a>"%>&nbsp;</td>
		</tr>
		<tr><td colspan="2">&nbsp;</td></tr>
		<tr>
			<td colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>FileName</b></td>
						<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Owner</b></td>
						<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Category</b></td>
						<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Archive Date</b></td>
						<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Destroy Date</b></td>
					</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select FileName,Owner,Category,Archive_Date,Destroy_Date,ID from tbl_Files where Location='" & Location & "';")
%>
<%do until rstemp.eof %>
<%Q=Q+1%>
<%If Q/2=Int(Q/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
						<td><a href="FS_Edit.asp?ID=<%=rstemp(5)%>"><%=rstemp(0)%></a>&nbsp;</td>
						<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
						<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
						<td style="text-align: center" width="100"><%=rstemp(3)%>&nbsp;</td>
						<td style="text-align: center" width="100"><%=rstemp(4)%>&nbsp;</td>
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
			</td>
		</tr>
	</table>
</div>
<%End Sub%>

</body>

</html>