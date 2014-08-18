<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>File Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>


<form method="POST" name="frm1" action="New_Cab_index.asp" webbot-action="--WEBBOT-SELF--">
<table border="0" cellspacing="0" cellpadding="0">
<caption><b>Please Select Location Type</b></caption>
	<tr>
		<td><%If Request.form("R1")="Drawer" then sel="checked" else sel=""%>
		<input type="radio" value="Drawer" name="R1" onclick="frm1.submit()" <%=sel%>></td>
		<td>Drawer</td>
		<td><%If Request.form("R1")="BookCase" then sel="checked" else sel=""%>
		<input type="radio" value="BookCase" name="R1" onclick="frm1.submit()" <%=sel%>>&nbsp;</td>
		<td>BookCase</td>
		<td><%If Request.form("R1")="Box" then Sel="checked" else sel=""%>
		<input type="radio" value="Box" name="R1" onclick="frm1.submit()" <%=sel%>>&nbsp;</td>
		<td>Box</td>
	</tr>
</table>
</form>

<%If Request.form("R1")="" and Request.form("B1")<>"Submit" then Response.end%>

<%If Request.form("B1") <> "Submit" then Session("Loc_type")=Request.form("R1")%>
<form method="POST" action="New_Cab_index.asp" webbot-action="--WEBBOT-SELF--">
<table border="0" cellspacing="0" cellpadding="0" id="table1">
	<tr>
		<td>Location</td>
		<td># of Files</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2">
<Select name="Location" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Distinct Location from tbl_Links where Location_Type='" & Session("Loc_Type") & "';")
%>
<%do until rstemp.eof %>
<%FileCount=conntemp.execute("Select Count(ID) from tbl_Files where Location='" & rstemp(0) & "';")%>
<%LocArray=split(rstemp(0),",")%>
<%If Request.form("Location")= rstemp(0) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(0)%> - (<%=Filecount(0)%>)</option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td><td><input type="submit" value="Submit" name="B1"></td>
	</tr>
</table>
<input type="hidden" name="R1" value="<%=Session("Loc_type")%>">
</form>
<%If Request.form("B1")="Submit" then GetData%>

<%sub GetData%>
<%Location=Request.form("Location")%>

<table border="0" width="100%" cellspacing="0" cellpadding="0" style="border: 2px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
<caption><%=Location%></caption>
	<tr>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">&nbsp;</td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><b>Category</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><b>FileName</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><b>Owner</b></td>
	</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select ID,Category,FileName,Owner from tbl_Files where Location='" & Location & "';")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
		<td align="center"><a href="FS_Edit.asp?ID=<%=rstemp(0)%>"><%=X%></a>&nbsp;</td>
		<td align="center"><%=rstemp(1)%></td>
		<td><%=rstemp(2)%></td>
		<td align="center"><%=rstemp(3)%></td>
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
<%end sub%>

</body>

</html>
