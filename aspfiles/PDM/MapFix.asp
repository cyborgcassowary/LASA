<%Response.buffer=False%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Fix Map Pages from Approach</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%Server.ScriptTimeout=999%>

<%Dim MapValue(10000)%>
<%Dim Contract(10000)%>
<%Dim Page(10000)%>
<%Dim Account(10000)%>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt; font-family: Tahoma; border: 2px outset #0000FF">
		<tr>
			<td></td>
			<td style="text-align: center"><b>dbID</b></td>
			<td style="text-align: center"><b>Original Value</b></td>
			<td style="text-align: center"><b>Account</b></td>
			<td style="text-align: center"><b>Map Value</b></td>
			<td style="text-align: center"><b>Contract&nbsp;</b></td>
			<td style="text-align: center"><b>Map Page&nbsp;</b></td>
			<td style="text-align: center"><b>New Value&nbsp;</b></td>
			<td></td>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select ID,Map,Account from map_Permit where Map <> Null;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%account(X)=rstemp(2)%>

<%Map=rstemp(1)%>
<%Map=Replace(Map,"/","")%>
<%Map=Replace(Map,"+","-")%>
<%Map=Replace(Map,"WA-","WA")%>
<%Map=Replace(Map,"_","-")%>
<%Map=Replace(Map,"`","")%>
<%Map=Replace(Map," ","")%>
<%Map=replace(Map,"SHT","")%>
<%Map=replace(map,"200","20")%>
<%Map=replace(map,"112","12")%>
<%Map=Replace(map,"Sht#","")%>
<%map=replace(map,"#","")%>
<%Map=Replace(Map,"LINEN ","LINEN-")%>
<%Map=replace(Map,"'","")%>
<%If Instr(Map,"WA10") and NOT Instr(Map,"-") then Map=Replace(Map,"WA10","WA10-")%>
<%If instr(Map,"WAWA") then Map=Replace(Map,"WAWA","WA")%>
<%If Instr(Map,"-") then NewVal=Split(Map,"-")%>
<%NewVal(0)=Trim(NewVal(0))%>
<%NewVal(1)=trim(NewVal(1))%>
<%If IsNumeric(Left(NewVal(0),1)) then NewVal(0)="WA" & NewVal(0)%>
<%If Instr(NewVal(1),"&") then NewVal(1)=Left(NewVal(1),Int(Instr(NewVal(1),"&"))-1)%>
<%MapVal=NewVal(0) & "-" & NewVal(1)%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<%MapValue(X)=MapVal%>
<%Contract(X)=NewVal(0)%>
<%Page(X)=NewVal(1)%>
		</tr>
				<tr bgcolor="<%=BgColor%>">
			<td><%=X%></td>
			<td><%=rstemp(0)%></td>
			<td><%=rstemp(2)%></td>
			<td><%=rstemp(1)%></td>
			<td><%=Map%>&nbsp;</td>
			<td><%=NewVal(0)%>&nbsp;</td>
			<td><%=NewVal(1)%>&nbsp;</td>
			<td><%=MapVal%>&nbsp;</td>
			<td><%If MapVal<>"-" then Update%></td>
		</tr>
<%Map=""%>
<%ReDim NewVal(2)%>
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

<%sub Update%>

<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%If Contract(I)<> "" and Page(I)<>"" and account(I)<> "" then RS = Conn.Execute("Update tbl_Permit Set Contract='" & Contract(I) & "', Map_Page='" & Page(I) & "', Map='" & MapValue(I) & "' Where Account='" & Account(I) & "';")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
Updated
<%End Sub%>
</body>

</html>