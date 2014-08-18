<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Batch Print Confirm</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="No" then Cancel%>
<%If Request.form("B1")="Cancel" then ReturnProject%>
<%If Request.form("B1")="Yes" then StartProcess%>
<%OID=Request.form("OID")%>
<%PID=Request.form("PID")%>
<%If OID="" then InputError%>

<form method="POST" action="Permit_Print.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1">
		<tr>
			<td colspan="2" style="text-align: center"><b><font size="3">Do you 
			want to Print and Close the <i>selected</i> Open Permits?</font></b></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td style="text-align: center">
<input type="submit" value="Yes" name="B1"></td>
			<td style="text-align: center">
			<input type="submit" value="No" name="B1"></td>
		</tr>
	</table>
</div>
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="OID" value="<%=OID%>">
</form>


<%Sub StartProcess%>
<%OID=Request.form("OID")%>
<%PID=Request.form("PID")%>
<%Dim FieldName(55)%>
<%Dim FieldVar(55)%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Last(Permit) from tbl_Permit;")
%>
<%do until rstemp.eof %>
<%Permit=rstemp(0)%>
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
set rstemp=conntemp.execute("Select * from permit_fld;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%FieldName(X)=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%OIDArray=Split(OID,",")%>
<%For C=0 to Ubound(OIDArray)%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_open where ID=" & OIDArray(C) & ";")
%>
<%do until rstemp.eof %>
<%for Q = 1 to X%>
<%Response.write "."%>
<%FieldVar(Q)=Replace(rstemp(Q),"'","&#039")%>
<%next%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%Permit=Permit+1%>
<p align="center">Creating Permit</p>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%=Permit%><br>
<%RS = Conn.Execute("Insert Into tbl_Permit (Permit) Values ('" & Permit & "')")%>
<%Response.write "Permit created"%><br>
<%NewID = Conn.Execute("Select Max(ID) from tbl_Permit;")%>
<%Response.write "Got Next ID"%><br>
<%PermitID=NewID(0)%>
<%RS = Conn.Execute("Update tbl_Permit Set Lot= '" & Fieldvar(3) & "',P_number= '" & Fieldvar(4) & "',P_2= '" & Fieldvar(5) & "',P_street= '" & Fieldvar(6) & "',P_citystate= '" & Fieldvar(7) & "',P_zip= '" & Fieldvar(8) & "',Development= '" & Fieldvar(9) & "',Owner= '" & Fieldvar(10) & "',Owner_2= '" & Fieldvar(11) & "',O_Number= '" & Fieldvar(12) & "',O_street= '" & Fieldvar(13) & "',O_Addr2= '" & Fieldvar(14) & "',O_City= '" & Fieldvar(15) & "',O_State= '" & Fieldvar(16) & "',O_Zip= '" & Fieldvar(17) & "',C_type= '" & Fieldvar(18) & "',EST_GPD= '" & Fieldvar(19) & "',Utility= '" & Fieldvar(20) & "',Taps= '" & Fieldvar(21) & "',A_Date= '" & Fieldvar(22) & "',I_Date= '" & Fieldvar(23) & "',Ins_fee= '" & Fieldvar(24) & "',Con_Fee= '" & Fieldvar(25) & "',Tap_fee= '" & Fieldvar(26) & "',Preparer= '" & Fieldvar(27) & "',Inspector= '" & Fieldvar(28) & "',Ins_Date= '" & Fieldvar(29) & "',B_Date= '" & Fieldvar(30) & "',Phase= '" & Fieldvar(31) & "',Section= '" & Fieldvar(32) & "',Lateral_L= '" & Fieldvar(33) & "',Lateral_D= '" & Fieldvar(34) & "',Lat_Sta= '" & Fieldvar(35) & "',Map= '" & Fieldvar(36) & "',Municipality= '" & Fieldvar(37) & "',DS_Pump= '" & Fieldvar(38) & "',Field39= '" & Fieldvar(39) & "',Class= '" & Fieldvar(40) & "',DS_Manhole= '" & Fieldvar(41) & "',Contract= '" & Fieldvar(42) & "',Map_Page= '" & Fieldvar(43) & "',App_Name= '" & Fieldvar(44) & "',App_Address= '" & Fieldvar(45) & "',App_Phone= '" & Fieldvar(46) & "',App_city= '" & Fieldvar(47) & "',App_state= '" & Fieldvar(48) & "',App_Zip= '" & Fieldvar(49) & "',PID= '" & Fieldvar(50) & "' Where ID=" & PermitID & ";")%>
<!--RS = Conn.Execute("DELETE FROM tbl_open where ID = " & OIDArray(C) & ";")-->
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%IDArray = IDArray & "," & PermitID%>
<%Next%>
<%IDArray=Mid(IDArray,2,Len(IDArray))%>

<form method="POST" action="Forms/ConnectionPermit.asp" name="Update" target="_blank">
<input type="hidden" name="X" Value="<%=IDArray%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
    	</script>
<%Response.end%>

<%End Sub%>



<%Sub Cancel%>
<%PID=Request.form("PID")%>
<form method="POST" action="Permit_open.asp" name="Update">
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="Refresh" value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub ReturnProject%>
<%PID=Request.form("PID")%>
<form method="POST" action="Project_Add_3.asp" name="Update">
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="Refresh" value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub InputError%>
<%PID=Request.form("PID")%>
<p align="center"><b>You didn't select a permit to Print & Close</b></p>
<form method="POST" action="Permit_open.asp" name="Update">
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="Refresh" value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 2000); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>