<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Create New Permit</title>
<!-- #include file="calctl.inc" -->
<script language="JavaScript" src="overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<%If Request.form("B1")="Create" then InsertData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select top 1 Permit,ID from tbl_Permit order by ID DESC;")
%>
<%NewPermit=rstemp(0)+1%>
<%NewID=rstemp(1)+1%>
<%
rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<div align="center">
<form method="POST" action="Permit_Add.asp" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellpadding="0" cellspacing="0" width="90%" id="table1" style="font-size: 8pt; border-style: solid; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
		<tr>
			<td colspan="8">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td>
						<p style="text-align: right"><b>Account# </b><input name="Account" size="6" maxlength="6" style="float: left"></td>
						<td style="text-align: right"><b>Permit #</b></td>
						<td><%=NewPermit%>&nbsp;</td>
						<td style="text-align: right"><b>Lot#</b></td>
						<td style="text-align: right">
						<input name="Field3" size="10" style="float: left"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" colspan="4">
			<b>Permit Information</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" colspan="4">
			<b>Owner Information</b></td>
		</tr>
		<tr>
			<td width="6%">Street#</td>
			<td width="5%">
			<input type="text" name="Field4" size="5"></td>
			<td width="4%" style="border-left-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			Street</td>
			<td width="29%" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<input type="text" name="Field6" size="40"></td>
			<td width="15%" colspan="3">
			<p style="text-align: right">Owner Name</td>
			<td width="32%">
			<input type="text" name="Field10" size="40"></td>
		</tr>
		<tr>
			<td width="12%" colspan="2" style="text-align: right">Address 2</td>
			<td width="36%" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="2">
			<input type="text" name="Field5" size="47"></td>
			<td width="15%" colspan="3">
			<p style="text-align: right">Owner 2</td>
			<td width="32%">
			<input type="text" name="Field11" size="40"></td>
		</tr>
		<tr>
			<td width="12%" colspan="2" style="text-align: right">City/State</td>
			<td width="36%" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="2">
			<input type="text" name="Field7" size="47"></td>
			<td width="6%">Street#</td>
			<td width="5%">
			<input type="text" name="Field12" size="5"></td>
			<td width="4%">Street</td>
			<td width="32%">
			<input type="text" name="Field13" size="40"></td>
		</tr>
		<tr>
			<td colspan="2">
			<p style="text-align: right">Zip</td>
			<td width="36%" colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<input type="text" name="Field8" size="10"></td>
			<td width="12%" colspan="2">
			<p style="text-align: right">Address 2</td>
			<td width="37%" colspan="2">
			<input type="text" name="Field14" size="47"></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right">
			<p style="text-align: right">Applicant Name</td>
			<td width="36%" colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<input type="text" name="Field44" size="40"></td>
			<td width="49%" colspan="4">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3" style="font-size: 8pt">
					<tr>
						<td align="right" width="26%">City</td>
						<td><input name="Field15" size="28" style="float: left"></td>
						<td align="right">State</td>
						<td><input type="text" name="Field16" size="2"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; text-align:right" nowrap>
			<p style="text-align: right">Applicant Address</td>
			<td width="36%" colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; ">
			<input type="text" name="Field45" size="40"></td>
			<td width="12%" colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; ">
			<p style="text-align: right">Zip</td>
			<td width="37%" colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; ">
			<input type="text" name="Field17" size="10" style="font-size: 8pt"></td>
		</tr>
		<tr>
			<td colspan="4" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:right; border-right-style:solid">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table10" style="font-size: 8pt">
					<tr>
						<td width="94">
						<p style="text-align: right">Applicant City</td>
						<td><input name="Field47" size="20" style="float: left"></td>
						<td>&nbsp;</td>
						<td>
<Select name="Field48" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=SiteSettings;"
set rstemp=conntemp.execute("Select ABBV from States order by ABBV ASC;")
%>
<%do until rstemp.eof %>
<%If rstemp(0)="PA" then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(0)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
						<td>&nbsp;</td>
						<td><input type="text" name="Field49" size="10"></td>
					</tr>
				</table>
			</td>
			<td width="12%" colspan="2" style="border-left-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			&nbsp;</td>
			<td width="37%" colspan="2" style="border-left-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:right">
			Applicant Phone</td>
			<td width="36%" colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px; border-bottom-style:solid">
			<input type="text" name="Field46" size="20"></td>
			<td width="12%" colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			&nbsp;</td>
			<td width="37%" colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			&nbsp;</td>
		</tr>
		<tr>
			<td colspan="4" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4" style="font-size: 8pt">
					<tr>
						<td>
						<p style="text-align: right">Customer Type</td>
						<td>
						<input type="checkbox" name="Field18" value="R"></td>
						<td>Residential</td>
						<td>
						<input type="checkbox" name="Field18" value="C"></td>
						<td>Commercial</td>
						<td>
						<input type="checkbox" name="Field18" value="I"></td>
						<td>Industrial</td>
					</tr>
				</table>
			</td>
			<td colspan="4" style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5" style="font-size: 8pt">
					<tr>
						<td width="132">
						<p style="text-align: right">Estimated Gallons per Day</td>
						<td>
						<input type="text" name="Field19" size="3"></td>
						<td>
						<p style="text-align: right">Map Page</td>
						<td>
						<input type="text" name="Field36" size="12"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			Municipality</td>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<select size="1" name="Field37">
			<option>Select Municipality</option>
			<option>EAST PETERSBURG BOROUGH</option>
			<option>EAST HEMPFIELD</option>
			<option>LANCASTER TOWNSHIP</option>
			<option>MANHEIM TOWNSHIP</option>
			<option>MANOR TOWNSHIP</option>
			<option>MOUNTVILLE BOROUGH</option>
			<option>WEST HEMPFIELD TOWNSHIP</option>
			</select></td>
			<td colspan="4" style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table6">
					<tr>
						<td width="131">
						<p style="text-align: right">Supplying Water Utility</td>
						<td>
						<select size="1" name="Field20">
						<option>Select Water Utility</option>
						<option>COLUMBIA</option>
						<option>EAST HEMPFIELD</option>
						<option>EAST PETERSBURG</option>
						<option>LANCASTER CITY</option>
						<option>WELL</option>
						</select></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			# of IDU's</td>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<input type="text" name="Field21" size="3"></td>
			<td colspan="4" style="border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table7">
					<tr>
						<td width="132">
						<p style="text-align: right">Contract #</td>
						<td>
						<p style="text-align: center">
						<input name="Field42" size="4" style="float: left"></td>
						<td>
						<p style="text-align: right">Page #</td>
						<td width="74">
						<input name="Field43" size="4" style="float: left"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: center; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px"><b>Lateral Information</b></td>
			<td colspan="4" style="text-align: center"><b>Inspection Information</b></td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table8" style="font-size: 8pt">
					<tr>
						<td width="100">
						<p style="text-align: right">Length</td>
						<td width="69">
						<input name="Field33" size="10" style="float: left"></td>
						<td>
						<p style="text-align: right">Depth</td>
						<td width="97">
						<input type="text" name="Field34" size="10"></td>
					</tr>
				</table>
			</td>
			<td colspan="2" style="text-align: right">Inspector</td>
			<td colspan="2">
			<input type="text" name="Field28" size="20"></td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table9" style="font-size: 8pt">
					<tr>
						<td width="100">
						<p style="text-align: right">Lateral Station</td>
						<td>
						<input name="Field35" size="10" style="float: left"></td>
						<td>
						<p style="text-align: right">DS Manhole</td>
						<td width="96">
						<input type="text" name="Field41" size="10"></td>
					</tr>
				</table>
			</td>
			<td colspan="2" style="text-align: right">Inspection Date</td>
			<td colspan="2">
			<input type="text" name="Field29" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('FrontPage_Form1.Field29');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px">DS Pump Station</td>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<Select name="Field38" size="1">
<option value="99">Select Pump Station</option>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select * from tbl_PS;")
%>
<%do until rstemp1.eof %>
<option value="<%=rstemp1(0)%>"><%=rstemp1(1)%></option>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
</Select>
			</td>
			<td colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:right">
			Inspection Fee</td>
			<td colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
<input type="text" name="Field24" size="20"></td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
			<p style="text-align: center"><b>Permit Information</b></td>
			<td colspan="4">
			<p style="text-align: center"><b>Miscellaneous Information</b></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right">Application Date</td>
			<td colspan="2" style="text-align: left; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
<input type="text" name="Field22" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('FrontPage_Form1.Field22');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">			
			&nbsp;</td>
			<td colspan="2" style="text-align: right">Prepared by</td>
			<td colspan="2">
			<input type="text" name="Field27" size="20"></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right">Issue Date</td>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<input type="text" name="Field23" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('FrontPage_Form1.Field23');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			&nbsp;</td>
			<td colspan="2" style="text-align: right">Phase</td>
			<td colspan="2">
			<input type="text" name="Field31" size="20"></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right">Billing Date</td>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<input type="text" name="Field30" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('FrontPage_Form1.Field30');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			&nbsp;</td>
			<td colspan="2" style="text-align: right">Section</td>
			<td colspan="2">
			<input type="text" name="Field32" size="20"></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right">Connection Fee</td>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<input type="text" name="Field25" size="10"></td>
			<td colspan="2" style="text-align: right">Development</td>
			<td colspan="2">
			<input type="text" name="Field9" size="20"></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right">Tap Fee</td>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<input type="text" name="Field26" size="10"></td>
			<td colspan="2" style="text-align: right">Class</td>
			<td colspan="2">
			<input type="text" name="Field40" size="20"></td>
		</tr>
		</table>
<input type="hidden" name="NewPermit" value="<%=NewPermit%>">
<input typClass</td>
			<td colspan="2">
			<input type="text" name="Field40" size="20"></td>
		</tr>
		</table>
<input type="hidden" name="NewPermit" value="<%=NewPermit%>">
<input type="hidden" name="NewID" value="<%=NewID%>">
<input type="submit" value="Create" name="B1"><input type="submit" value="Cancel" name="B1"></p>
</form>
</div>

<%Sub InsertData%>
<%Account=Request.form("Account")%>
<%NewPermit=Request.form("NewPermit")%>
<%NewID=Request.form("NewID")%>
<%DIM Fields(45)%>

<%Dim FieldName(45)%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select permit_fld_name from permit_fld;")
%>
<%do until rstemp.eof %>
<%Q=Q+1%>
<%FieldName(Q)=rstemp(0)%>

<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
NewPermit:<%=NewPermit%><br>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Insert Into tbl_Permit (Permit,Account) Values ('" & NewPermit & "','" & Account & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%For I = 3 to 49%>
<%Data="Field"& I%>
<%Fields(I)=Request.form(Data)%>
<%Fields(I)=Replace(Fields(I),"'","''")%>
<%If I=22 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=23 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=29 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=30 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%=I%>:<%=Fieldname(I)%>:<%=Fields(I)%><br>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%If I=22 or I = 23 or I=29 or I=30 then RS = Conn.Execute("Update tbl_Permit Set " & FieldName(I) & " = #" & Fields(I) & "# Where ID=" & NewID & ";") else RS = Conn.Execute("Update tbl_Permit Set " & FieldName(I) & " = '" & Fields(I) & "' Where ID=" & NewID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%next%>
<%Response.Cookies("PDM")("MaxID")=NewID%>
<p align="center"><b>Permit has been added to the database.		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<form method="POST" action="Permit_Detail.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>