<%Response.buffer=false%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Edit Permit Details</title>
<!-- #include file="calctl.inc" -->
<script language="JavaScript" src="overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>


<%OID=Request.QueryString("OID")%>
<%If OID="" then OID=Request.form("OID")%>
<%PID=Request.QueryString("PID")%>
<%If PID="" then PID=Request.Form("PID")%>
<%If Request.form("B1")="Update" then Updatedata%>
<%If Request.form("B1")="Cancel" then CancelUpdate%>
<%If Request.form("B1")="Print & Close" then GetPermit%>
<div align="center">

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_open where ID=" & OID & ";")
%>
<%SFFeeID=rstemp(53)%>
<%US_Manhole=rstemp(54)%>
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<form method="POST" action="Open_Edit.asp" name="permit" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="font-size: 8pt; border-style: solid; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
		<tr>
			<td colspan="4">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td><table><tr><td><b>dbID</b></td><td><%=rstemp(0)%></td></tr></table></td>
						<td><table><tr><td><b>Account #</b></td><td><input name="Field1" size="6" maxlength="6" value="<%=rstemp(1)%>"></td></tr></table></td>
						<td><table><tr><td><b>Lot#</b></td><td><input type="text" name="Field3" size="15" value="<%=rstemp(3)%>"></td></tr></table></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" colspan="2">
			<b>Permit Information</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" colspan="2">
			<b>Owner Information</b></td>
		</tr>
		<tr>
			<td width="44%" colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table10">
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table11">
								<tr>
									<td width="92" style="text-align: right">Street#</td>
									<td width="35">
			<input type="text" name="Field4" size="5" value="<%=rstemp(4)%>"></td>
									<td style="text-align: right" width="44">Street</td>
									<td>
			<input type="text" name="Field6" size="29" value="<%=rstemp(6)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table12" style="font-size: 8pt">
								<tr>
									<td width="92" style="text-align: right">Address 2</td>
									<td>
			<input type="text" name="Field5" size="47" value="<%=rstemp(5)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table13" style="font-size: 8pt">
								<tr>
									<td width="92" style="font-size: 8pt; text-align:right">City/State</td>
									<td>
			<input type="text" name="Field7" size="47" value="<%=rstemp(7)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table28">
								<tr>
									<td style="text-align: right" width="92">Zip</td>
									<td>
			<input type="text" name="Field8" size="10" value="<%=rstemp(8)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table14" style="font-size: 8pt">
								<tr>
									<td width="92" style="text-align: right">Applicant Name</td>
									<td>
			<input type="text" name="Field44" size="47" value="<%=rstemp(44)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table15" style="font-size: 8pt">
								<tr>
									<td width="92" style="text-align: right">Applicant Address</td>
									<td>
			<input type="text" name="Field45" size="47" value="<%=rstemp(45)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table16" style="font-size: 8pt">
								<tr>
									<td width="92" style="text-align: right">Applicant City</td>
									<td width="148">
			<input type="text" name="Field47" size="28" value="<%=rstemp(47)%>"></td>
									<td style="text-align: right" width="48">State</td>
									<td><Select name="Field48" size="1">
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=SiteSettings;"
set rstemp1=conntemp1.execute("Select ABBV from States;")
%>
<%do until rstemp1.eof %>
<%If rstemp(48)=rstemp1(0) then sel="selected" else sel=""%>
<option value="<%=rstemp1(0)%>" <%=sel%>><%=rstemp1(0)%></option>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
</Select></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table29">
								<tr>
									<td style="text-align: right" width="92">Zip</td>
									<td width="59">
									<input type="text" name="Field49" size="10" value="<%=rstemp(49)%>"></td>
									<td style="text-align: right" width="92">Applicant Phone</td>
									<td>
			<input type="text" name="Field46" size="15" value="<%=rstemp(46)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table17">
								<tr>
									<td style="font-size: 8pt; text-align:right">
									&nbsp;</td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					</table>
			</div>
			</td>
			<td width="47%" colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table18">
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table19">
								<tr>
									<td width="92" style="text-align: right">Owner Name</td>
									<td>
			<input type="text" name="Field10" size="47" value="<%=rstemp(10)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table20">
								<tr>
									<td width="92" style="text-align: right">Owner 2</td>
									<td>
			<input type="text" name="Field11" size="47" value="<%=rstemp(11)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table21">
								<tr>
									<td width="92" style="text-align: right">Street#</td>
									<td width="33">
			<input type="text" name="Field12" size="5" value="<%=rstemp(12)%>"></td>
									<td width="44" style="text-align: right">Street</td>
									<td>
			<input type="text" name="Field13" size="29" value="<%=rstemp(13)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table22">
								<tr>
									<td width="92" style="text-align: right">Address 2</td>
									<td>
			<input type="text" name="Field14" size="47" value="<%=rstemp(14)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table23">
								<tr>
									<td width="92" style="text-align: right">City</td>
									<td width="154">
						<input name="Field15" size="28" value="<%=rstemp(15)%>" style="float: left"></td>
									<td style="text-align: right" width="48">State</td>
									<td><Select name="Field16" size="1">
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=SiteSettings;"
set rstemp1=conntemp1.execute("Select ABBV from States;")
%>
<%do until rstemp1.eof %>
<%If rstemp(16)=rstemp1(0) then sel="selected" else sel=""%>
<option value="<%=rstemp1(0)%>" <%=sel%>><%=rstemp1(0)%></option>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
</Select></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table30">
								<tr>
									<td style="text-align: right" width="92">Zip</td>
									<td>
									<input type="text" name="Field17" size="10" value="<%=rstemp(17)%>"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
				</table>
			</div></td>
</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4" style="font-size: 8pt">
					<tr>
						<td style="text-align: right">
						Customer Type</td>
						<%If Int(Instr(rstemp(18),"R")) > 0 then chk="checked" else chk=""%>
						<td style="text-align: center">
						<input type="checkbox" name="Field18" value="R" <%=chk%>></td>
						<td>Residential</td>
						<%If Int(Instr(rstemp(18),"C")) > 0 then chk="checked" else chk=""%>
						<td style="text-align: center">
						<input type="checkbox" name="Field18" value="C" <%=chk%>></td>
						<td>Commercial</td>
						<%If Int(Instr(rstemp(18),"I")) > 0 then chk="checked" else chk=""%>
						<td style="text-align: center">
						<input type="checkbox" name="Field18" value="I" <%=chk%>></td>
						<td>Industrial</td>
					</tr>
				</table>
			</td>
			<td colspan="2" style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5" style="font-size: 8pt">
					<tr>
						<td width="135" style="text-align: right">
						Estimated Gallons per Day</td>
						<td>
						<input type="text" name="Field19" size="3" value="<%=rstemp(19)%>"></td>
						<td style="text-align: right"><b>Map Page:</b></td>
						<td width="234"><%=rstemp(36)%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: right; " colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table25">
					<tr>
						<td style="text-align: right" width="92">Municipality</td>
						<td>
<Select name="Field37" size="1">
<option>SELECT MUNICIPALITY</option>
<% 
set mtemp=server.createobject("adodb.connection")
mtemp.open "DSN=PDM;"
set mrstemp=mtemp.execute("Select * from tbl_municipality;")
%>
<%do until mrstemp.eof %>
<%If rstemp(37)=mrstemp(1) then sel="selected" else sel=""%>
<option value="<%=mrstemp(1)%>" <%=sel%>><%=mrstemp(1)%></option>
<%
mrstemp.movenext
loop

mrstemp.close
set mrstemp=nothing
mtemp.close
set mtemp=nothing
%>
</Select></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="2" style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table6">
					<tr>
						<td width="135" style="text-align: right">
						Supplying Water Utility</td>
						<td>
<Select name="Field20" size="1">
<option>SELECT UTILITY</option>
<% 
set utemp=server.createobject("adodb.connection")
utemp.open "DSN=PDM;"
set urstemp=utemp.execute("Select * from tbl_Utility;")
%>
<%do until urstemp.eof %>
<%If rstemp(20)=urstemp(1) then sel="selected" else sel=""%>
<option value="<%=urstemp(1)%>" <%=sel%>><%=urstemp(1)%></option>
<%
urstemp.movenext
loop

urstemp.close
set urstemp=nothing
utemp.close
set utemp=nothing
%>
</Select></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: right; border-left-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table26">
					<tr>
						<td style="text-align: right" width="92"># of IDU's</td>
						<td>
			<input type="text" name="Field21" size="3" value="<%=rstemp(21)%>"></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="2" style="border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table7" style="font-size: 8pt">
					<tr>
						<td width="135" style="text-align: right">
						Contract #</td>
						<td width="34">
<Select name="Field42" size="1">
<% 
set ctemp=server.createobject("adodb.connection")
ctemp.open "DSN=PDM;"
set seltemp=ctemp.execute("Select Contract from tbl_Contract order by Contract ASC;")
%>
<%do until seltemp.eof %>
<%If SelTemp(0)=rstemp(42) Or Instr(rstemp(42),seltemp(0)) then sel="selected" else sel=""%>
<option value="<%=seltemp(0)%>" <%=sel%>><%=seltemp(0)%></option>
<%
seltemp.movenext
loop

seltemp.close
set seltemp=nothing
ctemp.close
set ctemp=nothing
%>
</Select></td>
						<td width="43" style="text-align: right">
						Page #</td>
						<td>
<input name="Field43" size="4" value="<%=rstemp(43)%>" style="float: left"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: center; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px"><b>Lateral Information</b></td>
			<td colspan="2" style="text-align: center"><b>Inspection Information</b></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
				&nbsp;</td>
			<td style="text-align: right" width="10%" colspan="2">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table34" style="font-size: 8pt">
					<tr>
						<td style="text-align: right">
						Inspection Status</td>
						<%If Int(Instr(rstemp(59),"Ready for Inspection")) > 0 then chk="checked" else chk=""%>
						<td style="text-align: center">
						<input type="checkbox" name="Field59" value="Ready for Inspection" <%=chk%>></td>
						<td>Ready for Inspection</td>
						<%If Int(Instr(rstemp(59),"Hold")) > 0 then chk="checked" else chk=""%>
						<td style="text-align: center">
						<input type="checkbox" name="Field59" value="Hold" <%=chk%>></td>
						<td>Hold</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table8" style="font-size: 8pt">
					<tr>
						<td width="92" style="text-align: right">
						Length</td>
						<td width="69">
						<input name="Field33" size="10" style="float: left" value="<%=rstemp(33)%>"></td>
						<td style="text-align: right">
						Depth</td>
						<td width="146">
						<input type="text" name="Field34" size="10" value="<%=rstemp(34)%>"></td>
					</tr>
				</table>
			</td>
			<td style="text-align: right" width="10%">Inspector</td>
			<td>
			<input type="text" name="Field28" size="20" value="<%=rstemp(28)%>"></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table9" style="font-size: 8pt">
					<tr>
						<td width="92" style="text-align: right">
						Lateral Station</td>
						<td width="69">
						<input name="Field35" size="10" value="<%=rstemp(35)%>" style="float: left"></td>
						<td style="text-align: right">
						DS/US Manhole</td>
						<td width="73">
						<input type="text" name="Field41" size="10" value="<%=rstemp(41)%>"></td>
						<td width="73">
						<input type="text" name="Field54" size="10" value="<%=US_Manhole%>"></td>
					</tr>
				</table>
			</td>
			<td style="text-align: right" width="10%">Inspection Date</td>
			<%If rstemp(29)="1/1/1900" then Field29="" else Field29=rstemp(29)%>
			<td><input type="text" name="Field29" value="<%=Field29%>" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('permit.Field29');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			</td>
		</tr>
		<tr>
			<td style="text-align: right; border-left-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table24">
					<tr>
						<td style="text-align: right">
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table31">
								<tr>
									<td width="104">DS Pump Station</td>
									<td width="129">
			<Select name="Field38" size="1">
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select Station_Number,PS_Name from tbl_PS;")
%>
<%do until rstemp1.eof %>
<%If Int(rstemp(38))=Int(rstemp1(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp1(0)%>" <%=sel%>><%=rstemp1(1)%></option>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
</Select></td>
									<td align="right" colspan="2">&nbsp;</td>
									<td colspan="4">&nbsp;</td>
								</tr>
								<tr>
									<td width="104">Facility Fee</td>
									<td width="129">
			<Select name="Field53" size="1">
<option value="0">None</option>
<% 
set ffconntemp=server.createobject("adodb.connection")
ffconntemp.open "DSN=PDM;"
set ffrstemp=ffconntemp.execute("Select * from tbl_Facility order by FacilityName ASC;")
%>
<%do until ffrstemp.eof %>
<%If Int(SFFeeID)=Int(ffrstemp(0)) then sel="selected" else sel=""%>
<option value="<%=ffrstemp(0)%>" <%=sel%>><%=ffrstemp(1)%> - $<%=ffrstemp(2)%></option>
<%
ffrstemp.movenext
loop

ffrstemp.close
set ffrstemp=nothing
ffconntemp.close
set ffconntemp=nothing
%>
</Select></td>
									<td>
									<%If rstemp(56)="Escrow" then sel="checked" else sel=""%>
									<input type="radio" name="Field56" value="Escrow" <%=sel%>></td>
									<td>Escrow</td>
									<td>
									<%If rstemp(56)="Payment Due" then sel="checked" else sel=""%>
									<input type="radio" name="Field56" value="Payment Due" <%=sel%>></td>
									<td>Payment Due</td>
									<td>
									<%If rstemp(56)="None" then sel="checked" else sel=""%>
									<input type="radio" name="Field56" value="None" <%=sel%>></td>
									<td>None</td>
								</tr>
							</table>
						</div>
</td>
		<%if Instr(rstemp(24),",") then
			TmpArray=Split(rstemp(24),",")
			InsFee=TmpArray(0)
			P_type=Trim(TmpArray(1))
		else
			InsFee=rstemp(24)
		End If%>
				</tr>
				</table>
			</div>
			</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:right" width="10%">
			Inspection Fee</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<table width="100%" style="font-size: 8pt; font-family: Tahoma">
			<tr>
				<td><input type="text" name="Field24" size="20" value="<%=InsFee%>"></td>
				<td style="text-align: right">Escrow</td><td>
				<%If P_Type="Escrow" then chk="checked" else chk=""%>
				<input type="radio" name="Field24" value="Escrow" <%=chk%>></td>
				<td style="text-align: right">Payment Due</td><td>
				<%If P_Type="Payment Due" then chk="checked" else chk=""%>
				<input type="radio" name="Field24" value="Payment Due" <%=chk%>></td>
				<td style="text-align: right">None</td><td>
				<%If P_Type="None" then chk="checked" else chk=""%>
				<input type="radio" name="Field24" value="None" <%=chk%>></td>
			</tr>
		</table>
		</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
			<p style="text-align: center"><b>Permit Information</b></td>
			<td colspan="2">
			<p style="text-align: center"><b>Miscellaneous Information</b></td>
		</tr>
		<tr>
			<td style="text-align: right">Application Date</td>
			<td style="text-align: left; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
<%If rstemp(22)="1/1/1900" then Field22="" else Field22=rstemp(22)%>
<input type="text" name="Field22" value="<%=Field22%>" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('permit.Field22');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">			
			&nbsp;</td>
			<td style="text-align: right">Prepared by</td>
			<td>
			<input type="text" name="Field27" size="20" value="<%=rstemp(27)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right">Issue Date</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
<%If rstemp(23)="1/1/1900" then Field23="" else Field23=rstemp(23)%>
			<input type="text" name="Field23" value="<%=Field23%>" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('permit.Field23');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			&nbsp;</td>
			<td style="text-align: right">Phase</td>
			<td>
			<input type="text" name="Field31" size="20" value="<%=rstemp(31)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right">Billing Date</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<%If rstemp(30)="1/1/1900" then Field30="" else Field30=rstemp(30)%>
			<input type="text" name="field30" value="<%=Field30%>" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('permit.field30');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			&nbsp;</td>
			<td style="text-align: right">Section</td>
			<td>
			<input type="text" name="Field32" size="20" value="<%=rstemp(32)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right">Connection Fee</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table32">
				<tr>
					<td width="58">
					<input type="text" name="Field25" size="10" value="<%=rstemp(25)%>"></td>
					<td style="text-align: center" width="32">
					
					<%If rstemp(57)="Escrow" then sel="checked" else sel=""%>
					<input type="radio" name="Field57" value="Escrow" <%=sel%>></td>
					<td width="45">Escrow</td>
					<td style="text-align: center">
					<%If rstemp(57)="Payment Due" then sel="checked" else sel=""%>
					<input type="radio" name="Field57" value="Payment Due" <%=sel%>></td>
					<td>Payment Due</td>
					<td>
					<%If rstemp(57)="None" then sel="checked" else sel=""%>
					<input type="radio" name="Field57" value="None" <%=sel%>></td>
					<td>None</td>
				</tr>
			</table>
			</td>
			<td style="text-align: right">Development</td>
			<td>
			<input type="text" name="Field9" size="20" value="<%=rstemp(9)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right">Tap Fee</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table33">
				<tr>
					<td width="59">
					<input type="text" name="Field26" size="10" value="<%=rstemp(26)%>"></td>
					<td width="32" style="text-align: center">
					<%If rstemp(58)="Escrow" then sel="checked" else sel=""%>
					<input type="radio" name="Field58" value="Pmt Due from Escrow" <%=sel%>></td>
					<td width="46">Escrow</td>
					<td style="text-align: center">
					<%If rstemp(58)="Payment Due" then sel="checked" else sel=""%>
					<input type="radio" name="Field58" value="Payment Due" <%=sel%>></td>
					<td>Payment Due</td>
					<td>
					<%If rstemp(58)="None" then sel="checked" else sel=""%>
					<input type="radio" name="Field58" value="None" <%=sel%>></td>
					<td>None</td>
				</tr>
			</table>
			</td>
			<td style="text-align: right">Check#</td>
			<td>
			<input type="text" name="Field40" size="20" value="<%=rstemp(40)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px" colspan="4">&nbsp;</td>
		</tr>
		<tr>
			<td style="text-align: right" colspan="4">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table27">
					<tr>
						<td>
<input type="submit" value="Update" name="B1"></td>
						<td style="text-align: center">
						<input type="submit" value="Print &amp; Close" name="B1"></td>
						<td style="text-align: right"><input type="submit" value="Cancel" name="B1"></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		</table>
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="OID" value="<%=OID%>">
        <input type="hidden" name="Field50" value="<%=PID%>">
</form>
</div>
<%
rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<%Sub CancelUpdate%>
<%PID=Request.form("PID")%>
<%OID=Request.form("OID")%>
<form method="POST" action="Permit_open.asp" name="Update" target="ibody">
        <input type="hidden" name="PID" Value="<%=PID%>">
        <input type="hidden" name="OID" Value="<%=OID%>">
        <input type="hidden" name="Field50" value="<%=PID%>">
        <input type="hidden" name="Field55" value="Issued">
        <input type="hidden" name="Refresh" value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub UpdateData%>
Update Data
<%PID=Request.form("PID")%>
<%OID=Request.form("OID")%>
<%DIM Fields(60)%>
<%Dim FieldName(60)%>
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
<%stp=1%>
<%For I = 1 to 59 step stp%>
<%Data="Field"& I%>
<%Fields(I)=Request.form(Data)%>
<%If I=22 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=23 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=29 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=30 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%if Fields(24)=", None" then Fields(24)=""%>
<%If I=36 then Request.form("Field42") & "-" & Request.form("Field43")%>
<%Fields(I)=Replace(Fields(I),"'","&#039")%>
<%If Fields(20)="SELECT WATER UTILITY" THEN FIELDS(20)=""%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_open Set " & FieldName(I) & " = '" & Fields(I) & "' Where ID=" & OID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%=I%>:<%=FieldName(I)%>:<%=Fields(I)%><br>
<%If I=50 then stp=3%>
<%If I > 53 then stp=1%>
<%next%>
<p align="center"><b>Permit has been updated in the Database.</b>
<form method="POST" action="Permit_open.asp" name="Update1" target="ibody">
        <input type="hidden" name="PID" Value="<%=PID%>">
        <input type="hidden" name="Refresh" value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update1.submit();", 5000); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub GetPermit%>
Get Permit
<%PID=Request.form("PID")%>
<%OID=Request.form("OID")%>
<%DIM Fields(60)%>

<%Dim FieldName(60)%>
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
<%stp=1%>
<%For I = 1 to 58 step stp%>
<%Data="Field"& I%>
<%Fields(I)=Request.form(Data)%>
<%If I=22 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=23 and Fields(I)="" then Fields(I)="01/01/1900"%>

<%If I=29 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=30 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=36 then Fields(I)=Request.form("Field42") & "-" & Request.form("Field43")%>
<%Fields(I)=Replace(Fields(I),"'","&#039")%>
<%If Fields(20)="SELECT WATER UTILITY" THEN FIELDS(20)=""%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_open Set " & FieldName(I) & " = '" & Fields(I) & "' Where ID=" & OID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%=I%>:<%=FieldName(I)%>:<%=Fields(I)%><br>
<%if I=50 then stp=3%>
<%If I > 53 then stp=1%>
<%next%>

	<script language="javascript">
		window.open('Permit_Confirm.asp?OID=<%=OID%>&PID=<%=PID%>','', 'toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=200,left = 250,top = 200');
	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>