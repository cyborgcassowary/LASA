<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Create New Project Step 3</title>
<!-- #include file="calctl.inc" -->
<script language="JavaScript" src="overlib_mini.js"></script>
<%PID=Request.QueryString("PID")%>
<%If Request.QueryString("PID")="" then PID=Request.Form("PID")%>
<!--#include file="inc/Page3_Queries.asp"-->
<script language="javascript" type="text/javascript">
<!--
function part1()
{
window.open('P_Edit_1.asp?PID=<%=PID%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=510,height=300,left = 250,top = 200');
}
function part2()
{
window.open('P_edit_2.asp?PID=<%=PID%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=550,height=240,left = 250,top = 200');
}
//-->
function assignments()
{
window.open('Assign_popup.asp?ID=<%=AID%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=700,height=500,left = 250,top = 200');
}
//-->
function comments()
{
window.open('Comments_popup.asp?PID=<%=PID%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=350,left = 250,top = 200');
}
//-->
function addassign()
{
window.open('Assign_Add_popup.asp?PID=<%=PID%>&Project=<%=Project%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=300,left = 250,top = 200');
}
//-->
function crp_ws()
{
window.open('CRP_Edit.asp?PID=<%=PID%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=550,height=370,left = 250,top = 200');
}
//-->
function srl_ws()
{
window.open('srl.asp?PID=<%=PID%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=300,left = 250,top = 200');
}
//-->
function table_fix()
{
window.open('Project_Fix.asp','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=300,left = 250,top = 200');
}
function OLS_Edit()
{
window.open('OLS_Edit.asp?PID=<%=PID%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=300,left = 250,top = 200');
}
//-->
</script>
<script type="text/javascript" language="JavaScript1.2" src="inc/stmenu.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->
<%Session("Open_Date")=""%>
<%Session("Comments")=""%>
<%Session("FileCreated")=""%>
<%Session("Request_Rec")=""%>
<%Session("Review_Rec")=""%>
<%Session("Reserve_Date")=""%>
<%Session("Expire_Date")=""%>
<%Session("Renew_date")=""%>
<%Session("Expire_Notice")=""%>
<%Session("Res_Amount")=""%>
<%Session("Res_Type")=""%>
<%Session("Capacity_Close")=""%>
<%session("Record_type")=""%>
<%If Request.form("B1")="Submit" then InsertData%>

<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<form method="POST" action="Project_Add_3.asp" name="project" language="JavaScript" webbot-action="--WEBBOT-SELF--">
<div align="center">

	<table border="0" cellspacing="1" id="table1" style="border: 2px outset #0000FF" width="100%">
		<tr>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td width="27%"><b>Project Name:</b></td>
						<td width="73%"><%=Project%></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3">
					<tr>
						<td width="22%"><b>Municipality:</b></td>
						<td width="78%"><%=M_Name%></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4">
					<tr>
						<td width="27%"><b>Treatment Plant:</b></td>
						<td width="73%"><%=T_Plant%></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5">
					<tr>
						<td width="22%"><b>Pump Station:</b></td>
						<td width="78%"><%=PS_name%></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table6">
					<tr>
						<td width="27%"><b>Water Utility:</b></td>
						<td width="73%"><%=Utility%></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table7">
					<tr>
						<td width="22%"><b>Connection Type:</b></td>
						<td width="78%"><%=C_Type%></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td colspan="6">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table12">
					<tr>
						<td width="127"><b>GIS Parcel ID:</b></td>
						<td width="364"><%=GISID%>&nbsp;</td>
						<td>
						<table border="0" width="100%" id="table13" cellspacing="0" cellpadding="0">
							<tr>
								<%If Res_Type="IDU" then Multiplier=100 else Multiplier=.42%>
								<%SanityCheck=CLng(Trim(Res_Amount)) * Multiplier%>
								<td><b>Original Requested Capcity <%=Res_Type%>:</b>&nbsp;&nbsp;<%=Res_Amount%> X <%=Multiplier%> = <%=FormatCurrency(SanityCheck)%></td>
								<td>&nbsp;</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td colspan="6">
			&nbsp;</td>
		</tr>
		<tr>
			<td colspan="6"><hr width="98%"></td>
		</tr>
		<tr>
			<td style="text-align: center" colspan="3"><b>Developer Information</b></td>
			<td style="text-align: center" colspan="3"><b>Engineer Information</b></td>
		</tr>
		<tr>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table8">
					<tr>
						<td><%=D_Name%></td>
					</tr>
					<tr>
						<td><%=D_Street%></td>
					</tr>
					<tr>
						<td><%=D_City%>, <%=D_State%>  <%=D_Zip%></td>
					</tr>
					<tr>
						<td><%=D_Number%></td>
					</tr>
					<tr>
					<td><b>Project Contact:</b> <%=D_Contact%></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table9">
					<tr>
						<td><%=E_Name%></td>
					</tr>
					<tr>
						<td><%=E_Street%></td>
					</tr>
					<tr>
						<td><%=E_City%>, <%=E_State%>  <%=E_Zip%></td>
					</tr>
					<tr>
						<td><%=E_Number%></td>
					</tr>
					<tr>
					<td><b>Project Contact:</b> <%=E_Contact%></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
		<td colspan="6">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="6" align="center" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">Project inquiries should be directed to the:  <font color="red"><b><%=P_contact%></b></font></td>
		</tr>
		<tr>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table10">
					<tr>
						<td style="text-align: right">
			<b>Phase:</b></td>
						<td>
			<b><select size="1" name="Phase">
			<option value="">Select Phase</option>
			<%if Phase="Capacity" then sel="selected" else sel=""%>
			<option <%=sel%>>Capacity</option>
			<%if Phase="Plan Review" then sel="selected" else sel=""%>
			<option <%=sel%>>Plan Review</option>
			<%if Phase="Approval" then sel="selected" else sel=""%>
			<option <%=sel%>>Approval</option>
			<%if Phase="Construction" then sel="selected" else sel=""%>
			<option <%=sel%>>Construction</option>
			<%if Phase="Closeout" then sel="selected" else sel=""%>
			<option <%=sel%>>Closeout</option>
			</select></b></td>
						<td style="text-align: center">
						<%If AID <> "" and Status = "Completed" then Response.write "<a href=javascript:assignments()>View Assignment</a>" else Response.write "<a href=javascript:addassign()>Add Assignment</a>"%></td>
						<td style="text-align: center">&nbsp;</td>
						<td style="text-align: center"><!--#include file="inc/Project_options.inc"--></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table11">
					<tr>
						<td style="text-align: center">
						&nbsp;</td>
						<td style="text-align: center">&nbsp;</td>
						<td style="text-align: center">&nbsp;</td>
						<td style="text-align: right">&nbsp;</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>

		<tr>
			<td colspan="6" style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">&nbsp;</td>
		</tr>

		<tr>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px" colspan="6"><font color="#FF0000"><b>
			Check All Applicable Requirements for this Project</b></font></td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<b>REQ</b></td>
			<td style="text-align: center" colspan="2" bgcolor="#C0C0C0"><!--#include file="inc/ph_capacity.asp"-->
			<b>Capacity Phase</b></td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			<b>REQ</b></td>
			<td style="text-align: center" colspan="2" bgcolor="#999999"><!--#include file="inc/ph_board.asp"-->
			<b>Board Approval</b></td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Reservation Required")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Reservation Required" <%=chk%>></td>
			<td nowrap colspan="2" bgcolor="#C0C0C0">Reservation Required</td>
			<td nowrap style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			<%if Int(Instr(REQ,"Financial Security Required")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Financial Security Required" <%=chk%>></td>
			<td style="text-align: left" colspan="2" bgcolor="#999999">
			Financial Security Required</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="SRF" size="10" value="<%=SRF%>">
<a href="javascript:ggPosX=200;ggPosY=330;name=1;show_calendar('project.SRF');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a>
			</td>
			<td bgcolor="#C0C0C0">Sewer Request&nbsp; Received ($225)</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			&nbsp;</td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="CEA" size="10" value="<%=CEA%>">
<a href="javascript:name=2;show_calendar('project.CEA');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a>
</td>
			<td bgcolor="#999999">Construction Estimate Approval</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="LOCA" size="10" value="<%=LOCA%>">
<a href="javascript:show_calendar('project.LOCA');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Letter of Capacity Approval</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			&nbsp;</td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="FSA" size="15" value="<%=FSA%>"></td>
			<td bgcolor="#999999">Financial Security Amount
			<font color="#FF0000">($)</font></td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="ROC" size="10" value="<%=ROC%>">
<a href="javascript:show_calendar('project.ROC');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Reservation of Capacity</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			&nbsp;</td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="FSR" size="10" value="<%=FSR%>">
<a href="javascript:show_calendar('project.FSR');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Financial Security Received</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="RFA" size="15" value="<%=RFA%>"></td>
			<td bgcolor="#C0C0C0">Reservation Fee Amount <font color="#FF0000">
			($)</font></td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			&nbsp;</td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="FSE" size="10" value="<%=FSE%>">
<a href="javascript:show_calendar('project.FSE');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Financial Security Expiration</td>
		</tr>
		<tr>
			<td style="text-align: center" colspan="3" bgcolor="#999999"><!--#include file="inc/ph_development.asp"-->
			<b>Development Review</b></td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			<%if Int(Instr(REQ,"Builder&#039s Agreement")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Builder's Agreement" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="BA" size="10" value="<%=BA%>">
<a href="javascript:show_calendar('project.BA');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Builder's Agreement Received</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#999999">
			<%if Int(Instr(REQ,"Development Review Fee")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Development Review Fee" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="DRF" size="10" value="<%=DRF%>">
<a href="javascript:show_calendar('project.DRF');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Development Review Fee Received</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			<%if Int(Instr(REQ,"Approved Contractor")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Approved Contractor" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="AC" size="10" value="<%=AC%>">
<a href="javascript:show_calendar('project.AC');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Approved Contractor</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#999999">
			<%if Int(Instr(REQ,"Final Plan Approval")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Final Plan Approval" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="FPA" size="10" value="<%=FPA%>">
<a href="javascript:show_calendar('project.FPA');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Final Plan Approval</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			<%if Int(Instr(REQ,"Road Opening Permit")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Road Opening Permit" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="ROP" size="10" value="<%=ROP%>">
<a href="javascript:show_calendar('project.ROP');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Road Opening Permit</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#999999">
			<%if Int(Instr(REQ,"Planning Module Approved")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Planning Module Approved" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="PMA" size="10" value="<%=PMA%>">
<a href="javascript:show_calendar('project.PMA');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Planning Module Approved</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			<%if Int(Instr(REQ,"ACT 14 Notification")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="ACT 14 Notification" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="ACT" size="10" value="<%=ACT%>">
<a href="javascript:show_calendar('project.ACT');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">ACT 14 Notification</td>
		</tr>
		<tr>
			<td style="text-align: center" colspan="3" bgcolor="#C0C0C0"><!--#include file="inc/ph_construction.asp"-->
			<b>Construction Phase</b></td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			<%if Int(Instr(REQ,"PA DEP Approvals")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="PA DEP Approvals" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="DEP" size="10" value="<%=DEP%>">
<a href="javascript:show_calendar('project.DEP');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">PA DEP Approvals</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Approved Construction Permit")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Approved Construction Permit" <%=chk%>></td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="ACP" size="10" value="<%=ACP%>">
<a href="javascript:show_calendar('project.ACP');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">App. for Construction Received ($100)</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">
			<%if Int(Instr(REQ,"Board Approval")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Board Approval" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="BAD" size="10" value="<%=BAD%>">
<a href="javascript:show_calendar('project.BAD');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Board Approval Date</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Construction Schedule Received")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Construction Schedule Received" <%=chk%>></td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="CSR" size="10" value="<%=CSR%>">
<a href="javascript:show_calendar('project.CSR');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Construction Schedule Received</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">&nbsp;</td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="EFN" size="15" value="<%=EFN%>"></td>
			<td bgcolor="#999999">Escrow File Number</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<input type="checkbox" name="REQ" value="Construction Permit Number" <%=chk%>></td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="CPN" size="15" value="<%=CPN%>"></td>
			<td bgcolor="#C0C0C0">Construction Permit Number</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">&nbsp;</td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="EFR" size="15" value="<%=EFR%>"></td>
			<td bgcolor="#999999">Escrow Funds Required</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Construction Permit Number")) > 0 then chk="checked" else chk=""%>
			</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="CPD" size="10" value="<%=CPD%>">
<a href="javascript:show_calendar('project.CPD');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Construction Permit Date</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#999999">&nbsp;</td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="EFRc" size="10" value="<%=EFRC%>">
<a href="javascript:show_calendar('project.EFRc');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Escrow Funds Received</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Rights of Way Received")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Rights of Way Received" <%=chk%>></td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="ROWR" size="10" value="<%=ROWR%>">
<a href="javascript:show_calendar('project.ROWR');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Rights of Way Received</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" colspan="3" bgcolor="#C0C0C0"><!--#include file="inc/ph_indemn.asp"-->
			<b>Dedication</b></td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Shop Drawing Approval")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Shop Drawing Approval" <%=chk%>></td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="SDA" size="10" value="<%=SDA%>">
<a href="javascript:show_calendar('project.SDA');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">All Shop Drawing Approved</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Indemnification Needed")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Indemnification Needed" <%=chk%>></td>
			<td nowrap colspan="2" bgcolor="#C0C0C0">Indemnification Needed</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			<input type="text" name="Copies" size="1" value="<%=Copies%>"></td>
			<td nowrap bgcolor="#C0C0C0"><input type="text" name="CopyDate" size="10" value="<%=CopyDate%>">
<a href="javascript:show_calendar('project.CopyDate');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a>
			&nbsp;</td>
			<td bgcolor="#C0C0C0">Additional Construction Drawings Received</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<select size="1" name="TOI">
			<%If TOI="" then sel="selected" else sel=""%>
			<option <%=sel%>></option>
			<%If TOI="N/A" then sel="selected" else sel=""%>
			<option <%=sel%>>N/A</option>
			<%If TOI="State Road" then sel="selected" else sel=""%>
			<option <%=sel%>>State Road</option>
			<%If TOI="Township Road" then sel="selected" else sel=""%>
			<option <%=sel%>>Township Road</option>
			<%If TOI="Standard" then sel="selected" else sel=""%>
			<option <%=sel%>>Standard</option>
			</select></td>
			<td bgcolor="#C0C0C0">Type of Indemnification</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0"><input type="text" name="HOP_APP" size="20" value="<%=HOP_APP%>"></td>
			<td bgcolor="#C0C0C0">HOP Application #</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="IAR" size="10" value="<%=IAR%>"><a href="javascript:show_calendar('project.IAR');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Indemnification Agreement Received</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0"><input type="text" name="HOP_Permit" size="20" value="<%=HOP_Permit%>"></td>
			<td bgcolor="#C0C0C0">HOP Permit #</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="ICR" size="10" value="<%=ICR%>"><a href="javascript:show_calendar('project.ICR');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Insurance Certificate Received</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#C0C0C0">
			&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="HOP_EXP" size="10" value="<%=HOP_EXP%>"><a href="javascript:show_calendar('project.HOP_EXP');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">HOP Expiration Date</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			<input type="text" name="IED" size="10" value="<%=IED%>"><a href="javascript:show_calendar('project.IED');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Insurance Expiration Date</td>
		</tr>
		<tr>
			<td style="text-align: center" colspan="3" bgcolor="#999999"><!--#include file="inc/ph_close.asp"-->
			<b>Project Closeout</b></td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			&nbsp;<Select name="SFF" size="1">
<option value="0">None</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Facility order by FacilityName ASC;")
%>
<%do until rstemp.eof %>
<%If Int(SFF)=Int(rstemp(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%> - $<%=rstemp(2)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
			<td bgcolor="#C0C0C0">Special Facility Fees</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#999999">
			<%if Int(Instr(REQ,"Final Inspection")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Final Inspection" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="FI" size="10" value="<%=FI%>">
<a href="javascript:show_calendar('project.FI');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Final Inspection</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Deed of Dedication")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Deed of Dedication" <%=chk%>></td>
			<td nowrap bgcolor="#C0C0C0">
			&nbsp;<input type="text" name="DODD" size="10" value="<%=DODD%>"><a href="javascript:show_calendar('project.DODD');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">Deed of Dedication Date</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#999999">
			<%if Int(Instr(REQ,"Escrow Refunded")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Escrow Refunded" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="ER" size="10" value="<%=ER%>">
<a href="javascript:show_calendar('project.ER');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Escrow Refunded</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"As Builts")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="As Builts" <%=chk%>></td>
			<td nowrap bgcolor="#C0C0C0">
			&nbsp;<input type="text" name="ABR" size="10" value="<%=ABR%>"><a href="javascript:show_calendar('project.ABR');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#C0C0C0">As Built Received</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#999999">
			<%if Int(Instr(REQ,"Project Complete & Closed")) > 0 then chk="checked" else chk=""%>
			<input type="checkbox" name="REQ" value="Project Complete &amp; Closed" <%=chk%>></td>
			<td nowrap bgcolor="#999999">
			<input type="text" name="PCC" size="10" value="<%=PCC%>">
<a href="javascript:show_calendar('project.PCC');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
			<td bgcolor="#999999">Project Complete and Closed</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">&nbsp;</td>
			<td nowrap bgcolor="#C0C0C0">
			&nbsp;</td>
			<td bgcolor="#C0C0C0">&nbsp;</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#999999">&nbsp;</td>
			<td nowrap bgcolor="#999999">&nbsp;</td>
			<td bgcolor="#999999">&nbsp;</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"Deed of Dedication")) > 0 then chk="checked" else chk=""%>
			</td>
			<td nowrap bgcolor="#C0C0C0">
			&nbsp;</td>
			<td bgcolor="#C0C0C0">&nbsp;</td>
		</tr>
		<tr>
			<td style="text-align: center" bgcolor="#999999">&nbsp;</td>
			<td nowrap bgcolor="#999999">&nbsp;</td>
			<td bgcolor="#999999">&nbsp;</td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" bgcolor="#C0C0C0">
			<%if Int(Instr(REQ,"As Builts")) > 0 then chk="checked" else chk=""%>
			</td>
			<td nowrap bgcolor="#C0C0C0">
			&nbsp;</td>
			<td bgcolor="#C0C0C0">&nbsp;</td>
		</tr>
		<tr>
		<td colspan="6">
		<table width="100%">
		<tr>
			<td width="33%"><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></td>
			<td width="33%" style="text-align: center">
			<a href="Project_Index.asp">Back to Project Index</a></td>
			<td width="33%"></td>
		</tr>
		</table>
		</td>
		</tr>
		</table>
</div>
<input type="hidden" name="PID" value="<%=Session("ProjectID")%>">
</form>

<%Sub InsertData%>
<%PID=Request.form("PID")%>
<%REQ=Request.form("REQ")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%SRF=Request.form("SRF")%>
<%If NOT IsDate(SRF) then SRF="1/1/1990"%>
<%LOCA=Request.form("LOCA")%>
<%If NOT IsDate(LOCA) then LOCA="1/1/1990"%>
<%ROC=Request.form("ROC")%>
<%If NOT IsDate(ROC) then ROC="1/1/1990"%>
<%RFA=Request.form("RFA")%>
<%RFA=Replace(RFA,"$","")%>
<%RS = Conn.Execute("Update ph_capacity Set SRF='" & SRF & "', LOCA='" & LOCA & "', ROC='" & ROC & "', RFA='" & RFA & "' Where PID='" & PID & "';")%>
<%DRF=Request.form("DRF")%>
<%If NOT IsDate(DRF) then DRF="1/1/1990"%>
<%FPA=Request.form("FPA")%>
<%If NOT IsDate(FPA) then FPA="1/1/1990"%>
<%PMA=Request.form("PMA")%>
<%If NOT IsDate(PMA) then PMA="1/1/1990"%>
<%RS = Conn.Execute("Update ph_development Set DRF='" & DRF & "', FPA='" & FPA & "', PMA='" & PMA & "' Where PID='" & PID & "';")%>
<%ACP=Request.form("ACP")%>
<%If NOT IsDate(ACP) then ACP="1/1/1990"%>
<%CSR=Request.form("CSR")%>
<%If NOT IsDate(CSR) then CSR="1/1/1990"%>
<%CPN=Request.form("CPN")%>
<%CPD=Request.form("CPD")%>
<%If NOT IsDate(CPD) then CPD="1/1/1990"%>
<%ROWR=Request.form("ROWR")%>
<%If NOT IsDate(ROWR) then ROWR="1/1/1990"%>
<%SDA=Request.form("SDA")%>
<%If NOT IsDate(SDA) then SDA="1/1/1990"%>
<%Copies=Request.form("Copies")%>
<%CopyDate=Request.form("CopyDate")%>
<%If NOT IsDate(CopyDate) then CopyDate=("1/1/1990")%>
<%HOP_APP=Request.form("HOP_APP")%>
<%HOP_Permit=Request.form("HOP_Permit")%>
<%HOP_EXP=Request.form("HOP_EXP")%>
<%If NOT IsDate(HOP_EXP) then CopyDate=("1/1/1990")%>
<%RS = Conn.Execute("Update ph_construction Set ACP='" & ACP & "', CSR='" & CSR & "', CPN='" & CPN & "', CPD='" & CPD & "', ROWR='" & ROWR & "', SDA='" & SDA & "', Copies='" & Copies & "', CopyDate='" & CopyDate & "', HOP_APP='" & HOP_APP & "', HOP_Permit='" & HOP_Permit & "', HOP_EXP='" & HOP_EXP & "' Where PID='" & PID & "';")%>
<%FI=Request.form("FI")%>
<%If NOT IsDate(FI) then FI="1/1/1990"%>
<%ER=Request.form("ER")%>
<%If NOT IsDate(ER) then ER="1/1/1990"%>
<%PCC=Request.form("PCC")%>
<%If NOT IsDate(PCC) then PCC="1/1/1990"%>
<%RS = Conn.Execute("Update ph_close Set FI='" & FI & "', ER='" & ER & "', PCC='" & PCC & "' Where PID='" & PID & "';")%>
<%CEA=Request.form("CEA")%>
<%If NOT IsDate(CEA) then CEA="1/1/1990"%>
<%FSA=Request.form("FSA")%>
<%FSA=Replace(FSA,"$","")%>
<%FSR=Request.form("FSR")%>
<%If NOT IsDate(FSR) then FSR="1/1/1990"%>
<%FSE=Request.form("FSE")%>
<%IF NOT IsDate(FSE) then FSE="1/1/1990"%>
<%BA=Request.form("BA")%>
<%If NOT IsDate(BA) then BA="1/1/1990"%>
<%AC=Request.form("AC")%>
<%If NOT IsDate(AC) then AC="1/1/1990"%>
<%ROP=Request.form("ROP")%>
<%If NOT IsDate(ROP) then ROP="1/1/1990"%>
<%ACT=Request.form("ACT")%>
<%If NOT IsDate(ACT) then ACT="1/1/1990"%>
<%DEP=Request.form("DEP")%>
<%If NOT IsDate(DEP) then DEP="1/1/1990"%>
<%BAD=Request.form("BAD")%>
<%If NOT IsDate(BAD) then BAD="1/1/1990"%>
<%EFR=Request.form("EFR")%>
<%EFRC=Request.form("EFRC")%>
<%If NOT IsDate(EFRC) then EFRC="1/1/1990"%>
<%RS = Conn.Execute("Update ph_board Set CEA='" & CEA & "', FSA='" & FSA & "', FSR='" & FSR & "', FSE='" & FSE & "', BA='" & BA & "', AC='" & AC & "', ROP='" & ROP & "', ACT='" & ACT & "', DEP='" & DEP & "', BAD='" & BAD & "', EFR='" & EFR & "', EFRC='" & EFRC & "' Where PID='" & PID & "';")%>
<%TOI=Request.form("TOI")%>
<%IAR=Request.form("IAR")%>
<%If NOT IsDate(IAR) then IAR="1/1/1990"%>
<%ICR=Request.form("ICR")%>
<%If NOT IsDate(ICR) then ICR="1/1/1990"%>
<%IED=Request.form("IED")%>
<%If NOT IsDate(IED) then IED="1/1/1990"%>
<%SFF=Request.form("SFF")%>
<%DODD=Request.form("DODD")%>
<%If NOT IsDate(DODD) then DODD="1/1/1990"%>
<%ABR=Request.form("ABR")%>
<%If NOT IsDate(ABR) then ABR="1/1/1990"%>
<%RS = Conn.Execute("Update ph_indemn Set TOI='" & TOI & "', IAR='" & IAR & "', ICR='" & ICR & "', IED='" & IED & "', SFF='" & SFF & "', DODD='" & DODD & "', ABR='" & ABR & "' Where PID='" & PID & "';")%>
<%REQ=Request.form("REQ")%>
<%EFN=Request.form("EFN")%>
<%REQ=Replace(REQ,"'","&#039")%>
<%LastUpdate=FormatDateTime(now)%>
<%Phase=Request.form("Phase")%>
<%RS = Conn.Execute("Update tbl_Project Set Required='" & REQ & "', EscrowFile='" & EFN & "', LastUpdate='" & LastUpdate & "', Phase='" & Phase & "' Where PID=" & PID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>

<%End Sub%>

</body>

</html>