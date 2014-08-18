<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Create New Permit</title>
<!-- #include file="calctl.inc" -->
<script language="JavaScript" src="overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%Server.Scripttimeout=900%>
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<%If Request.form("B1")="Open" then InsertData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<%PID=Request.QueryString("PID")%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project,Municipality,PumpStation,Utility,User_Type from tbl_Project where PID=" & PID & ";")
%>
<%do until rstemp.eof %>
<%ProjectTitle=rstemp(0)%>
<%Municipality=rstemp(1)%>
<%PumpStation=rstemp(2)%>
<%Utility=rstemp(3)%>
<%UserType=rstemp(4)%>
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
set rstemp=conntemp.execute("Select SFF from ph_indemn where PID='" & PID & "';")
%>
<%do until rstemp.eof %>
<%SFF=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<div align="center">
<b><font size="4">Permit Worksheet
</font></b>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.Permit_Count.value == "")
  {
    alert("Please enter a value for the \"# of Permits\" field.");
    theForm.Permit_Count.focus();
    return (false);
  }

  if (theForm.Permit_Count.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"# of Permits\" field.");
    theForm.Permit_Count.focus();
    return (false);
  }

  if (theForm.Permit_Count.value.length > 2)
  {
    alert("Please enter at most 2 characters in the \"# of Permits\" field.");
    theForm.Permit_Count.focus();
    return (false);
  }

  var checkOK = "0123456789-,";
  var checkStr = theForm.Permit_Count.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"# of Permits\" field.");
    theForm.Permit_Count.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseInt(allNum);
  if (chkVal != "" && !(prsVal >= "0" && prsVal <= "50"))
  {
    alert("Please enter a value greater than or equal to \"0\" and less than or equal to \"50\" in the \"# of Permits\" field.");
    theForm.Permit_Count.focus();
    return (false);
  }

  if (theForm.Field37.selectedIndex == 0)
  {
    alert("The first \"Municipality\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Field37.focus();
    return (false);
  }

  if (theForm.Field20.selectedIndex == 0)
  {
    alert("The first \"Supplying Water Utility\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Field20.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789-";
  var checkStr = theForm.Field43.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter and digit characters in the \"Map Page\" field.");
    theForm.Field43.focus();
    return (false);
  }

  if (theForm.Field38.selectedIndex == 0)
  {
    alert("The first \"Downstream Pump Station\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Field38.focus();
    return (false);
  }

  var checkOK = "0123456789-.";
  var checkStr = theForm.Field25.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Connection Fee\" field.");
    theForm.Field25.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"Field25\" field.");
    theForm.Field25.focus();
    return (false);
  }

  var checkOK = "0123456789-.";
  var checkStr = theForm.Field26.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Tap Fee\" field.");
    theForm.Field26.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"Field26\" field.");
    theForm.Field26.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="Permit_WorkSheet.asp" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="font-size: 8pt; border-style: solid; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
		<tr>
			<td colspan="5" align="center"><size="2"><b><%=ProjectTitle%></b></size></td>
		</tr>
		<tr>
			<td colspan="4">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2" style="font-family: Tahoma; font-size: 8pt">
					<tr>
						<td style="text-align: right">
						<b># of Permits to Open:</b></td>
						<td width="26">
						<!--webbot bot="Validation" s-display-name="# of Permits" s-data-type="Integer" s-number-separators="," b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2" s-validation-constraint="Greater than or equal to" s-validation-value="0" s-validation-constraint="Less than or equal to" s-validation-value="50" --><input type="text" name="Permit_Count" size="3" title="Enter Total number of Permits to open" maxlength="2"></td>
						<td style="text-align: right"><b>Starting Lot#</b></td>
						<td style="text-align: right">
						<input name="LOT" size="10" style="float: left" title="Enter your starting Lot #"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table10" style="font-size: 8pt">
					<tr>
						<td style="text-align: center">
			<b>Permit Information</b></td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table11" style="font-size: 8pt">
								<tr>
									<td width="92" style="text-align: right">Street#</td>
									<td width="36">
			<input type="text" name="Field4" size="5"></td>
									<td style="text-align: right" width="44">Street</td>
									<td>
			<input type="text" name="Field6" size="29"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table12">
								<tr>
									<td width="92" style="text-align: right">Address 2</td>
									<td>
			<input type="text" name="Field5" size="47"></td>
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
									<td width="92" style="text-align: right">City/State</td>
									<td>
			<input type="text" name="Field7" size="47"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table27">
								<tr>
									<td style="text-align: right" width="92">Zip
									</td>
									<td>
			<input type="text" name="Field8" size="10"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table14">
								<tr>
									<td width="92" style="text-align: right">Applicant Name</td>
									<td>
			<input type="text" name="Field44" size="47"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table15">
								<tr>
									<td width="92" style="text-align: right">Applicant Address</td>
									<td>
			<input type="text" name="Field45" size="47"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table16">
								<tr>
									<td width="92" style="text-align: right">Applicant City</td>
									<td width="108">
			<input type="text" name="Field47" size="20"></td>
									<td width="48" style="text-align: right">State</td>
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
									<td width="60"><input type="text" name="Field49" size="10"></td>
									<td style="text-align: right" width="92">Applicant Phone</td>
									<td><input name="Field46" size="15"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table17">
								<tr>
									<td style="text-align: right">&nbsp;</td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
				</table>
			</div>
</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" colspan="2" width="53%">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table18">
					<tr>
						<td style="text-align: center">
			<b>Owner Information</b></td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table19" style="font-size: 8pt">
								<tr>
									<td width="92" style="text-align: right">Owner Name</td>
									<td>
			<input type="text" name="Field10" size="47"></td>
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
			<input type="text" name="Field11" size="47"></td>
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
									<td width="34">
			<input type="text" name="Field12" size="5"></td>
									<td width="44" style="text-align: right">Street</td>
									<td>
			<input type="text" name="Field13" size="29"></td>
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
									<td width="92" style="text-align: right">&nbsp;Address 2</td>
									<td>
			<input type="text" name="Field14" size="47"></td>
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
									<td width="149"><input name="Field15" size="28"></td>
									<td width="48" style="text-align: right">State</td>
									<td><Select name="Field16" size="1">
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
									<td>
			<input type="text" name="Field17" size="10" style="font-size: 8pt"></td>
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
					<tr>
						<td>&nbsp;</td>
					</tr>
					</table>
			</div></td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4" style="font-size: 8pt">
					<tr>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px; text-align:right" width="92">
						Customer Type</td>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px; text-align:center">
						<%If Instr(UserType,"R") then chk="checked" else chk=""%>
						<input type="checkbox" name="Field18" value="R" <%=chk%>></td>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">Residential</td>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px; text-align:center">
						<%If Instr(UserType,"C") then chk="checked" else chk=""%>
						<input type="checkbox" name="Field18" value="C" <%=chk%>></td>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">Commercial</td>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px; text-align:center">
						<%If Instr(UserType,"I") then chk="checked" else chk=""%>
						<input type="checkbox" name="Field18" value="I" <%=chk%>></td>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">Industrial</td>
					</tr>
				</table>
			</td>
			<td colspan="2" style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="53%">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5" style="font-size: 8pt">
					<tr>
						<td width="135" style="text-align: right; border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">
						Estimated Gallons per Day</td>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" width="30">
						<input type="text" name="Field19" size="3"></td>
						<td style="text-align: right; border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" width="57">
						&nbsp;</td>
						<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">
						&nbsp;</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: right; border-right-width: 1px; border-right-style:solid; border-left-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table24">
					<tr>
						<td style="text-align: right" width="92">
						<font color="#FF0000">* Municipality</font></td>
						<td>&nbsp;<!--webbot bot="Validation" s-display-name="Municipality" b-disallow-first-item="TRUE" --><Select name="Field37" size="1">
<option>SELECT MUNICIPALITY</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Municipality;")
%>
<%do until rstemp.eof %>
<%If Int(Municipality)=Int(rstemp(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp(1)%>" <%=Sel%>><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="2" style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="53%">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table6">
					<tr>
						<td width="135" style="text-align: right">
						<font color="#FF0000">*
						Supplying Water Utility</font></td>
						<td>&nbsp;<!--webbot bot="Validation" s-display-name="Supplying Water Utility" b-disallow-first-item="TRUE" --><Select name="Field20" size="1">
<option>SELECT UTILITY</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Utility;")
%>
<%do until rstemp.eof %>
<%If Int(Utility)=Int(rstemp(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp(1)%>" <%=sel%>><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: right; border-right-width: 1px; border-right-style:solid; border-left-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table25">
					<tr>
						<td style="text-align: right" width="92"># of IDU's</td>
						<td>
			<input type="text" name="Field21" size="3"></td>
					</tr>
				</table>
			</div>
			</td>
			<td colspan="2" style="border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="53%">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table7" style="font-size: 8pt">
					<tr>
						<td width="135" style="text-align: right">
						Contract #</td>
						<td width="41">
						<p style="text-align: center">
						<Select name="Field42" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Contract from tbl_Contract order by Contract ASC;")
%>
<%do until rstemp.eof %>
<option value="<%=rstemp(0)%>"><%=rstemp(0)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
						<td style="text-align: right">
						Page #</td>
						<td width="273">
						&nbsp;<!--webbot bot="Validation" s-display-name="Map Page" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" --><input name="Field43" size="4" style="float: left"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: center; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px"><b>Lateral Information</b></td>
			<td colspan="2" style="text-align: center" width="53%"><b>Inspection Information</b></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table8" style="font-size: 8pt">
					<tr>
						<td width="92" style="text-align: right">
						Length</td>
						<td width="70">
						<input name="Field33" size="10" style="float: left"></td>
						<td style="text-align: right">
						Depth</td>
						<td width="175">
						<input type="text" name="Field34" size="10"></td>
					</tr>
				</table>
			</td>
			<td style="text-align: right" width="10%">Inspector</td>
			<td width="42%">
			<input type="text" name="Field28" size="20"></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table9" style="font-size: 8pt">
					<tr>
						<td width="92" style="text-align: right">
						Lateral Station</td>
						<td width="70">
						<input name="Field35" size="10" style="float: left"></td>
						<td style="text-align: right">
						DS/US Manhole</td>
						<td width="73">
						<input type="text" name="Field41" size="10"></td>
						<td width="87">
						<input type="text" name="Field54" size="10"></td>
					</tr>
				</table>
			</td>
			<td style="text-align: right" width="10%">Inspection Date</td>
			<td width="42%">
			<input type="text" name="Field29" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('FrontPage_Form1.Field29');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			</td>
		</tr>
		<tr>
			<td style="text-align: right; border-right-width: 1px; border-right-style: solid; border-left-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px" width="45%" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table26">
					<tr>
						<td style="text-align: right">
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table30" style="font-size: 8pt">
								<tr>
									<td width="20%">
						<font color="#FF0000">* DS Pump Station</font></td>
									<td width="31%">
									<!--webbot bot="Validation" s-display-name="Downstream Pump Station" b-disallow-first-item="TRUE" --><Select name="Field38" size="1">
<option>SELECT PUMPSTATION</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_PS;")
%>
<%do until rstemp.eof %>
<%If Int(Pumpstation)=Int(rstemp(2)) then sel="selected" else sel=""%>
<option value="<%=rstemp(2)%>" <%=sel%>><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
									<td style="text-align: right" colspan="2">&nbsp;</td>
									<td width="30%" colspan="4">			
									</td>
								</tr>
								<tr>
									<td width="20%">
						Facility Fee</td>
									<td width="31%">
									<Select name="Field53" size="1">
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
									<td width="4%">
									<input type="radio" name="Field56" value="Escrow"></td>
									<td width="14%">
									<p style="text-align: left">Escrow</td>
									<td width="4%">			
									<input type="radio" name="Field56" value="Payment Due"></td>
									<td width="9%">			
									Payment Due</td>
									<td width="8%">			
									<input type="radio" name="Field56" value="None"></td>
									<td width="8%">			
									None</td>
								</tr>
							</table>
						</div>
					</td>
					</tr>
				</table>
			</div>
			</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:right" width="10%">
			Inspection Fee</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="42%">
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
			<td colspan="2" width="53%">
			<p style="text-align: center"><b>Miscellaneous Information</b></td>
		</tr>
		<tr>
			<td style="text-align: right" width="11%">Application Date</td>
			<td style="text-align: left; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px" width="34%">
<input type="text" name="Field22" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('FrontPage_Form1.Field22');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">			
			&nbsp;</td>
			<td style="text-align: right" width="10%">Prepared by</td>
			<td width="42%">
			<input type="text" name="Field27" size="20" value="<%=Request.Cookies("Portal")("FullName")%>"></td>
		</tr>
		<tr>
			<td style="text-align: right" width="11%">Issue Date</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="34%">
			<input type="text" name="Field23" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('FrontPage_Form1.Field23');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			&nbsp;</td>
			<td style="text-align: right" width="10%">Phase</td>
			<td width="42%">
			<input type="text" name="Field31" size="20"></td>
		</tr>
		<tr>
			<td style="text-align: right" width="11%">Billing Date</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="34%">
			<input type="text" name="Field30" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('FrontPage_Form1.Field30');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			&nbsp;</td>
			<td style="text-align: right" width="10%">Section</td>
			<td width="42%">
			<input type="text" name="Field32" size="20"></td>
		</tr>
		<tr>
			<td style="text-align: right" width="11%">Connection Fee</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="34%">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table31">
				<tr>
					<td width="58">
					<!--webbot bot="Validation" s-display-name="Connection Fee" s-data-type="Number" s-number-separators="x." --><input type="text" name="Field25" size="10"></td>
					<td style="text-align: center" width="32">
					<input type="radio" name="Field57" value="Escrow"></td>
					<td width="45">Escrow</td>
					<td style="text-align: center">
					<input type="radio" name="Field57" value="Payment Due"></td>
					<td>Payment Due</td>
					<td>
					<input type="radio" name="Field57" value="None"></td>
					<td>None</td>
				</tr>
			</table>
			</td>
			<td style="text-align: right" width="10%">Development</td>
			<td width="42%">
			<input type="text" name="Field9" size="20"></td>
		</tr>
		<tr>
			<td style="text-align: right" width="11%">Tap Fee</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="34%">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table32">
				<tr>
					<td width="59">
					<!--webbot bot="Validation" s-display-name="Tap Fee" s-data-type="Number" s-number-separators="x." --><input type="text" name="Field26" size="10"></td>
					<td width="32" style="text-align: center">
					<input type="radio" name="Field58" value="Pmt Due from Escrow"></td>
					<td width="46">Escrow</td>
					<td style="text-align: center">
					<input type="radio" name="Field58" value="Payment Due"></td>
					<td>Payment Due</td>
					<td>
					<input type="radio" name="Field58" value="None"></td>
					<td>None</td>
				</tr>
			</table>
			</td>
			<td style="text-align: right" width="10%">Check#</td>
			<td width="42%">
			<input type="text" name="Field40" size="20"></td>
		</tr>
		<tr>
			<td style="text-align: right; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px" width="97%" colspan="4">
			<font color="#FF0000">* Required Fields</font></td>
		</tr>
		</table>
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="Field50" value="<%=PID%>">
<input type="submit" value="Open" name="B1">
	<a href="Project_Add_3.asp?PID=<%=PID%>">Back to Project</a>
</form>
</div>

<%Sub InsertData%>
<%Account=Request.form("Account")%>
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
<%Lot=Request.form("Lot")%>
<%If Lot="" then Lot=0%>
<%PID=Request.form("PID")%>
<%Permit_Count=Request.form("Permit_Count")%>
<%For X = 1 to Permit_Count%>
<%If X <> 1 then Lot=Lot+1%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Insert Into tbl_open (Account,PID,Lot) Values ('" & Account & "','" & PID & "','" & Lot & "')")%>
<%If X = 1 then RS = Conn.Execute("Insert Into tbl_Reservation (Permit,PID) Values ('" & Permit_Count & "','" & PID & "')")%>

<%
set RS=nothing
conn.close
Set conn=nothing
%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Max(ID) from tbl_open;")
%>
<%do until rstemp.eof %>
<%NewID=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%Stp=1%>
<%For I = 4 to 58 Step stp%>
<%Data="Field"& I%>
<%Fields(I)=Request.form(Data)%>
<%Fields(I)=Replace(Fields(I),"'","''")%>
<%If I=22 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=23 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=29 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=30 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If Fields(24)=", None" then Fields(24)=""%>
<%If I=36 then Fields(I)=Request.form("Field42") & "-" & Request.form("Field43")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%If I=22 or I = 23 or I=29 or I=30 then RS = Conn.Execute("Update tbl_open Set " & FieldName(I) & " = " & Fields(I) & " Where ID=" & NewID & ";") else RS = Conn.Execute("Update tbl_open Set " & FieldName(I) & " = '" & Fields(I) & "' Where ID=" & NewID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%=I%>:<%=Fieldname(I)%>:<%=Fields(I)%><br>
<%If I = 50 then Stp=3%>
<%If I > 53 then stp=1%>
<%next%>
<%next%>
<%Response.Cookies("PDM")("MaxID")=NewID%>
<p align="center"><b><%=Permit_Count%> Permits have been successfully opened.</b>
<form method="POST" action="Project_Add_3.asp" name="Update" target="ibody">
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="Refresh" value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<%CancelPID=Request.form("CancelPID")%>
<form method="POST" action="Project_add_3.asp" name="Update" target="ibody">
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="Refresh" value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>