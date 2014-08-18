<%response.buffer=False%>
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
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>

<%X=Request.form("X")%>
<%If X="" then X=Request.QueryString("X")%>
<%If Request.form("B1")="Update" then Updatedata%>
<%If Request.form("B1")="Cancel" then CancelUpdate%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_permit where ID=" & X & ";")
%>
<%SFFeeID=rstemp(53)%>
<%US_Manhole=rstemp(54)%>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

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
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="Permit_Edit.asp" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" webbot-action="--WEBBOT-SELF--">
	<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="font-size: 8pt; border-style: solid; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
		<tr>
			<td colspan="4">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td width="47"><select size="1" name="Field55">
						<%if rstemp(55)="Issued" then sel="selected" else sel=""%>
						<option <%=sel%>>Issued</option>
						<%if rstemp(55)="Not-Connected" then sel="selected" else sel=""%>
						<option <%=sel%>>Not-Connected</option>
						<%if rstemp(55)="Void" then sel="selected" else sel=""%>
						<option <%=sel%>>Void</option>
						</select>						
						</td>
						<td><%=rstemp(0)%>&nbsp;</td>
						<td style="text-align: right" width="147"><b>Account #</b></td>
						<td> <input type="text" name="Field1" size="20" value="<%=rstemp(1)%>"></td>
						<td style="text-align: right"><b>Project</b></td>
						<td>
<Select name="Field50" size="1">
<option value="0">None</option>
<% 
set PIDtemp=server.createobject("adodb.connection")
PIDtemp.open "DSN=PDM;"
set Ptemp=PIDtemp.execute("Select PID,Project from tbl_Project order by Project ASC;")
%>
<%do until Ptemp.eof %>
<%If Int(rstemp(50))=Int(Ptemp(0)) then sel="selected" else sel=""%>
<option value="<%=Ptemp(0)%>" <%=Sel%>><%=PTemp(1)%></option>
<%
Ptemp.movenext
loop

Ptemp.close
set Ptemp=nothing
PIDtemp.close
set PIDtemp=nothing
%>
</Select></td>
						<td style="text-align: right" width="262"><b>Permit #</b></td>
						<td width="17"><%=rstemp(2)%>&nbsp;</td>
						<td style="text-align: right" width="63"><b>Lot#</b></td>
						<td style="text-align: right" width="61">
						<input type="text" name="Field3" size="10" value="<%=rstemp(3)%>"></td>
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
			<td width="44%" colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table10">
					<tr>
						<td width="92" style="text-align: right">Street#</td>
						<td width="35"><input type="text" name="Field4" size="5" value="<%=rstemp(4)%>"></td>
						<td style="text-align: right" width="44">Street</td>
						<td>
						<input type="text" name="Field6" size="29" value="<%=rstemp(6)%>"></td>
					</tr>
				</table>
			</td>
			<td width="47%" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table14">
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
			<td width="48%" colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table11">
					<tr>
						<td width="92" style="text-align: right">Address 2</td>
						<td><input type="text" name="Field5" size="47" value="<%=rstemp(5)%>"></td>
					</tr>
				</table>
			</div>
			</td>
			<td width="47%" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table15">
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
			<td width="48%" colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table12">
					<tr>
						<td width="92" style="text-align: right">City/State</td>
						<td><input type="text" name="Field7" size="47" value="<%=rstemp(7)%>"></td>
					</tr>
				</table>
			</div>
			</td>
			<td width="47%" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table16">
					<tr>
						<td width="92" style="text-align: right">Street#</td>
						<td width="35">
			<input type="text" name="Field12" size="5" value="<%=rstemp(12)%>"></td>
						<td style="text-align: right" width="44">Street</td>
						<td>
			<input type="text" name="Field13" size="29" value="<%=rstemp(13)%>"></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table13">
					<tr>
						<td width="92" style="text-align: right">Zip</td>
						<td><input type="text" name="Field8" size="10" value="<%=rstemp(8)%>"></td>
					</tr>
				</table>
			</div>
			</td>
			<td width="49%" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table17">
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
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table19">
					<tr>
						<td width="92" style="text-align: right">Applicant Name</td>
						<td>
						<input type="text" name="Field44" size="47" value="<%=rstemp(44)%>"></td>
					</tr>
				</table>
			</div>
			</td>
			<td width="49%" colspan="2">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3" style="font-size: 8pt">
					<tr>
						<td width="92" style="text-align: right">City</td>
						<td width="154"><input name="Field15" size="28" value="<%=rstemp(15)%>"></td>
						<td style="text-align: right" width="48">State</td>
						<td>
<Select name="Field16" size="1">
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=Sitesettings;"
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
			</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table19">
					<tr>
						<td style="text-align: right" width="92">Applicant Address</td>
						<td>
						<input type="text" name="Field45" size="47" value="<%=rstemp(45)%>"></td>
					</tr>
				</table></td>
			<td width="49%" colspan="2">
				<div align="center">
					<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table18">
						<tr>
							<td width="92" style="text-align: right">Zip</td>
							<td>
			<input type="text" name="Field17" size="10" value="<%=rstemp(17)%>"></td>
						</tr>
					</table>
				</div>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table19">
					<tr>
						<td style="text-align: right" width="92">Applicant City</td>
						<td width="150">
						<input type="text" name="Field47" size="28" value="<%=rstemp(47)%>"></td>
						<td width="48" style="text-align: right">State</td>
						<td>
<Select name="Field48" size="1">
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=Sitesettings;"
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
				</table></td>
			<td width="49%" colspan="2">
				<div align="center">
					<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table23" style="font-size: 8pt">
						<tr>
							<td width="91" style="text-align: right">Tax Parcel 
							ID</td>
							<td>
							<input type="text" name="ParcelID" size="20" value="<%=rstemp(52)%>"></td>
						</tr>
					</table>
				</div>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table19">
					<tr>
						<td style="text-align: right" width="92">Applicant Zip</td>
						<td width="62">
						<input type="text" name="Field49" size="10" value="<%=rstemp(49)%>"></td>
						<td style="text-align: right" width="92">
						Applicant Phone</td>
						<td>
						<input type="text" name="Field46" size="15" value="<%=rstemp(46)%>"></td>
					</tr>
				</table></td>
			<td width="49%" colspan="2">
				&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; border-right-style:solid">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table19">
					<tr>
						<td style="text-align: right">&nbsp;</td>
					</tr>
				</table></td>
			<td width="49%" colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4" style="font-size: 8pt">
					<tr>
						<td style="text-align: right" width="92">
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
						<td width="27">
						<input type="text" name="Field19" size="3" value="<%=rstemp(19)%>"></td>
						<td style="text-align: right">
						Map Page</td>
						<td>
						<input type="text" name="Field36" size="12" value="<%=rstemp(36)%>"></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: right; border-left-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table20">
					<tr>
						<td style="text-align: right" width="92">Municipality</td>
						<td>
			<Select name="Field37" size="1">
<% 
set mtemp=server.createobject("adodb.connection")
mtemp.open "DSN=PDM;"
set mrstemp=mtemp.execute("Select Municipality from tbl_Municipality;")
%>
<%do until mrstemp.eof %>
<%If mrstemp(0)=rstemp(37) then sel="selected" else sel=""%>
<option value="<%=mrstemp(0)%>" <%=sel%>><%=mrstemp(0)%></option>
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
set ustemp=utemp.execute("Select Utility from tbl_utility;")
%>
<%do until ustemp.eof %>
<%If rstemp(20)=ustemp(0) then sel="selected" else sel=""%>
<option value="<%=ustemp(0)%>" <%=sel%>><%=ustemp(0)%></option>
<%
ustemp.movenext
loop

ustemp.close
set ustemp=nothing
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
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table21">
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
						<p style="text-align: center">
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
						<td style="text-align: right">
						Page #</td>
						<td width="185">
						&nbsp;<!--webbot bot="Validation" s-display-name="Map Page" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" --><input name="Field43" size="4" value="<%=rstemp(43)%>" style="float: left"></td>
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
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table8" style="font-size: 8pt">
					<tr>
						<td width="92" style="text-align: right">
						Length</td>
						<td width="64">
						<input name="Field33" size="10" style="float: left" value="<%=rstemp(33)%>"></td>
						<td style="text-align: right">
						Depth</td>
						<td width="144">
						<input type="text" name="Field34" size="10" value="<%=rstemp(34)%>"></td>
					</tr>
				</table>
			</td>
			<td style="text-align: right">Inspector</td>
			<td>
			<input type="text" name="Field28" size="20" value="<%=rstemp(28)%>"></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right; border-left-width:1px; border-right-style:solid; border-right-width:1px; border-top-width:1px; border-bottom-width:1px">
				&nbsp;</td>
			<td style="text-align: right" colspan="2">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table27" style="font-size: 8pt">
					<tr>
						<td style="text-align: right" width="92">
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
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table9" style="font-size: 8pt">
					<tr>
						<td width="92" style="text-align: right">
						Lateral Station</td>
						<td width="64">
						<input name="Field35" size="10" value="<%=rstemp(35)%>" style="float: left"></td>
						<td style="text-align: right">
						DS/US Manhole</td>
						<td width="72">
						<input type="text" name="Field41" size="10" value="<%=rstemp(41)%>"></td>
						<td width="72">
						<input type="text" name="Field54" size="10" value="<%=US_Manhole%>"></td>
					</tr>
				</table>
			</td>
			<td style="text-align: right">Inspection Date</td>
			<%If rstemp(29)="1/1/1900" then Field29="" else Field29=rstemp(29)%>
			<td><input type="text" name="Field29" value="<%=Field29%>" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('Permit.Field29');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			</td>
		</tr>
		<tr>
			<td style="text-align: right; border-left-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px" colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 8pt" id="table22">
					<tr>
						<td style="text-align: right">
						<div align="center">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table24" style="font-size: 8pt; font-family: Tahoma">
								<tr>
									<td width="93">DS Pump Station</td>
									<td width="118">
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
									<td style="text-align: right" colspan="3">&nbsp;</td>
									<td width="145" colspan="3">&nbsp;</td>
								</tr>
								<tr>
									<td width="93">Facility Fee</td>
									<td width="118">
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
									<td style="text-align: left">
									<%If rstemp(56)="Escrow" then Sel="checked" else sel=""%>
									<input type="radio" name="Field56" value="Escrow" <%=sel%>></td>
									<td style="text-align: left">Escrow</td>
									<td style="text-align: left">
									<%If rstemp(56)="Payment Due" then sel="checked" else sel=""%>
									<input type="radio" name="Field56" value="Payment Due" <%=sel%>></td>
									<td style="text-align: left">Payment Due</td>
									<td style="text-align: left">
									<%if rstemp(56)="None" then sel="checked" else sel=""%>
									<input type="radio" name="Field56" value="None" <%=sel%>></td>
									<td style="text-align: left">None</td>
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
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:right">
			Inspection Fee</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<table width="100%" style="font-size: 8pt; font-family: Tahoma">
			<tr>
				<td>
				<input type="text" name="Field24" size="20" value="<%=InsFee%>" style="height: 19px"></td>
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
		</table></td>
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
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('Permit.Field22');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
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
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('Permit.Field23');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
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
			<input type="text" name="Field30" value="<%=Field30%>" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('Permit.Field30');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top">
			&nbsp;</td>
			<td style="text-align: right">Section</td>
			<td>
			<input type="text" name="Field32" size="20" value="<%=rstemp(32)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right">Connection Fee</td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table25">
				<tr>
					<td width="61">
					<!--webbot bot="Validation" s-display-name="Connection Fee" s-data-type="Number" s-number-separators="x." -->
					<input type="text" name="Field25" size="10" value="<%=rstemp(25)%>" style="height: 19px"></td>
					<td width="23">
					<%if rstemp(57)="Escrow" then sel="checked" else sel=""%>
					<input type="radio" name="Field57" value="Escrow" <%=sel%>></td>
					<td width="48">Escrow</td>
					<td width="23">
					<%If rstemp(25)="Payment Due" then sel="checked" else sel=""%>
					<input type="radio" name="Field57" value="Payment Due" <%=sel%>></td>
					<td width="85">Payment Due</td>
					<td width="25">
					<%If rstemp(25)="None" then sel="checked" else sel=""%>
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
			<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table26">
				<tr>
					<td width="61">
					<!--webbot bot="Validation" s-display-name="Tap Fee" s-data-type="Number" s-number-separators="x." --><input type="text" name="Field26" size="10" value="<%=rstemp(26)%>"></td>
					<td width="22">
					<%If rstemp(58)="Escrow" then sel="checked" else sel=""%>
					<input type="radio" name="Field58" value="Pmt Due from Escrow" <%=sel%>></td>
					<td width="49">Escrow</td>
					<td width="22">
					<%If rstemp(58)="Payment Due" then sel="checked" else sel=""%>
					<input type="radio" name="Field58" value="Payment Due" <%=sel%>></td>
					<td width="86">Payment Due</td>
					<td width="25">
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
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px" colspan="2">
<input type="submit" value="Update" name="B1"></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px" colspan="2"><input type="submit" value="Cancel" name="B1"></td>
		</tr>
		</table>
	</div>
<input type="hidden" name="X" value="<%=X%>">
<input type="hidden" name="Field2" value="<%=rstemp(2)%>">
</form>

<%
rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<%Sub CancelUpdate%>
<%X=Request.form("X")%>
<p align="center"><b>Changes have been Cancelled.</b>
<form method="POST" action="Permit_Detail.asp" name="Update" target="ibody">
        <input type="hidden" name="X" Value="<%=X%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub UpdateData%>
<%X=Request.form("X")%>
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
<%Stp = 1%>
<%For I = 1 to 59 step stp%>
<%Data="Field"& I%>
<%Fields(I)=Request.form(Data)%>
<%If I=22 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=23 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=29 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If I=30 and Fields(I)="" then Fields(I)="01/01/1900"%>
<%If Instr(Fields(I),"LINEN") then Fields(I)=Replace(Fields(I),"WA","")%>
<%If Fields(24)=", None" then Fields(24)=""%>
<%Fields(I)=Replace(Fields(I),"'","''")%>
<%If Fields(20)="SELECT WATER UTILITY" THEN FIELDS(20)=""%>
<%If I=29 or I=22 or I=23 or I=30 then Fields(I)=Cdate(Fields(I))%>
<%=I%>:<%=FieldName(I)%>:<%=Fields(i)%><br>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%Faketime=" 00:00:00"%>
<%If I=22 or I=23 or I=29 or I=30 then
	RS = Conn.Execute("Update tbl_Permit Set " & FieldName(I) & " = '" & Cdate(Fields(I)) & "' Where ID=" & X & ";")

	else
	 RS = Conn.Execute("Update tbl_Permit Set " & FieldName(I) & " = '" & Fields(I) & "' Where ID=" & X & ";")
End if%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%If I=50 then stp=3%>
<%If I > 53 then stp=1%>
<%next%>
<%ParcelID=Request.form("ParcelID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Permit Set ParcelID='" & ParcelID & "' Where ID=" & X & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
52:<%=ParcelID%>
<p align="center"><b>Permit has been updated in the Database.</b>
<form method="POST" action="Permit_Detail.asp" name="Update" target="ibody">
        <input type="hidden" name="X" Value="<%=X%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>

</body>

</html>