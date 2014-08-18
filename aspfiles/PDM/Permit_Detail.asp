<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Permit Detail</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="Edit" then PermitEdit%>
<%If Request.form("B1")="Add" then PermitAdd%>
<%If Request.form("B1")="Permit" then GetPermit%>
<%If Request.form("B1")="Refresh" then GetMaxID%>
<%If Request.form("X")="" then X=Request.QueryString("X") else X=Request.Form("X")%>
<%If Request.Cookies("PDM")("MaxID")="" Or X=0 then GetMaxID%>
<%If X=0 then X=1%>
<%If Request.form("B1")="Next" then X=X+1%>
<%If Request.form("B1")="Previous" then X=X-1%>
<%If Request.form("B1")="Go" then PermitID%>
<%If Request.form("B2")="Go" then AccountID%>
<%If Request.form("B1")="Last" then X=Request.Cookies("PDM")("MaxID")%>
<%If Request.form("B1")="First" then X=1%>
<%If X <= 0 then X=1%>
<%If Int(X) > Int(request.cookies("PDM")("MaxID")) then X=Int(Request.cookies("PDM")("MaxID"))%>

<%if Request.QueryString("OLS") <> "" then GetID%>

<%Sub GetID%>
<%Permit=Request.QueryString("OLS")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select ID from tbl_permit where Permit=" & Permit & ";")
%>
<%X=rstemp(0)%>
<%
rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%End Sub%>


<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.AccountVal.value.length > 6)
  {
    alert("Please enter at most 6 characters in the \"Account Number\" field.");
    theForm.AccountVal.focus();
    return (false);
  }

  var checkOK = "0123456789-.";
  var checkStr = theForm.AccountVal.value;
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
    alert("Please enter only digit characters in the \"Account Number\" field.");
    theForm.AccountVal.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"AccountVal\" field.");
    theForm.AccountVal.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.JumpVal.value;
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
    allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Permit #\" field.");
    theForm.JumpVal.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseInt(allNum);
  if (chkVal != "" && !(prsVal >= "1"))
  {
    alert("Please enter a value greater than or equal to \"1\" in the \"Permit #\" field.");
    theForm.JumpVal.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="Permit_Detail.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1" webbot-action="--WEBBOT-SELF--">

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Permit where ID=" & Int(X) & ";")
%>
<%Permit=rstemp(2)%>
<%GISID=rstemp(52)%>
<%I_Date=rstemp(23)%>
<%US_Manhole=rstemp(54)%>
<%set sff=conntemp.execute("Select FacilityFeeID from tbl_Permit where ID=" & Int(x) & ";")%>
<%SFFID=sff(0)%>

<script language="javascript" type="text/javascript">
<!--
function comments()
{
window.open('Permit_Notes.asp?ID=<%=X%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=500,left = 250,top = 200');
}
//-->
function OLS_Edit()
{
window.open('OLS_Edit_Permit.asp?Permit=<%=Permit%>&GISID=<%=GISID%>&I_Date=<%=I_Date%>&PS=<%=rstemp(38)%>&Taps=<%=rstemp(21)%>','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=300,left = 250,top = 200');
}
//-->
</script>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="border:2px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
	<caption><b><font size="3" color="#000080">Permit Browser</font></b></caption>
		<tr>
			<td colspan="2">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td width="37"><b>Status:</b><%=rstemp(55)%></td>
						<td style="text-align: center"><a href="javascript:comments()">View Notes</a></td>
						<td><b>Account #</b></td>
						<td><%=rstemp(1)%></td>
					</tr>
				</table>
			</td>
			<td colspan="4">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3">
					<tr>
						<td><a href="javascript:OLS_Edit()">Septic Editor</a></td>
						<td>
						<p style="text-align: right"><b>Permit #</b></td>
						<td><%=rstemp(2)%></td>
						<td>
						<p style="text-align: right"><b>Lot #</b></td>
						<td><%=rstemp(3)%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="text-align: center; border-top-style:solid; border-top-width:1px; " colspan="2"><b>Permit Infomation</b></td>
			<td style="text-align: center; border-left-style:solid; border-left-width:1px; border-top-style:solid; border-top-width:1px; " colspan="4"><b>Owner Information</b></td>
		</tr>
		<tr>
			<td colspan="2" style="border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4">
					<tr>
						<td><%=rstemp(4)%>&nbsp;<%=rstemp(6)%><%If rstemp(5)<>"" then response.write"<br>":response.write rstemp(5)%></td>
					</tr>
					<tr>
						<td><%=rstemp(7)%>&nbsp;&nbsp;<%=rstemp(8)%></td>
					</tr>
					<tr><td>&nbsp;</td></tr>
				</table>
			</td>
			<td colspan="4" style="border-left-style: solid; border-left-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4">
					<tr>
						<td colspan="2"><%=rstemp(10)%>&nbsp;<%If rstemp(11) <> "" then response.write"<br>":response.write rstemp(11)%></td>
					</tr>
					<tr>
						<td colspan="2"><%=rstemp(12)%>&nbsp;<%=rstemp(13)%><%If rstemp(14)<>"" then response.write"<br>":response.write rstemp(14)%></td>
					</tr>
					<tr>
						<td colspan="2"><%=rstemp(15)%>, <%=rstemp(16)%>&nbsp;&nbsp;<%=rstemp(17)%></td>
					</tr>
					<tr>
					<td align="right" width="22%"><b>Tax Parcel ID:</b></td>
					<td width="78%"><%=rstemp(52)%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="18%"><b>Customer Type</b></td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="28%"><%=rstemp(18)%></td>
			<td colspan="3"><b>Estimated Gallons Per Day</b></td>
			<td width="19%"><%=rstemp(19)%></td>
		</tr>
		<tr>
			<td width="18%"><b>Municipality</b></td>
			<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="28%"><%=rstemp(37)%>&nbsp;</td>
			<td colspan="3"><b>Supplying Water Utility</b></td>
			<td width="19%"><%=rstemp(20)%>&nbsp;</td>
		</tr>
		<tr>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; " width="18%"><b># of 
			IDU's</b></td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-right-style:solid; border-bottom-width:1px" width="28%"><%=rstemp(21)%>&nbsp;</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; " colspan="3"><b>Map Page</b></td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; " width="19%"><%=rstemp(36)%>&nbsp;</td>
		</tr>
		<tr>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="18%">&nbsp;</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; border-right-style:solid" width="28%">&nbsp;</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="19%">
			<b>Contract#</b></td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="3%"><%=rstemp(42)%>&nbsp;</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="8%">
			<p style="text-align: right">
			<b>Map Page#</b></td>


			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="19%">
<%If Instr(Rstemp(42),"WA") then Contract=Replace(rstemp(42),"WA","") else Contract=rstemp(42)%>
<%FileCheck="D:\Shared\Asbuilts\" & Contract & "\" & rstemp(43) & ".tif"%>

<%
dim fs
set fs=Server.CreateObject("Scripting.FileSystemObject")%>
<%Link="//intranet/asbuilts/" & Contract & "/" & rstemp(43) & ".tif"%>
<%NoLink="<font color=darkred><b>" & rstemp(43) & " </b></font>"%>
<%if fs.FileExists(Filecheck)=true then Response.write "<b><a href=":Response.write link:Response.write " target=_blank>":Response.write rstemp(43):Response.write" </a></b>" else Response.write NoLink%>
<%Set FS=nothing%>			
			&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; ">
			<p style="text-align: center"><b>Lateral Information</b></td>
			<td colspan="4" style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; ">
			<p style="text-align: center"><b>Inspection Information</b></td>
		</tr>
		<tr>
			<td width="18%"><b>Length</b></td>
			<td width="28%"><%=rstemp(33)%>&nbsp;</td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<b>Inspector</b></td>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%"><%=rstemp(28)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="18%"><b>Depth</b></td>
			<td width="28%"><%=rstemp(34)%>&nbsp;</td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<b>Inspection Date</b></td>
			<%If rstemp(29)="1/1/1900" then Ins_Date="" else Ins_Date=rstemp(29)%>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%"><%=Ins_Date%>&nbsp;</td>
		</tr>
		<tr>
			<td width="18%"><b>Lateral Station</b></td>
			<td width="28%"><%=rstemp(35)%>&nbsp;</td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<b>Inspection Fee</b></td>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%">$&nbsp;<%=rstemp(24)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="18%"><b>Down/Up stream Manhole</b></td>
			<td width="28%"><%=rstemp(41)%>/<%=US_Manhole%>&nbsp;</td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<strong>Ready for Inspection</strong></td>
			<td width="28%"><%=rstemp(59)%>&nbsp;</td>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%">&nbsp;</td>
		</tr>
		<tr>

<% 
set ffconntemp=server.createobject("adodb.connection")
ffconntemp.open "DSN=PDM;"
set ffrstemp=ffconntemp.execute("Select FacilityName,Fee from tbl_Facility where ID='" & SFFID &"';")
%>
<%do until ffrstemp.eof %>
<%FacilityFee=ffrstemp(0) & ", " & FormatCurrency(ffrstemp(1))%>
<%
ffrstemp.movenext
loop

ffrstemp.close
set ffrstemp=nothing
ffconntemp.close
set ffconntemp=nothing
%>
<%If FacilityFee="" then FacilityFee="None"%>	
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; " width="18%">
			<b>Special Facility Fee</b></td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; " width="28%"><%=FacilityFee%>&nbsp;</td>
			<td colspan="4" style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; ">&nbsp;</td>
		</tr>
		<tr>
<%If rstemp(38) <> "" then DS_Pump=Int(rstemp(38)) else DS_Pump="99"%>

<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select PS_Name from tbl_PS where Station_Number='" & DS_Pump & "';")
%>
<%Session("PumpStation")=rstemp1(0)%>
<%
rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>	
	
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="18%"><b>Downstream Pump Station</b></td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="28%"><%=Session("PumpStation")%>&nbsp;</td>
			<td colspan="4" style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px; border-bottom-style:solid">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2">
			<p style="text-align: center"><b>Permit Information</b></td>
			<td colspan="4" style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px">
			<p style="text-align: center"><b>Miscellaneous Information</b></td>
		</tr>
		<tr>
			<td width="18%"><b>Application Date</b></td>
			<%If rstemp(22)="1/1/1900" then A_Date="" else A_Date=rstemp(22)%>
			<td width="28%"><%=A_Date%>&nbsp;</td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<b>Prepared by</b></td>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%"><%=rstemp(27)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="18%"><b>Issue Date</b></td>
			<%If rstemp(23)="1/1/1900" then I_Date="" else I_Date=rstemp(23)%>
			<td width="28%"><%=I_Date%>&nbsp;</td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<b>Phase</b></td>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%"><%=rstemp(31)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="18%"><b>Billing Date</b></td>
			<%If rstemp(30)="1/1/1900" then B_Date="" else B_Date=rstemp(30)%>
			<td width="28%"><%=B_Date%>&nbsp;</td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<b>Section</b></td>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%"><%=rstemp(32)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="18%"><b>Connection Fee</b></td>
			<td width="28%">$&nbsp;<%=rstemp(25)%></td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<b>Development</b></td>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%"><%=rstemp(9)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="18%"><b>Tap Fee</b></td>
			<td width="28%">$&nbsp;<%=rstemp(26)%></td>
			<td style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" colspan="3">
			<b>Check#</b></td>
			<td style="border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px" width="19%"><%=rstemp(40)%>&nbsp;</td>
		</tr>
	</table>

<%
rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5" style="font-size: 8pt">
		<tr>
			<td><input type="submit" value="Previous" name="B1"><input type="submit" value="First" name="B1"><input type="submit" value="Refresh" name="B1"></td>
			<td style="text-align: right">
			<p style="text-align: left"><input type="submit" value="Add" name="B1"><input type="submit" value="Edit" name="B1"><input type="submit" value="Permit" name="B1"></td>
			<td style="text-align: right"><b>Jump to Account#</b></td>
			<td>
			&nbsp;<!--webbot bot="Validation" s-display-name="Account Number" s-data-type="Number" s-number-separators="x." i-maximum-length="6" --><input type="text" name="AccountVal" size="6" maxlength="6"><input type="submit" value="Go" name="B2"></td>
			<td>
			<p style="text-align: right"><b>Jump to Permit#</b></td>
			<td>
			&nbsp;<!--webbot bot="Validation" s-display-name="Permit #" s-data-type="Integer" s-number-separators="x" s-validation-constraint="Greater than or equal to" s-validation-value="1" --><input type="text" name="JumpVal" size="6"><input type="submit" value="Go" name="B1"></td>
			<td style="text-align: right">
			<input type="submit" value="Last" name="B1"><input type="submit" value="Next" name="B1"></td>
		</tr>
	</table>
</div>
<input type="hidden" name="X" value="<%=X%>">
</form>

<%Sub PermitID%>
<%JumpVal=Request.form("JumpVal")%>
<%X=0%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select ID from tbl_Permit where Permit=" & JumpVal & ";")
%>
<%do until rstemp.eof%>
<%X=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%If X = 0 then ErrorOut%>
<%End sub%>

<%Sub AccountID%>
<%AccountVal=Request.form("AccountVal")%>
<%X=0%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select ID from tbl_Permit where Account='" & AccountVal & "';")
%>
<%do until rstemp.eof%>
<%X=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%If X = 0 then ErrorOut%>
<%End sub%>

<%Sub ErrorOut%>
<%X=Request.form("X")%>
<p align="center"><b>Record not found.</b></p>
<form method="POST" action="Permit_Detail.asp" name="Update" target="ibody">
        <input type="hidden" name="X" Value="<%=X%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub GetMaxID%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Top 1 ID from tbl_Permit order by ID DESC;")
%>
<%do until rstemp.eof %>
<%Response.Cookies("PDM")("MaxID")=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%End Sub%>
<%Sub DSLookup%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select PS_Name from tbl_PS where ID=" & Int(DS_Pump) & ";")
%>
<%Session("PumpStation")=rstemp1(0)%>
<%
rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>	
<%End Sub%>
<%Sub PermitEdit%>
<%X=Request.form("X")%>
<form method="POST" action="Permit_Edit.asp" name="Update" target="ibody">
        <input type="hidden" name="X" Value="<%=X%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
<%Sub PermitAdd%>
<%X=Request.form("X")%>
<form method="POST" action="Permit_Add.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
<%Sub GetPermit%>
<%X=Request.form("X")%>
<form method="POST" action="Forms/ConnectionPermit.asp" name="Update" target="_blank">
<input type="hidden" name="X" Value="<%=X%>">
<input type="hidden" name="bleh" value="bleh">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%End Sub%>
</body>

</html>