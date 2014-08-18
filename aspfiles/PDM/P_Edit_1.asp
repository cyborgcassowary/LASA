<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Update Project Header(1)</title>
<script language="JavaScript">
<!--
function closewinder()
{
 self.close();
}
//-->
</script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->
<%PID=Request.QueryString("PID")%>
<%If Request.form("B1")="Update" then UpdateData%>
<%If Request.form("B1")="Cancel" then CancelData%>
<%If Request.form("D_Name") <> "" then Session("D_Name")=Request.form("D_Name") else Session("D_Name")=""%>
<%If Request.form("E_Name") <> "" then Session("E_Name")=Request.form("E_Name") else Session("E_Name")=""%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_project where PID=" & PID & ";")
%>
<%do until rstemp.eof %>
<%Project=rstemp(1)%>
<%Municipality=rstemp(6)%>
<%Plant=rstemp(5)%>
<%User_Type=rstemp(4)%>
<%DID=rstemp(8)%>
<%EID=rstemp(9)%>
<%Utility=rstemp(15)%>
<%PumpStation=rstemp(14)%>
<%GISID=rstemp(23)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<div align="center">
<b><font size="3">Update Project Information</font></b>
</div>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.P_Name.value == "")
  {
    alert("Please enter a value for the \"Project Name\" field.");
    theForm.P_Name.focus();
    return (false);
  }

  if (theForm.Developer.selectedIndex == 0)
  {
    alert("The first \"Developer\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Developer.focus();
    return (false);
  }

  if (theForm.Engineer.selectedIndex == 0)
  {
    alert("The first \"Engineer\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Engineer.focus();
    return (false);
  }

  if (theForm.Municipality.selectedIndex == 0)
  {
    alert("The first \"Municipality\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Municipality.focus();
    return (false);
  }

  if (theForm.Service_Type.selectedIndex == 0)
  {
    alert("The first \"Service Type\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Service_Type.focus();
    return (false);
  }

  if (theForm.Treatment_Plant.selectedIndex == 0)
  {
    alert("The first \"Treatment Plant\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Treatment_Plant.focus();
    return (false);
  }

  if (theForm.Utility.selectedIndex == 0)
  {
    alert("The first \"Water Utility\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Utility.focus();
    return (false);
  }

  if (theForm.PumpStation.selectedIndex == 0)
  {
    alert("The first \"Pump Station\" option is not a valid selection.  Please choose one of the other options.");
    theForm.PumpStation.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="P_Edit_1.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
		<tr>
			<td style="text-align: right"><b>Project Name</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Project Name" b-value-required="TRUE" --><input type="text" name="P_Name" size="40" value="<%=Project%>"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Developer</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Developer" b-disallow-first-item="TRUE" --><Select name="Developer" size="1">
<option>Select Developer</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select DID,D_Name from tbl_Developer order by D_Name ASC;")
%>
<%do until rstemp.eof %>
<%IF Int(DID)=Int(rstemp(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%></option>
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
		<tr>
			<td style="text-align: right"><b>Engineer</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Engineer" b-disallow-first-item="TRUE" --><Select name="Engineer" size="1">
<option>Select Engineer</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select EID,E_Name from tbl_Engineer order by E_Name ASC;")
%>
<%do until rstemp.eof %>
<%if Int(EID)=int(rstemp(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></font></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Municipality</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Municipality" b-disallow-first-item="TRUE" --><Select name="Municipality" size="1">
<option value="0">Select Municipality</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select MID,Municipality from tbl_Municipality order by Municipality ASC;")
%>
<%do until rstemp.eof %>
<%If int(Municipality)=int(rstemp(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select>
			</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>&nbsp;Service Type</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Service Type" b-disallow-first-item="TRUE" --><select size="1" name="Service_Type">
			<option value="0">Select Type</option>
			<%If User_Type="C" then sel="selected" else sel=""%>
			<option value="C" <%=sel%>>Commercial</option>
			<%If User_Type="I" then sel="selected" else sel=""%>
			<option value="I" <%=sel%>>Industrial</option>
			<%If User_Type="R" then sel="selected" else sel=""%>
			<option value="R" <%=sel%>>Residential</option>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Treatment Plant</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Treatment Plant" b-disallow-first-item="TRUE" --><select size="1" name="Treatment_Plant">
			<option value="0">Select Plant</option>
			<%If Plant="LASA" then sel="selected" else sel=""%>
			<option <%=sel%>>LASA</option>
			<%If Plant="Lancaster City" then sel="selected" else sel=""%>
			<option <%=sel%>>Lancaster City</option>
			<%If Plant="N/A" then sel="selected" else sel=""%>
			<option <%=sel%>>N/A</option>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Water Utility</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Water Utility" b-disallow-first-item="TRUE" --><Select name="Utility" size="1">
			<option value="0">Select Water Utility</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select UID,Utility from tbl_Utility order by Utility ASC;")
%>
<%do until rstemp.eof %>
<%If Int(utility)=Int(rstemp(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%></option>
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
		<tr>
			<td style="text-align: right"><b>Downstream Pump Station</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Pump Station" b-disallow-first-item="TRUE" --><Select name="PumpStation" size="1">
			<option value="0">Select Pump Station</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Station_Number,PS_Name from tbl_PS order by Station_Number ASC;")
%>
<%do until rstemp.eof %>
<%If int(pumpstation)=int(rstemp(0)) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%></option>
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
		<tr>
			<td style="text-align: right"><b>GIS Parcel ID</b></td>
			<td><input type="text" name="GISID" size="40" value="<%=GISID%>"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td style="text-align: right">
			<input type="submit" value="Update" name="B1"></td>
		</tr>
	</table>
</div>
<input type="hidden" name="PID" value="<%=PID%>">
</form>
<%Sub UpdateData%>
<%PID=Request.form("PID")%>
<%P_Name=Request.form("P_Name")%>
<%P_Name=Replace(P_Name,"'","&#039")%>
<%D_Name=Request.form("Developer")%>
<%E_Name=Request.form("Engineer")%>
<%Municipality=Request.form("Municipality")%>
<%Service_Type=Request.form("Service_type")%>
<%Treatment_Plant=Request.form("Treatment_Plant")%>
<%Utility=Request.form("Utility")%>
<%PumpStation=Request.form("Pumpstation")%>
<%OpenDate=FormatDatetime(now)%>
<%GISID=Request.form("GISID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_project Set Project='" & P_Name & "', User_type='" & Service_Type & "', Treatment_location='" & Treatment_Plant & "', Municipality='" & Municipality & "', Developer='" & D_Name & "', Engineer='" & E_Name & "', PumpStation='" & PumpStation & "', Utility='" & Utility & "', LastUpdate='" & OpenDate & "', GISID='" & GISID & "' Where PID=" & PID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">Information has been updated<br><br>
You must continue to the next step to finish the update<br>
<a href="P_Edit_2.asp?PID=<%=PID%>">Proceed</a></p>
<form method="POST" action="Project_Add_3.asp" name="Update" target="ibody">
        <input type="hidden" name="Refresh" Value="Updated">
        <input type="hidden" name="PID" Value="<%=PID%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%response.end%>
<%End Sub%>

<%Sub CancelData%>

<%End Sub%>

</body>

</html>
