<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Create New Project</title>
<script language="javascript" type="text/javascript">
<!--
function developer()
{
window.open('d_add.asp','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=510,height=220,left = 250,top = 200');
}
function engineer()
{
window.open('e_add.asp','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=510,height=220,left = 250,top = 200');
}
//-->
</script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->
<%If Request.form("B1")="Submit" then InsertData%>
<%If Request.form("B1")="Cancel" then CancelData%>
<%If Request.form("B3")="Add" then AddDeveloper%>
<%If Request.form("B4")="Add" then AddEngineer%>
<%If Request.form("D_Name") <> "" then Session("D_Name")=Request.form("D_Name") else Session("D_Name")=""%>
<%If Request.form("E_Name") <> "" then Session("E_Name")=Request.form("E_Name") else Session("E_Name")=""%>
<div align="center">
<b><font size="3">Start New Project (Step 1)</font></b>
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
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="Project_Add.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
		<tr>
			<td style="text-align: right"><b>Project Name</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Project Name" b-value-required="TRUE" --><input type="text" name="P_Name" size="40"></td>
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
<%If rstemp(1)=Session("D_Name") then sel="Selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select><font size="1"><a href="javascript:developer()">Add New</a></font></td>
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
<%If rstemp(1)=Session("E_Name") then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select><font size="1"><a href="javascript:engineer()">Add New</a></font></td>
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
<option value="<%=rstemp(0)%>"><%=rstemp(1)%></option>
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
			<select size="1" name="Service_Type">
			<option value="0">Select Type</option>
			<option value="C">Commercial</option>
			<option value="I">Industrial</option>
			<option value="R">Residential</option>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Treatment Plant</b></td>
			<td>
			<select size="1" name="Treatment_Plant">
			<option value="0">Select Plant</option>
			<option>N/A</option>
			<option>LASA</option>
			<option>Lancaster City</option>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Water Utility</b></td>
			<td>
			<Select name="Utility" size="1">
			<option value="0">Select Water Utility</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select UID,Utility from tbl_Utility order by Utility ASC;")
%>
<%do until rstemp.eof %>
<option value="<%=rstemp(0)%>"><%=rstemp(1)%></option>
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
			<Select name="PumpStation" size="1">
			<option value="0">Select Pump Station</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Station_Number,PS_Name from tbl_PS order by Station_Number ASC;")
%>
<%do until rstemp.eof %>
<option value="<%=rstemp(0)%>"><%=rstemp(1)%></option>
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
			<td><input type="text" name="GISID" size="40"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td style="text-align: right">
			<input type="submit" value="Submit" name="B1"></td>
		</tr>
	</table>
</div>
</form>
<%Sub InsertData%>
<%P_Name=Request.form("P_Name")%>
<%P_Name=Replace(P_Name,"'","&#039")%>
<%D_Name=Request.form("Developer")%>
<%E_Name=Request.form("Engineer")%>
<%Municipality=Request.form("Municipality")%>
<%Service_Type=Request.form("Service_type")%>
<%Treatment_Plant=Request.form("Treatment_Plant")%>
<%Utility=Request.form("Utility")%>
<%PumpStation=Request.form("Pumpstation")%>
<%OpenDate=FormatDatetime(now,VbShortDate)%>
<%GISID=Request.form("GISID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Insert Into tbl_Project (Project,User_Type,Treatment_Location,Municipality,Developer,Engineer,PumpStation,Utility,OpenDate,Status,GISID) Values ('" & P_Name & "','" & Service_Type & "','" & Treatment_Plant & "','" & Municipality & "','" & D_Name & "','" & E_Name & "','" & Pumpstation & "','" & Utility & "','" & OpenDate & "','Open','" & GISID & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<div align="center">
Project has been created in the Database<br>
<a href="Project_Add_2.asp">Proceed</a> to Step 2
</div>
<%response.end%>
<%End Sub%>

<%Sub CancelData%>

<%End Sub%>




</body>

</html>