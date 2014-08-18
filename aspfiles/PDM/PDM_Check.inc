<%If Request.Cookies("Portal")("FullName") = "" then LoginError%>
<%If Int(Instr(Request.Cookies("Portal")("PDM"),"R")) = 0 then AccessError%>
<%Sub LoginError%>
<br>
<br>
<p align="center"><b>You must login before you can use this Application</b></p>
<form method="POST" action="../home.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End sub%>
<%Sub AccessError%>
<br><br>
<p align="center"><b>You do not have access to this Application</b></p>
<form method="POST" action="../home.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>