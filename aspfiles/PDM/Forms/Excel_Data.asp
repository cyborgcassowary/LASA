<%X=Request.form("X")%>
<%If Request.form("Bleh") <> "bleh" then Response.Cookies("PDM")("MaxID")=Int(Request.Cookies("PDM")("MaxID"))+1%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Permit where ID=" & X & ";")
%>
<%If rstemp(53) <> "" then SFFID=rstemp(53) else SFFID=0%>
<%PID=rstemp(50)%>
<%C1=rstemp(1)%>
<%PermitK1=rstemp(2)%>
<%IF rstemp(3) <> "" then LOT=" LOT# " & rstemp(3)%>
<%D7=rstemp(4) & " " & rstemp(6) & " " & rstemp(7) & " " & rstemp(8) & LOT%>
<%If rstemp(31) <> "" then Phase=" Phase " & rstemp(31)%>
<%If rstemp(32) <> "" then Section=" Section " & rstemp(32)%>
<%D8=rstemp(9) & Phase &  Section%>
<%D9=rstemp(10)%>
<%owner2=rstemp(11)%>
<%If owner2 <> "" then D9=D9 & " " & owner2%>
<%D10=rstemp(12) & " " & rstemp(13)%>
<%D11=rstemp(15) & " " & rstemp(16) & " " & rstemp(17)%>
<%If Int(Instr(rstemp(18),"R")) > 0 then E13="X" else E13=""%>
<%If Int(Instr(rstemp(18),"C")) > 0 then H13="X" else H13=""%>
<%If Int(Instr(rstemp(18),"I")) > 0 then K13="X" else K1=""%>
<%I15=rstemp(19)%>
<%D18=rstemp(20)%>
<%J18=rstemp(21)%>
<%H21=rstemp(37)%>
<%I32=rstemp(22)%>
<%I33=rstemp(23)%>
<%I36=rstemp(26)%>
<%If IsNull(I36) then I36=0%>
<%I37=rstemp(25)%>
<%if IsNull(I37) then I37=0%>
<%If Instr(rstemp(24),",") then
	 	I38Array=split(rstemp(24),",")
		I38=I38Array(0)
		I38b=I38Array(1)
	 else
	 	I38=rstemp(24)
End If%>
<%If I38b="None" then I38b=""%>
<%If IsNull(I38) then I38=0%>
<%I40=rstemp(40)%>
<%I41=rstemp(27)%>
<%I43=rstemp(28) & " " & rstemp(29)%>
<%If Instr(I43,"1/1/1900") then I43=replace(I43,"1/1/1900","")%>
<%C47=rstemp(33)%>
<%C48=rstemp(34)%>
<%C49=rstemp(35)%>
<%C46=rstemp(36)%>
<%C51=rstemp(38)%>
<%C50=rstemp(41)%>
<%C50US=rstemp(54)%>
<%B36=rstemp(44)%>
<%CRLF=Chr(10) & Chr(13)%>
<%B38=rstemp(45)%>
<%B39=rstemp(47) & " " & rstemp(48) & " " & rstemp(49)%>
<%B42=rstemp(46)%>

<%FF_Status=rstemp(56)%>
<%If FF_Status="None" then FF_Status=""%>
<%CF_Status=rstemp(57)%>
<%If CF_Status="None" then CF_Status=""%>
<%TF_Status=rstemp(58)%>
<%IF TF_Status="None" then TF_Status=""%>
<%
rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Fee from tbl_Facility where ID=" & SFFID & ";")
%>
<%do until rstemp.eof %>
<%I35=rstemp(0)%>
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
set rstemp=conntemp.execute("Select EscrowFile from tbl_Project where PID=" & PID & ";")
%>
<%do until rstemp.eof %>
<%K2=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%If IsNull(I35) then I35=0%>

