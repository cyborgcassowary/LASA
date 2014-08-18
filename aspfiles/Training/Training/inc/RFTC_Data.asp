<%Req_Date=Request.form("Req_Date")%>
<%If Req_Date="" then Req_Date=Request.QueryString("Req_Date")%>
<%If Req_Date="" then Req_Date="5/1/2006 4:26:31 PM"%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select * from tbl_users where Req_Date=#" & Req_Date & "#;")
%>
<%do until rstemp.eof %>
<%C5=rstemp(1)%>
<%H5=rstemp(23)%>
<%C7=rstemp(2)%>
<%H7=rstemp(16)%>
<%C9=rstemp(3)%>
<%C11=rstemp(4)%>
<%C13=rstemp(5)%>
<%H13=rstemp(19)%>
<%C15=rstemp(6)%>
<%Days=rstemp(22)%>
<%I15=rstemp(17)%>
<%C17=rstemp(25)%>
<%I17=rstemp(26)%>
<%C19=rstemp(8)%>
<%C21=rstemp(7)%>
<%E26=rstemp(9)%>
<%E28=rstemp(10)%>
<%E30=rstemp(11)%>
<%E32=rstemp(12)%>
<%E34=rstemp(13)%>
<%G34=rstemp(14)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%E36=Int(E26)+int(E28)+int(E30)+int(E32)+int(E34)%>
<%EndDate=DateAdd("d",Days,C15)%>
<%If Days <> 1 then C15=C15 & " - " & Enddate%>