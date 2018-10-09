<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<meta http-equiv="Content-Type" content="application/ms-excel; charset=gb2312" />
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

SQL=session("SQL")
rs.open SQL,conn,1,3
Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=HoldHistory.xls"
 
response.write "<Table border=1 bordercolor='#009999' >"
response.write "<tr>"
	 
for i=0 to rs.Fields.count-1
    response.write "<td height='20' class='t-t-Borrow'><div align='center'>" & rs.Fields(i).name &"</div></td>"
next 
response.write("</tr>")	
for j=0 to rs.recordcount-1
	response.write "<tr>"	
	for i=0 to rs.Fields.count-1
		response.write "<td height='20' ><div align='center'>"& rs(i).value &"</div></td>"
  	next
	rs.movenext
	response.write "</tr>"
next 
response.write "</Table>"
response.end
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
