
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

SQL=session("SQL-SubAuto")
if(SQL="") THEN
response.end
END IF 
rs.open SQL,conn,1,3
Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=JobList.xls"
 
response.write "<Table border=1 bordercolor='#009999' >"
response.write "<tr>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>NO</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>JobNumber</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>SheetNumber</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>LineName</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>ModelName</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>StartTime</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>EndTime</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>StartQty</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>GoodQty</div></td>"
response.write "<td height='20' class='t-t-Borrow'><div align='center'>RejectQty</div></td>"
response.write("</tr>")	
for j=0 to rs.recordcount-1
	response.write "<tr>"	
		response.write "<td height='20' ><div align='center'>"& cstr(j+1) &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& rs("job_number") &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& rs("sheet_number") &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& rs("line_name") &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& rs("part_number_tag") &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& rs("start_time") &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& rs("close_time") &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& rs("JOB_START_QUANTITY") &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& rs("JOB_GOOD_QUANTITY") &"</div></td>"
		response.write "<td height='20' ><div align='center'>"& cstr(cint(rs("JOB_START_QUANTITY"))-cint(rs("JOB_GOOD_QUANTITY"))) &"</div></td>"
	rs.movenext
	response.write "</tr>"
next 
response.write "</Table>"
response.end
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->