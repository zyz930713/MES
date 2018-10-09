<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

SQL=session("SQL")
Detail=session("Detail")
rs.open SQL,conn,1,3
Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=Yield.xls"
 
response.write "<Table border=1 bordercolor='#009999' >"
response.write "<tr>"
	 
for i=0 to rs.Fields.count-1
    response.write "<td height='20' class='t-t-Borrow'><div align='center'>" & rs.Fields(i).name &"</div></td>"
next 
if detail="1" then
    response.write "<td height='20' class='t-t-Borrow'><div align='center'>BatchType</div></td>"
end if 
response.write("</tr>")	
for j=0 to rs.recordcount-1
	response.write "<tr>"	
	for i=0 to rs.Fields.count-1
		response.write "<td height='20' ><div align='center'>"& rs(i).value &"</div></td>"
  	next
	if Detail="1" then
		batchno=rs("BATCHNO")
		batch_type=""
		if(right(batchno,2)="-5") then
			batch_type="S"
		else
			sql="select * from label_print_history where batchno='"+batchno+"'"
			set rs2=server.createobject("adodb.recordset")
			rs2.open sql,conn,1,3
			if rs2.recordcount>0 then
				subjoblist=rs2("subjoblist")
				arrsubjoblist=split(subjoblist,"-")
				if(cint(arrsubjoblist(0))<100) then
					batch_type="N"
				end if 
				if(cint(arrsubjoblist(0))<600 and cint(arrsubjoblist(0))>=500) then
					batch_type="R"
				end if 
				if(cint(arrsubjoblist(0))>=700) then
					batch_type="C"
				end if 
			end if 
		end if 
		response.write "<td height='20' ><div align='center'>"& batch_type &"</div></td>"
	end if 
	response.write "</tr>"
	rs.movenext
next 
response.write "</Table>"
response.end
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->