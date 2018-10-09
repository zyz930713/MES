<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

SQL=session("SQL")

rs.open SQL,conn,1,3

Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=ExportExcel.xls"

response.write "<Table border=1 bordercolor='#009999' >"
response.write "<tr>"
response.write "<td>Job Number 工单号</td>"
response.write "<td>Model Name 型号</td>"
response.write "<td>Sub Job List 子工单号</td>"
response.write "<td>Qty 数量</td>"
response.write "<td>Actual Qty 实际数量</td>"
response.write "<td>Operator 工号</td>"
response.write "<td>DateTime 发生时间</td>"
response.write "<td>Remark 备注</td>"
response.write "<td>Job Release 是否已开工单 </td>"
response.write "<td>New Job Number 打印新工单号</td>"
response.write "<td>Transaction Type 类型</td>"
response.write "</tr>"

for i=0 to rs.recordcount-1
	response.write "<tr>"
	response.write "<td>" 
	response.write rs("JOB_NUMBER")
	response.write "</td>" 
	response.write "<td>" 
	response.write rs("part_number_tag")
	response.write "</td>" 
	response.write "<td>" 
	response.write rs("SUBJOBLIST")
	response.write "</td>" 
	response.write "<td>" 
	response.write rs("QTY")
	response.write "</td>" 
	response.write "<td>" 
	response.write rs("ACTUALQTY")
	response.write "</td>" 
	response.write "<td>" 
	response.write rs("OPERATOR")
	response.write "</td>" 
	response.write "<td>" 
	response.write rs("INTIMESTAMP")
	response.write "</td>" 
	response.write "<td>" 
	response.write rs("REMARK")
	response.write "</td>" 
	response.write "<td>" 
	if rs("ISPRINT")="1" then response.write "Yes" end if 
	if rs("ISPRINT")="0" then response.write "NO" end if 
	response.write "</td>" 
	response.write "<td>" 
	response.write rs("NEW_JOBNUMBER")
	response.write "</td>" 
	response.write "<td>" 
		if rs("TransactionType")="-1" then response.write "Normal" end if  
		if rs("TransactionType")="0"  then response.write "None" end if 
		if rs("TransactionType")="1" then response.write "Rework" end if
		if rs("TransactionType")="2" then response.write "Scrap" end if
		if rs("TransactionType")="3" then response.write "Readjust" end if 
		if rs("TransactionType")="4" then response.write "Retest" end if
		if rs("TransactionType")="5" then response.write "Slapping" end if 
	response.write "</td>" 
	response.write "</tr>" 
	rs.movenext
next  
response.write "</Table>"
response.end

%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
