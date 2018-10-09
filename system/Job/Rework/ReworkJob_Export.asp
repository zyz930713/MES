<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
dim SQLQuery
 SQLQuery=session("SQLQuery")
rs.Open SQLQuery,conn,1,3

call ExportToExcel
rs.Close
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
 <%
  sub ExportToExcel     
      Response.ContentType   =   "application/vnd.ms-excel"   
      Response.AddHeader   "Content-Disposition",   "attachment;Filename=ReworkJobList.xls"   
      Response.Write   "<body>"   
      Response.Write   "<table border=1>"  
	  if not rs.eof then 
	      Call   WriteTableData   
	  end if
      Response.Write   "</table>"   
      Response.Write   "</body>"   
      Response.Write   "</html>"     
  End Sub   
    
  Sub WriteTableData    
	  '   write   header   
	  Response.Write   "<tr>"   
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>NO</span> </td>"  
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>Rework Job Number</span> </td>" 
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>Rework Job Type</span> </td>" 
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>Rework Quantity</span> </td>" 
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>Good Quantity</span> </td>"  
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>Reject Quantity</span> </td>"  
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>Start Time</span> </td>"  
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>End Time</span> </td>"  
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>Operator Code</span> </td>"   
	   Response.Write   "<td style='background-color:#73A2EE;'><span style='color: #ffffff'>Yield</span> </td>" 
	  Response.Write   "</tr>"   
  '   write   data   rows   
 	  dim i  
	i=1
if not rs.eof then
while not rs.eof 
%>
<Tr>
  <td height="20"><div align="center"><%=i%></div></td>
  <td height="20"><div align="center"><%=rs("REWORK_JOB_NUMBER").value%></div></td>
  <td><div align="center">
  	<% if rs("REWORK_TYPE").value="0" then response.write "Readjust"  end if 
	   if rs("REWORK_TYPE").value="1" then response.write "Teardown" end if  
	   if rs("REWORK_TYPE").value="2" then response.write "Slapping"  end if %></div>
  </td>
  <td ><div align="center"><%=rs("SUB_JOB_REWORK_QTY").value%></div></td>
  <td ><div align="center"><%=rs("REWORK_GOOD_QTY").value%></td>
  <td ><div align="center"><%=rs("REWORK_REJECT_QTY").value%></div></td>
  <td ><div align="center"><%=rs("REWORK_START_TIME").value%></div></td>
  <td ><div align="center"><%=rs("REWORK_END_TIME").value%></div></td>
  <td ><div align="center"><%=rs("OPERATOR_CODE").value%></div></td>
  <td ><div align="center"><%=round(cdbl(rs("REWORK_GOOD_QTY").value)/cdbl(rs("SUB_JOB_REWORK_QTY").value),4)*100%>%</div></td>
 </tr>
<%	
	i=i+1
	rs.movenext
wend
end if
  End  Sub   
%>