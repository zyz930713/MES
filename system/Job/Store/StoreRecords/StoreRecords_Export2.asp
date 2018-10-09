<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
 path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=request.QueryString("jobnumber")
 
SQL=session("SQL")
Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=bom.xls"
rs.open SQL,conn,1,3
 
%>
<table width="98%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
	<td class="t-t-Borrow">入库时间</td>
	<td class="t-t-Borrow">生产数量</td>
	<td class="t-t-Borrow">入库数量</td>
	<td class="t-t-Borrow">型号</td>
	<td class="t-t-Borrow">BOM</td>
  </tr>
  
	<%  if(rs.State > 0 ) then	
			while not rs.eof
	%>
	<tr>
	 	<td><%=rs(0).value%></td>
		<td><%=rs(1).value%></td>
		<td><%=rs(2).value%> </td>
		<td><%=rs(3).value%> </td>
		<td><%=rs(4).value%> </td>

	</tr>
	<%
				rs.movenext
		 	wend
		end if
	%>
</table>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->