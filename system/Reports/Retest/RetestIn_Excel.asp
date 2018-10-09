<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

SQL=session("SQL")
rs.open SQL,conn,1,3
Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=RetestIn.xls"
function getMainJobNumber(JobNumber)
		arrJob=split(JobNumber,"-")
		NewJobNumber=""
		 if arrJob(1)="E" or arrJob(1)="R" then
			NewJobNumber=arrJob(0)&"-"&arrJob(1)
		else
			NewJobNumber=arrJob(0)
		end if
		getMainJobNumber=NewJobNumber
	end function
%>
 
<table width="98%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
	<td class="t-t-Borrow">BatchNO</td>
	<td class="t-t-Borrow">子工单号</td>
	<td class="t-t-Borrow">型号</td>
	<td class="t-t-Borrow">接受时间</td>
	<td class="t-t-Borrow">接受人</td>
	<td class="t-t-Borrow">接受数量</td>
	<td class="t-t-Borrow">工单类型</td>
	<td class="t-t-Borrow">线别</td>
  </tr>
  
	<%  if(rs.State > 0 ) then	
			for j=0 to rs.recordcount-1
	%>
	<tr>
	 	<td><%=rs(0).value%></td>
		<td><%=rs(4).value%></td>
 		<%
			SQL="SELECT * FROM label_print_history WHERE BATCHNO='"+rs(0).value+"' AND (SUBJOBLIST LIKE '%70%' OR SUBJOBLIST LIKE '%71%' OR SUBJOBLIST LIKE '%72%')"
			set rsModelType=server.createobject("adodb.recordset")
			rsModelType.open SQL,conn,1,3
			if(rsModelType.recordcount=0) then
				sql="select part_number_tag from job where job_number='"+getMainJobNumber(rs(0).value)+"' and job_type='N' "
				set rsModel1=server.createobject("adodb.recordset")
				rsModel1.open SQL,conn,1,3	
				if(rsModel1.recordcount>0) then
					modelname=rsModel1(0)
				end if 
			else
				sql="select part_number_tag from job where job_number='"+getMainJobNumber(rs(0).value)+"' and job_type='C' "
				set rsModel2=server.createobject("adodb.recordset")
				rsModel2.open SQL,conn,1,3
				if(rsModel2.recordcount>0) then
					modelname=rsModel2(0)
				end if 	
			end if
		%>
		<td><%=modelname%> </td>
		<td><%=rs(1).value%> </td>
		<td><%=rs(2).value%></td>
		<td><%=rs(3).value%></td>
		<td>
		  <%if rs("TransactionType")="-1" then response.write "Normal" end if %> 
		  <%if rs("TransactionType")="0"  then response.write "None" end if  %>  
		  <%if rs("TransactionType")="1" then response.write "Rework" end if %>  
		  <%if rs("TransactionType")="2" then response.write "Scrap" end if %>  
		  <%if rs("TransactionType")="3" then response.write "Readjust" end if %> 
		  <%if rs("TransactionType")="4" then response.write "Change Model" end if %> 
		  <%if rs("TransactionType")="5" then response.write "Slapping" end if %> 
		</td>
		<td><%=rs(6).value%></td>
		<%TotalQty=TotalQty+cint(rs(3).value)%>
		
	<%
				rs.movenext
		 	next	
	%>
	</tr>
	<%
		end if
	%>
	
  </tr>
  <tr>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td  class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">总数量</td>
	<td  class="t-t-Borrow"><%=TotalQty%>&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
  </tr>
   
 
</table>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->