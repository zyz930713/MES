<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

SQL=session("SQL")
rs.open SQL,conn,1,3
Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=FY_Daily_Tracking_Excel.xls"
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
   <tr>
    <td height="20" class="t-t-Borrow"><div align="center">Job_Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Job_Start_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">N_In_Store_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">S_In_Store_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">R_In_Store_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">C_In_Store_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Plan_Rework_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Output_WTD</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Projected_FG_Scrap_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Final_Scrap_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">FG_Scrap_overrun</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">FPY(%)</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Final_Yield(%)</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Final_Yield_Target(%)</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">SubSeries_Name</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Series_Name</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Family_Name</div></td>
  </tr>
  <%  if(rs.State > 0 ) then	
			for j=0 to rs.recordcount-1
	%>
	<tr>
	<td><%=rs("job_number")%></td>
	<td><%=rs("input_quantity")%></td>
	<td><%=rs("N_In_Store_Qty")%></td>
	<td><%=rs("S_In_Store_Qty")%></td>
	<td><%=rs("R_In_Store_Qty")%></td>
	<td><%=rs("C_In_Store_Qty")%></td>
	<%
		IF ISNULL(rs("N_In_Store_Qty")) THEN
			N_In_Qty=0
		else
			N_In_Qty=rs("N_In_Store_Qty")
		END IF 
		
		IF ISNULL(rs("input_quantity")) THEN
			input_quantity=0
		else
			input_quantity=rs("input_quantity")
		END IF 
		
		IF ISNULL(rs("Output_WTD")) THEN
			Output_WTD=0
		else
			Output_WTD=rs("Output_WTD")
		END IF 
		
		IF ISNULL(rs("final_scrap_quantity")) THEN
			final_scrap_quantity=0
		else
			final_scrap_quantity=rs("final_scrap_quantity")
		END IF 
		
		Plan_Rework_Qty=round(cdbl(input_quantity)*cdbl(rs("target_yield")/100)-cdbl(N_In_Qty),2)
	%>
	<td><%=Plan_Rework_Qty%></td>
	<td><%=rs("Output_WTD")%></td>
	<td><%=cdbl(input_quantity)-cdbl(Output_WTD)%></td>
	<td><%=final_scrap_quantity%></td>
	<td><%=cdbl(final_scrap_quantity)-(cdbl(input_quantity)-cdbl(Output_WTD))%></td>
	<%
		if(cdbl(input_quantity)<>0)then
			fpy=cstr(round((cdbl(N_In_Qty)/cdbl(input_quantity))*100,2))
		else
			fpy=0
		end if 
	%>
	<td><%=fpy%></td>
	<%
		if(cdbl(input_quantity)<>0)then
			finalyield=cstr(round((cdbl(Output_WTD)/cdbl(input_quantity))*100,2))
		else
			finalyield=0
		end if 
	%>
	<td><%=finalyield%></td>
	<td><%=rs("target_yield")%></td>
	<td><%=rs("subseries_name")%></td>
	<td><%=rs("SERIES_NAME")%></td>
	<td><%=rs("FAMILY_NAME")%></td>
	<%
			rs.movenext
			next 
		end if 
	%>
	
	</tr>
 
</table>

<!--#include virtual="/WOCF/BOCF_Close.asp" -->