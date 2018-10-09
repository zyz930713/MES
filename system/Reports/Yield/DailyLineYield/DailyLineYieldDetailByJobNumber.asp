<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
family_name=request.QueryString("family_name")
factory_id=request.QueryString("factory_id")
line_name=request.QueryString("line_name")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JOB_NUMBER,SHEET_NUMBER"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/Yield/DailyLineYield/DailyLineDetailbyJobNumber.asp.asp"
SQL="select INCLUDED_SYSTEM_ITEMS from SERIES_GROUP where SERIES_GROUP_NAME='"&family_name&"' and FACTORY_ID='"&factory_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
included_system_items=rs("INCLUDED_SYSTEM_ITEMS")
end if
rs.close
SQL="select * from JOB where FACTORY_ID='"&factory_id&"' and instr('"&included_system_items&"',PART_NUMBER_TAG)>0 and LINE_NAME='"&line_name&"' and CLOSE_TIME>=to_date('"&from_time&"','yyyy-mm-dd hh24:mi:ss') and CLOSE_TIME<=to_date('"&to_time&"','yyyy-mm-dd hh24:mi:ss') "&order
session("aerror")=SQL
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language=JavaScript src="/Functions/TablePlus.js" type=text/javascript></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">Browse Daily Line Yield -- <%=line_name%> (<%=from_time%> - <%=to_time%>) </td>
  </tr>
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">User:
      <% =session("User") %></td>
  </tr>
<tr>
  <td height="20" colspan="7">Sort by: 
    <input name="ByNormal" type="radio" value="1" onClick="javascript:window.open('DailyLineYieldDetail.asp?family_name=<%=family_name%>&factory_id=<%=factory_id%>&line_name=<%=line_name%>&from_time=<%=from_time%>&to_time=<%=to_time%>','_self')">
    Normal
    <input name="ByPart" type="radio" value="1" onClick="javascript:window.open('DailyLineYieldDetailByPartNumber.asp?family_name=<%=family_name%>&factory_id=<%=factory_id%>&line_name=<%=line_name%>&from_time=<%=from_time%>&to_time=<%=to_time%>','_self')">
    Part Number </td>
</tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Job Number</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Part Number</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">InputQuantity </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Output Quantity</div></td>
    <td height="20" class="t-t-Borrow"><div align="center"> Yield </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Close Time </div></td>
  </tr>
  <%
i=1
if not rs.eof then
dim this_jobnumber()
dim this_sheetnumber()
dim this_jobtype()
dim this_partnumbertag()
dim this_input_quantity()
dim this_output_quantity()
dim this_yield()
dim this_close_time()
redim this_jobnumber(0)
redim this_sheetnumber(0)
redim this_jobtype(0)
redim this_partnumbertag(0)
redim this_input_quantity(0)
redim this_output_quantity(0)
redim this_yield(0)
redim this_close_time(0)
total_input_quantity=0
total_output_quantity=0
total_yield=0
all_input=0
all_output=0
current_jobnumber=rs("JOB_NUMBER")
current_partnumber=rs("PART_NUMBER_TAG")
t=1
while not rs.eof
	if rs("JOB_NUMBER")=current_jobnumber then 'Get total parameters' value of all the same jobs
		this_jobnumber(UBound(this_jobnumber))=rs("JOB_NUMBER")
		this_sheetnumber(UBound(this_sheetnumber))=rs("SHEET_NUMBER")
		this_jobtype(UBound(this_jobtype))=rs("JOB_TYPE")
		this_partnumbertag(UBound(this_partnumbertag))=rs("PART_NUMBER_TAG")
		this_input_quantity(UBound(this_input_quantity))=rs("JOB_START_QUANTITY")
		this_output_quantity(UBound(this_output_quantity))=rs("JOB_GOOD_QUANTITY")
		if rs("JOB_START_QUANTITY")<>"0" then
		this_yield(UBound(this_yield))=cint(rs("JOB_GOOD_QUANTITY"))/cint(rs("JOB_START_QUANTITY"))
		else
		this_yield(UBound(this_yield))=0
		end if
		this_close_time(UBound(this_close_time))=rs("CLOSE_TIME")
		ReDim Preserve this_jobnumber(UBound(this_jobnumber)+1)
		ReDim Preserve this_sheetnumber(UBound(this_sheetnumber)+1)
		ReDim Preserve this_jobtype(UBound(this_jobtype)+1)
		ReDim Preserve this_partnumbertag(UBound(this_partnumbertag)+1)
		ReDim Preserve this_input_quantity(UBound(this_input_quantity)+1)
		ReDim Preserve this_output_quantity(UBound(this_output_quantity)+1)
		ReDim Preserve this_yield(UBound(this_yield)+1)
		ReDim Preserve this_close_time(UBound(this_close_time)+1)
		total_input_quantity=total_input_quantity+cint(rs("JOB_START_QUANTITY"))
		total_output_quantity=total_output_quantity+cint(rs("JOB_GOOD_QUANTITY"))
		if total_input_quantity<>0 then
		total_yield=total_output_quantity/total_input_quantity
		else
		total_yield=0
		end if
	else
	%>
  <tr class="t-b-blue">
    <td height="20"><div align="center">
      <% =t%>
    </div></td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=current_jobnumber%>')">
      <input name="tabstatus<%=current_jobnumber%>" type="hidden" value="0">
      <img src="/Images/Treeimg/plus.gif" name="tabimg<%=current_jobnumber%>" width="9" height="9">&nbsp;<%= current_jobnumber %></span></div></td>
    <td height="20"><div align="center"><%=current_partnumber%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_input_quantity%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_output_quantity%></div></td>
    <td height="20"><div align="center"><%=formatpercent(total_yield,2,-1)%>&nbsp;</div></td>
    <td height="20">&nbsp;</td>
  </tr>
  <tbody id="tab<%=current_jobnumber%>" style="display:none">
    <%
			for u=0 to ubound(this_jobnumber)-1 'Display detail of each sheet.
%>
    <tr>
      <td height="20"><div align="center"><%=u+1%></div></td>
      <td height="20">
        <div align="center"><span style="cursor:hand" class="d_link" onClick="javascript:window.open('/Job/SubJobs/JobDetail.asp?jobnumber=<%=this_jobnumber(u)%>&sheetnumber=<%=this_sheetnumber(u)%>&jobtype=<%=this_jobtype(u)%>','_blank')"><%=this_jobnumber(u)%>-<%=replace(this_jobtype(u),"N","")&repeatstring(this_sheetnumber(u),"0",3)%></span>&nbsp;
            <%if DBA=true then%>
            <span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('/Job/SubJobs/JobStationUpdate.asp?jobnumber=<%=this_jobnumber(u)%>&sheetnumber=<%=this_sheetnumber(u)%>&jobtype=<%=this_jobtype(u)%>&path=<%=path%>&query=<%=query%>','_self')">U</span>
            <%end if%>
        </div>
      </td>
      <td height="20"><div align="center"><%= this_partnumbertag(u) %>&nbsp;</div></td>
      <td height="20"><div align="center"><% =this_input_quantity(u)%></div></td>
      <td height="20"><div align="center"><%=this_output_quantity(u)%></div></td>
      <td height="20"><div align="center"><%=formatpercent(this_yield(u),2,-1)%>&nbsp;</div></td>
      <td height="20"><div align="center"><%=this_close_time(u)%></div></td>
    </tr>
    <%		next%>
  </tbody>
  <%
		redim this_jobnumber(0)
		redim this_sheetnumber(0)
		redim this_jobtype(0)
		redim this_partnumbertag(0)
		redim this_input_quantity(0)
		redim this_output_quantity(0)
		redim this_yield(0)
		redim this_close_time(0)
		total_input_quantity=0
		total_output_quantity=0
		this_jobnumber(UBound(this_jobnumber))=rs("JOB_NUMBER")
		this_sheetnumber(UBound(this_sheetnumber))=rs("SHEET_NUMBER")
		this_jobtype(UBound(this_jobtype))=rs("JOB_TYPE")
		this_partnumbertag(UBound(this_partnumbertag))=rs("PART_NUMBER_TAG")
		this_input_quantity(UBound(this_input_quantity))=rs("JOB_START_QUANTITY")
		this_output_quantity(UBound(this_output_quantity))=rs("JOB_GOOD_QUANTITY")
		if rs("JOB_START_QUANTITY")<>"0" then
		this_yield(UBound(this_yield))=cint(rs("JOB_GOOD_QUANTITY"))/cint(rs("JOB_START_QUANTITY"))
		else
		this_yield(UBound(this_yield))=0
		end if
		this_close_time(UBound(this_close_time))=rs("CLOSE_TIME")
		ReDim Preserve this_jobnumber(UBound(this_jobnumber)+1)
		ReDim Preserve this_sheetnumber(UBound(this_sheetnumber)+1)
		ReDim Preserve this_jobtype(UBound(this_jobtype)+1)
		ReDim Preserve this_partnumbertag(UBound(this_partnumbertag)+1)
		ReDim Preserve this_input_quantity(UBound(this_input_quantity)+1)
		ReDim Preserve this_output_quantity(UBound(this_output_quantity)+1)
		ReDim Preserve this_yield(UBound(this_yield)+1)
		ReDim Preserve this_close_time(UBound(this_close_time)+1)
		total_input_quantity=total_input_quantity+cint(rs("JOB_START_QUANTITY"))
		total_output_quantity=total_output_quantity+cint(rs("JOB_GOOD_QUANTITY"))
		if total_input_quantity<>0 then
		total_yield=total_output_quantity/total_input_quantity
		else
		total_yield=0
		end if
		t=t+1
	end if
	all_input=all_input+cint(rs("JOB_START_QUANTITY"))
	all_output=all_output+cint(rs("JOB_GOOD_QUANTITY"))
i=i+1
current_jobnumber=rs("JOB_NUMBER")
current_partnumber=rs("PART_NUMBER_TAG")
rs.movenext
wend
%>
  <tr class="t-b-blue">
    <td height="20"><div align="center">
      <% =t%>
    </div></td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=current_jobnumber%>')">
      <input name="tabstatus<%=current_jobnumber%>" type="hidden" value="0">
      <img src="/Images/Treeimg/plus.gif" name="tabimg<%=current_jobnumber%>" width="9" height="9">&nbsp;<%= current_jobnumber %></span></div></td>
    <td height="20"><div align="center"><%=current_partnumber%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_input_quantity%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_output_quantity%></div></td>
    <td height="20"><div align="center">
    <%=formatpercent(total_yield,2,-1)%>&nbsp;</div></td>
    <td height="20">&nbsp;</td>
  </tr>
  <tbody id="tab<%=current_jobnumber%>" style="display:none">
    <%
		for u=0 to ubound(this_jobnumber)-1 'Display detail of each sheet.
%>
    <tr>
      <td height="20"><div align="center"><%=u+1%></div></td>
      <td height="20">
        <div align="center"><span style="cursor:hand" class="d_link" onClick="javascript:window.open('/Job/SubJobs/JobDetail.asp?jobnumber=<%=this_jobnumber(u)%>&sheetnumber=<%=this_sheetnumber(u)%>&jobtype=<%=this_jobtype(u)%>','_blank')"><%=this_jobnumber(u)%>-<%=replace(this_jobtype(u),"N","")&repeatstring(this_sheetnumber(u),"0",3)%></span>&nbsp;
          <%if DBA=true then%>
          <span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('/Job/SubJobs/JobStationUpdate.asp?jobnumber=<%=this_jobnumber(u)%>&sheetnumber=<%=this_sheetnumber(u)%>&jobtype=<%=this_jobtype(u)%>&path=<%=path%>&query=<%=query%>','_self')">U</span>
          <%end if%>
</div></td>
      <td height="20"><div align="center"><%= this_partnumbertag(u) %>&nbsp;</div></td>
      <td height="20"><div align="center"><%=this_input_quantity(u)%></div></td>
      <td height="20"><div align="center"><%=this_output_quantity(u)%></div></td>
      <td height="20"><div align="center"><%=formatpercent(this_yield(u),2,-1)%>&nbsp;</div></td>
      <td height="20"><div align="center"><%=this_close_time(u)%></div></td>
    </tr>
	<% next%>
  </tbody>
	<tr>
      <td height="20">&nbsp;</td>
      <td height="20"><div align="center">Total</div></td>
      <td height="20">&nbsp;</td>
      <td height="20"><div align="center"><%=all_input%>&nbsp;</div></td>
      <td height="20"><div align="center"><%=all_output%></div></td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
    </tr>
  <%
else
%>
  <tr>
    <td height="20" colspan="7"><div align="center">No Records </div></td>
  </tr>
  <%end if
rs.close
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->