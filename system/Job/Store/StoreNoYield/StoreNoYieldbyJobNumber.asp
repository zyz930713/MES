<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/Store/StoreNoYieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Reports/Store/StoreNoYield/StoreNoYieldbyJobNumber.asp"
SQL=session("SQL")
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language=JavaScript src="/Functions/TablePlus.js" type=text/javascript></script>
<!--#include virtual="/Job/Store/StoreNoYield/NoYieldCheck.asp" -->
<!--#include virtual="/Language/Reports/Store/StoreNoYield/Lan_StoreNoYieldbyJobNumber.asp" -->
</head>

<body onLoad="language()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_User"></span>:
      <% =session("User") %></td>
  </tr>
  <tr>
    <td height="20" colspan="9"><span id="inner_SortTitle"></span>:
<input name="ByNormal" type="radio" value="1" onClick="javascript:location.href='StoreNoYield.asp'">
<span id="inner_SortNormal"></span>       <input name="ByPart" type="radio" value="1" onClick="javascript:location.href='StoreNoYieldbyPartNumber.asp'">
<span id="inner_SortPartNumber"></span>
</td>
  </tr>
<form name="formSort" method="post" action="">

<input name="store_nids" type="hidden" value="">
<input name="family_name" type="hidden" value="">
</form>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_NoYieldCheck"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_JobNumber"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_InputQuantity"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_StoreQuantity"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Yield"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_StoreType"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_StoreTime"></span></div></td>
  </tr>
  <%
i=1
if not rs.eof then
dim this_nid()
dim this_jobnumber()
dim this_jobNoYield()
dim this_partnumbertag()
dim this_input_quantity()
dim this_output_quantity()
dim this_yield()
dim this_store_type()
dim this_store_time()
redim this_nid(0)
redim this_jobnumber(0)
redim this_NoYield(0)
redim this_partnumbertag(0)
redim this_input_quantity(0)
redim this_output_quantity(0)
redim this_yield(0)
redim this_store_type(0)
redim this_store_time(0)
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
		this_nid(UBound(this_nid))=rs("NID")
		this_jobnumber(UBound(this_jobnumber))=rs("JOB_NUMBER")
		this_NoYield(UBound(this_jobnumber))=rs("NO_YIELD")
		this_partnumbertag(UBound(this_partnumbertag))=rs("PART_NUMBER_TAG")
		this_input_quantity(UBound(this_input_quantity))=rs("INPUT_QUANTITY")
		this_output_quantity(UBound(this_output_quantity))=rs("STORE_QUANTITY")
		if rs("INPUT_QUANTITY")<>"0" then
		this_yield(UBound(this_yield))=cint(rs("STORE_QUANTITY"))/cint(rs("INPUT_QUANTITY"))
		else
		this_yield(UBound(this_yield))=0
		end if
		this_store_type(UBound(this_store_type))=rs("STORE_TYPE")
		this_store_time(UBound(this_store_time))=rs("STORE_TIME")
		ReDim Preserve this_nid(UBound(this_nid)+1)
		ReDim Preserve this_jobnumber(UBound(this_jobnumber)+1)
		ReDim Preserve this_NoYield(UBound(this_NoYield)+1)
		ReDim Preserve this_partnumbertag(UBound(this_partnumbertag)+1)
		ReDim Preserve this_input_quantity(UBound(this_input_quantity)+1)
		ReDim Preserve this_output_quantity(UBound(this_output_quantity)+1)
		ReDim Preserve this_yield(UBound(this_yield)+1)
		ReDim Preserve this_store_type(UBound(this_store_type)+1)
		ReDim Preserve this_store_time(UBound(this_store_time)+1)
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
    <td>&nbsp;</td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=current_jobnumber%>')">
      <input name="tabstatus<%=current_jobnumber%>" type="hidden" value="0">
      <img src="/Images/Treeimg/plus.gif" name="tabimg<%=current_jobnumber%>" width="9" height="9">&nbsp;<%= current_jobnumber %></span></div></td>
    <td height="20"><div align="center"><%=current_partnumber%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_input_quantity%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_output_quantity%></div></td>
    <td height="20"><div align="center"><%=formatpercent(total_yield,2,-1)%>&nbsp;</div></td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
  </tr>
  <tbody id="tab<%=current_jobnumber%>" style="display:none">
    <%
			for u=0 to ubound(this_jobnumber)-1 'Display detail of each sheet.
%>
    <tr>
      <td height="20"><div align="center"><%=u+1%></div></td>
      <td><div align="center">
        <input name="NoYieldcheck<%=i%>" type="checkbox" id="NoYieldcheck<%=i%>" value="<%=this_nid(u)%>" <%if this_NoYield(u)="1" then%>checked<%end if%> onClick="NoYieldcheck(this,this.value)">
      </div></td>
      <td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=this_jobnumber(u)%>" target="_blank"><%= this_jobnumber(u) %></a></div></td>
      <td height="20"><div align="center"><%= this_partnumbertag(u) %>&nbsp;</div></td>
      <td height="20"><div align="center"><% =this_input_quantity(u)%></div></td>
      <td height="20"><div align="center"><%=this_output_quantity(u)%></div></td>
      <td height="20"><div align="center"><%=formatpercent(this_yield(u),2,-1)%>&nbsp;</div></td>
      <td height="20"><div align="center"><%=this_store_type(u)%></div></td>
      <td height="20"><div align="center"><%=this_store_time(u)%></div></td>
    </tr>
    <%		next%>
  </tbody>
  <%
		redim this_nid(0)
		redim this_jobnumber(0)
		redim this_NoYield(0)
		redim this_partnumbertag(0)
		redim this_input_quantity(0)
		redim this_output_quantity(0)
		redim this_yield(0)
		redim this_store_type(0)
		redim this_store_time(0)
		total_input_quantity=0
		total_output_quantity=0
		this_nid(UBound(this_nid))=rs("NID")
		this_jobnumber(UBound(this_jobnumber))=rs("JOB_NUMBER")
		this_NoYield(UBound(this_NoYield))=rs("NO_YIELD")
		this_partnumbertag(UBound(this_partnumbertag))=rs("PART_NUMBER_TAG")
		this_input_quantity(UBound(this_input_quantity))=rs("INPUT_QUANTITY")
		this_output_quantity(UBound(this_output_quantity))=rs("STORE_QUANTITY")
		if rs("INPUT_QUANTITY")<>"0" then
		this_yield(UBound(this_yield))=cint(rs("STORE_QUANTITY"))/cint(rs("INPUT_QUANTITY"))
		else
		this_yield(UBound(this_yield))=0
		end if
		this_store_type(UBound(this_store_type))=rs("STORE_TYPE")
		this_store_time(UBound(this_store_time))=rs("STORE_TIME")
		ReDim Preserve this_nid(UBound(this_nid)+1)
		ReDim Preserve this_jobnumber(UBound(this_jobnumber)+1)
		ReDim Preserve this_NoYield(UBound(this_NoYield)+1)
		ReDim Preserve this_partnumbertag(UBound(this_partnumbertag)+1)
		ReDim Preserve this_input_quantity(UBound(this_input_quantity)+1)
		ReDim Preserve this_output_quantity(UBound(this_output_quantity)+1)
		ReDim Preserve this_yield(UBound(this_yield)+1)
		ReDim Preserve this_store_type(UBound(this_store_type)+1)
		ReDim Preserve this_store_time(UBound(this_store_time)+1)
		total_input_quantity=total_input_quantity+cint(rs("INPUT_QUANTITY"))
		total_output_quantity=total_output_quantity+cint(rs("STORE_QUANTITY"))
		if total_input_quantity<>0 then
		total_yield=total_output_quantity/total_input_quantity
		else
		total_yield=0
		end if
		t=t+1
	end if
	all_input=all_input+cint(rs("INPUT_QUANTITY"))
	all_output=all_output+cint(rs("STORE_QUANTITY"))
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
    <td>&nbsp;</td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=current_jobnumber%>')">
      <input name="tabstatus<%=current_jobnumber%>" type="hidden" value="0">
      <img src="/Images/Treeimg/plus.gif" name="tabimg<%=current_jobnumber%>" width="9" height="9">&nbsp;<%= current_jobnumber %></span></div></td>
    <td height="20"><div align="center"><%=current_partnumber%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_input_quantity%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_output_quantity%></div></td>
    <td height="20"><div align="center">
    <%=formatpercent(total_yield,2,-1)%>&nbsp;</div></td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
  </tr>
  <tbody id="tab<%=current_jobnumber%>" style="display:none">
    <%
		for u=0 to ubound(this_jobnumber)-1 'Display detail of each sheet.
%>
    <tr>
      <td height="20"><div align="center"><%=u+1%></div></td>
      <td><div align="center">
        <input name="NoYieldcheck<%=i%>" type="checkbox" id="NoYieldcheck<%=i%>" value="<%=this_nid(u)%>" <%if this_NoYield(u)="1" then%>checked<%end if%> onClick="NoYieldcheck(this,this.value)">
      </div></td>
      <td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=this_jobnumber(u)%>" target="_blank"><%= this_jobnumber(u) %></a></div></td>
      <td height="20"><div align="center"><%= this_partnumbertag(u) %>&nbsp;</div></td>
      <td height="20"><div align="center"><%=this_input_quantity(u)%></div></td>
      <td height="20"><div align="center"><%=this_output_quantity(u)%></div></td>
      <td height="20"><div align="center"><%=formatpercent(this_yield(u),2,-1)%>&nbsp;</div></td>
      <td height="20"><div align="center"><%=this_store_type(u)%></div></td>
      <td height="20"><div align="center"><%=this_store_time(u)%></div></td>
    </tr>
	<% next%>
  </tbody>
	
  <%
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records </div></td>
  </tr>
  <%end if
rs.close
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->