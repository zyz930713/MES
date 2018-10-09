<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
store_nids=request.Form("store_nids")
family_name=trim(request.Form("family_name"))
pagename="/Reports/FinalYield/FinalFamilyYield/FinalFamilyDetailbyPartNumber.asp.asp"
SQL="select * from JOB_MASTER_STORE where instr('"&store_nids&"',NID)>0 order by PART_NUMBER_TAG,JOB_NUMBER"
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
    <td height="20" colspan="8" class="t-c-greenCopy">Browse Final Series Yield -- <%=family_name%></td>
  </tr>
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy">User:
      <% =session("User") %></td>
  </tr>
<form name="formSort" method="post" action="">
<tr>
  <td height="20" colspan="8">Sort by: 
    <input name="ByJob" type="radio" value="1" onClick="document.formSort.action='FinalFamilyDetailbyJobNumber.asp';document.all.store_nids.value='<%=store_nids%>';document.all.family_name.value='<%=family_name%>';document.formSort.submit()">
    Job Number
    <input name="ByNormal" type="radio" value="1" onClick="document.formSort.action='FinalFamilyDetail.asp';document.all.store_nids.value='<%=store_nids%>';document.all.family_name.value='<%=family_name%>';document.formSort.submit()">
    Normal </td>
</tr>
<input name="store_nids" type="hidden" value="">
<input name="family_name" type="hidden" value="">
</form>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Job Number</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Part Number</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">InputQuantity </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Store Quantity</div></td>
    <td height="20" class="t-t-Borrow"><div align="center"> Yield </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Store Type </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Store Time </div></td>
  </tr>
  <%
i=1
if not rs.eof then
dim this_jobnumber()
dim this_partnumbertag()
dim this_input_quantity()
dim this_output_quantity()
dim this_yield()
dim this_store_type()
dim this_store_time()
redim this_jobnumber(0)
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
	if rs("PART_NUMBER_TAG")=current_partnumber then 'Get total parameters' value of all the same jobs
		this_jobnumber(UBound(this_jobnumber))=rs("JOB_NUMBER")
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
		ReDim Preserve this_jobnumber(UBound(this_jobnumber)+1)
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
	else
	%>
  <tr class="t-b-blue">
    <td height="20"><div align="center">
      <% =t%>
    </div></td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=replace(current_partnumber,"-","_")%>')">
      <input name="tabstatus<%=replace(current_partnumber,"-","_")%>" type="hidden" value="0">
      <img src="/Images/Treeimg/plus.gif" name="tabimg<%=replace(current_partnumber,"-","_")%>" width="9" height="9">&nbsp;<%= current_partnumber %></span></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_input_quantity%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_output_quantity%></div></td>
    <td height="20"><div align="center"><%=formatpercent(total_yield,2,-1)%>&nbsp;</div></td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
  </tr>
  <tbody id="tab<%=replace(current_partnumber,"-","_")%>" style="display:none">
    <%
			for u=0 to ubound(this_jobnumber)-1 'Display detail of each sheet.
%>
    <tr>
      <td height="20"><div align="center"><%=u+1%></div></td>
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
		redim this_jobnumber(0)
		redim this_partnumbertag(0)
		redim this_input_quantity(0)
		redim this_output_quantity(0)
		redim this_yield(0)
		redim this_store_type(0)
		redim this_store_time(0)
		total_input_quantity=0
		total_output_quantity=0
		this_jobnumber(UBound(this_jobnumber))=rs("JOB_NUMBER")
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
		ReDim Preserve this_jobnumber(UBound(this_jobnumber)+1)
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
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=replace(current_partnumber,"-","_")%>')">
      <input name="tabstatus<%=replace(current_partnumber,"-","_")%>" type="hidden" value="0">
      <img src="/Images/Treeimg/plus.gif" name="tabimg<%=replace(current_partnumber,"-","_")%>" width="9" height="9">&nbsp;<%= current_partnumber %></span></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_input_quantity%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_output_quantity%></div></td>
    <td height="20"><div align="center">
    <%=formatpercent(total_yield,2,-1)%>&nbsp;</div></td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
  </tr>
  <tbody id="tab<%=replace(current_partnumber,"-","_")%>" style="display:none">
    <%
		for u=0 to ubound(this_jobnumber)-1 'Display detail of each sheet.
%>
    <tr>
      <td height="20"><div align="center"><%=u+1%></div></td>
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
	<tr>
      <td height="20">&nbsp;</td>
      <td height="20"><div align="center">Total</div></td>
      <td height="20">&nbsp;</td>
      <td height="20"><div align="center"><%=all_input%>&nbsp;</div></td>
      <td height="20"><div align="center"><%=all_output%></div></td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
    </tr>
  <%
else
%>
  <tr>
    <td height="20" colspan="8"><div align="center">No Records </div></td>
  </tr>
  <%end if
rs.close
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->