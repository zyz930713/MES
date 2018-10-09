<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->

<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->

<%
fromdate=request("fromdate")
todate=request("todate")
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
todate=request("todate")
if isnull(todate) or todate=""  then	
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if

code=trim(request("code"))

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>

<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">

</head>

<body onLoad="language_page();language(<%=session("language")%>);document.all.code.select();" >
<form action="/Reports/Production/WIIReport.asp" method="get" name="form1" target="_self">
               


<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="4" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td width="14%" height="20"><span id="inner_2DCode"></span></td>
    <td width="19%" height="20"><input name="code" type="text" id="code" value="<%=code%>" size="25">  </td>
	<td width="50%" height="20"><a href="2DcodeList.asp?Barcode=<%=code%>"><span id="inner_PTC"></span></a></td>
    <td width="17%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">	</td>
    </tr>
</table>




</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/Packing_Plan/addPacking_Plan.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="11"></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
 
  
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_JobNumber"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Stations"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_OpCode"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_StartTime"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_CloseTime"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Remark"></span></div></td>
  </tr>
  
  
<!---------------------------------------------------取得Routing信息的数据：------------------------------------------------------>  
  
<%
if code <> "" then
sql="select a.job_number,a.sheet_number,b.defect_chinese_name as defect from job_2d_code a left join defectcode b on a.defect_code_id=b.nid where code='"&code&"'"
rs.open sql,conn,1,3
if not rs.eof then
	jobNumber=rs("job_number")
	sheetNumber=rs("sheet_number")
	defect=rs("defect")
	subJobNumber=jobNumber&"-"&repeatstring(sheetNumber,"0",3)
end if
rs.close

sql="select * from job_master where job_number='"&jobNumber&"'"
rs.open sql,conn,1,3
if not rs.eof then
  part_number_tag=rs("part_number_tag")
end if
rs.close

sql="select b.station_name||' '||b.station_chinese_name as station_name,a.operator_code,a.start_time,a.close_time from job_stations a,station b where a.station_id=b.nid and a.job_number='"&jobNumber&"' and a.sheet_number='"&sheetNumber&"' order by a.start_time"
rs.open SQL,conn,1,3
i=1
if not rs.eof then

do while not rs.eof 
%>
<tr>
  <td height="20"><div align="center">
    <% =i%>
  </div>  </td>        
	<td><a href="EMR_List.asp?job_number=<%=jobNumber%>&sheet_number=<%=sheetNumber%>&partnumber=<%=part_number_tag%>"><%= subJobNumber%></a>&nbsp;</td>
	<td><%= rs("station_name")%>&nbsp;</td>
	<td><%= rs("operator_code")%>&nbsp;</td>
	<td><%= rs("start_time")%>&nbsp;</td>
	<td><%= rs("close_time")%>&nbsp;</td>
	<td><%if trim(rs("station_name"))=trim("700-Acoustic Test 声学测试")   then
		%><%=defect%>&nbsp;
		<a href="PVSTestData.asp?code=<%=code%>" >PVSTestData</a>&nbsp;<a href="TSDTestData.asp?code=<%=code%>" >TSD-TestData</a>
		<%if not isnull(defect) then
			i=i+1
			exit do
			end if
        elseif trim(rs("station_name"))=trim("711-FOI 外观检查") then
            response.Write(defect)
		end if%>
	&nbsp;</td>   
  </tr>
<%
i=i+1
rs.movenext
loop
else
%>
  <tr>
    <td height="20" colspan="11" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close

if jobNumber<>"" and isnull(defect) then
%>
<!---------------------------------------------------取得Label Print的数据------------------------------------------------------>
<%

sql="select 'Label Print' as station_name,b.operator,b.printtime as start_time,b.printtime as close_time from label_print_history b where b.job_number='"&jobNumber&"' and '"&sheetNumber&"' in (select column_value from table(cast(CHAR_SPLIT(b.subjoblist,'-') as char_table))) and b.transactiontype='-1'"
rs.open SQL,conn,1,3
if not rs.eof then

while not rs.eof 
%>
<tr>
  <td height="20"><div align="center">
    <% =i%>
  </div>  </td>
 	<td><%= subJobNumber%>&nbsp;</td>
	<td><%= rs("station_name")%>&nbsp;</td>
	<td><%= rs("operator")%>&nbsp;</td>
	<td><%= rs("start_time")%>&nbsp;</td>
	<td><%= rs("close_time")%>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
<%
rs.movenext
wend
else
%>
   <tr>
   <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div>  </td>    
	<td><%= subJobNumber%>&nbsp;</td>
	<td><%="Label Print"%>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>   
  </tr>
<%end if
i=i+1
rs.close%>


<!---------------------------------------------------取得OQC的数据------------------------------------------------------>

<%

sql="select 'OQC' as station_name,c.startoperator as operator,c.oqcstarttime as start_time,c.oqcendtime as close_time from label_print_history b,job_oqc c where b.job_number='"&jobNumber&"' and '"&sheetNumber&"' in (select column_value from table(cast(CHAR_SPLIT(b.subjoblist,'-') as char_table))) and instr(c.batchnolist,b.batchno)>0 "
rs.open SQL,conn,1,3
if not rs.eof then

while not rs.eof 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div>  </td>
	<td><%= subJobNumber%>&nbsp;</td>
	<td><%= rs("station_name")%>&nbsp;</td>
	<td><%= rs("operator")%>&nbsp;</td>
	<td><%= rs("start_time")%>&nbsp;</td>
	<td><%= rs("close_time")%>&nbsp;</td>
	<td>&nbsp;</td>
   
  </tr>
<%

rs.movenext
wend
else
%>
   <tr>
   <td height="20"><div align="center">
    <% =i%>
  </div>  </td>    
	<td><%= subJobNumber%>&nbsp;</td>
	<td><%="OQC"%>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>   
  </tr>
<%end if
i=i+1
rs.close%>





<!---------------------------------------------------取得Pre Store的资料------------------------------------------------------>


<%

sql="select 'Pre Store' as station_name,replace(a.store_code,'-0000','') as operator,a.store_time as start_time,a.store_time as close_time from job_master_store_pre a where a.job_number='"&jobNumber&"' and instr(a.sub_job_numbers,'"&subJobNumber&"')>0 "
rs.open SQL,conn,1,3
if not rs.eof then

while not rs.eof 
%>
<tr>
  <td height="20"><div align="center">
    <% =i%>
  </div>  </td>
 
    
	<td><%=subJobNumber%>&nbsp;</td>
	<td><%= rs("station_name")%>&nbsp;</td>
	<td><%= rs("operator")%>&nbsp;</td>
	<td><%= rs("start_time")%>&nbsp;</td>
	<td><%= rs("close_time")%>&nbsp;</td>
	<td>&nbsp;</td>
   
  </tr>
<%

rs.movenext
wend
else
%>
  <tr>
   <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div>  </td>    
	<td><%= subJobNumber%>&nbsp;</td>
	<td><%="Pre Store"%>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>   
  </tr>
<%end if
rs.close
i=i+1
end if
%>




<!---------------------------------------------------取得Packing的资料------------------------------------------------------>


<%

sql="select 'Packing' as station_name,pack_user as operator,pack_time as start_time,pack_time as close_time,box_id from job_2d_code where code='"&code&"' and is_packed =1 "
rs.open SQL,conn,1,3
if not rs.eof then

while not rs.eof 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div>  </td>
 
    
	<td><%= subJobNumber%>&nbsp;</td>
	<td><%= rs("station_name")%>&nbsp;</td>
	<td><%= rs("operator")%>&nbsp;</td>
	<td><%= rs("start_time")%>&nbsp;</td>
	<td><%= rs("close_time")%>&nbsp;</td>
	<td><%if rs("box_id")<>"" then%><span id="inner_BoxID"></span>:<%= rs("box_id")%><%end if%>&nbsp;</td>
   
  </tr>
<%
rs.movenext
wend
else
%>
  <tr>
   <td height="20"><div align="center">
    <% =i%>
  </div>  </td>    
	<td><%= subJobNumber%>&nbsp;</td>
	<td><%="Packing"%>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>   
  </tr>
<%end if
rs.close
i=i+1%>


<!---------------------------------------------------取得Shipping的资料------------------------------------------------------>


<%
if isnull(defect) then
sql="select get_sub_job_number(a.job_number,a.sheet_number) as job_number,'Shipping' as station_name,b.shipped_user as operator,b.stack_time as start_time,b.shipped_time as close_time,b.pallet_id,b.DELIVERY_ID from job_2d_code a,job_pack_detail b where a.box_id=b.box_id and a.job_number=b.job_number and  b.is_shipped=1 and b.WHREC_TIME is not NULL and code='"&code&"'"
rs.open SQL,conn,1,3
if not rs.eof then

while not rs.eof 
%>
<tr>
  <td height="20"><div align="center">
    <% =i%>
  </div>  </td>
 
    
	<td><%= subJobNumber%>&nbsp;</td>
	<td><%= rs("station_name")%>&nbsp;</td>
	<td><%= rs("operator")%>&nbsp;</td>
	<td><%= rs("start_time")%>&nbsp;</td>
	<td><%= rs("close_time")%>&nbsp;</td>
	<td><%if rs("pallet_id")<>"" then%><span id="inner_PalletID"></span>:<%=rs("pallet_id")%><%end if%><% if rs("DELIVERY_ID") <>"" then %>,<span id="td_DeliveryID"></span>:<%=rs("DELIVERY_ID")%><%end if%>&nbsp;</td>
   
  </tr>
<%
rs.movenext
wend
else
%>
  <tr>
   <td height="20"><div align="center">
    <% =i%>
  </div>  </td>    
	<td><%= subJobNumber%>&nbsp;</td>
	<td><%="Shipping"%>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>   
  </tr>
<%end if
rs.close
i=i+1
end if
%>
<%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->