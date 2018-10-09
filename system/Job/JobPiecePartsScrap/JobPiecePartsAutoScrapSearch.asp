<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
	MaterialPartNumber=request("txtMaterialPartNumber")
	Opcode=request("txtOpcode")
	LineName=request("txtLineName")
	ERP_Account=request("ERP_Account")
	ERP_Reason=request("ERP_Reason")
	ERP_Refer=request("ERP_Refer")
 
	fromdate=request("fromdate")
	fromtime=request("fromtime")
	todate=request("todate")
	totime=request("totime")
	ReportType=request("ReportType")

	time0=now   
	if isnull(fromdate) or fromdate=""  then
		fromdate=cstr(year(dateadd("d",-weekday(time0)+1,time0))) +"-"+cstr(month(dateadd("d",-weekday(time0)+1,time0)))+"-"+cstr(day(dateadd("d",-weekday(time0)+1,time0)))
	end if
	
	if isnull(fromtime) or fromtime=""  then
		fromtime="14:30:00"
	end if
	todate=request("todate")
	if isnull(todate) or todate=""  then
		todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
	end if
	totime=request("totime")
	if isnull(totime) or totime="" then
		totime="14:30:00"
	end if
	if(request.QueryString("Action")="2") then
		SQLStr="SELECT * FROM PIECE_PARTS_SCRAP WHERE SCRAP_TYPE=1"
		if(MaterialPartNumber<>"")then
			SQLStr=SQLStr+" AND MATERIAL_PART_NUMBER='"+MaterialPartNumber+"'"
		end if 
		if(Opcode<>"")then
			SQLStr=SQLStr+" AND SCRAP_OP_CODE='"+Opcode+"'"
		end if
		if(LineName<>"")then
			SQLStr=SQLStr+" AND LINE_NAME='"+LineName+"'"
		end if
 
		if(fromdate<>"")then
			SQLStr=SQLStr+" AND SCRAP_DATETIME>=to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		end if
		
		if(todate<>"")then
			SQLStr=SQLStr+" AND SCRAP_DATETIME<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		end if
 

		if(ReportType="1")then
			SQLStr=SQLStr+" AND IS_SENT_APPROVE='0'"
		end if 
		if(ReportType="2")then
			SQLStr=SQLStr+" AND IS_SENT_APPROVE='1'"
		end if 
		
		SQLStr=SQLStr+" order by SCRAP_DATETIME desc,MATERIAL_PART_NUMBER"
		rs.open SQLStr,conn,1,3
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="javascript">
 	 
	function SearchData()
	{
		form1.action="JobPiecePartsAutoScrapSearch.asp?Action=2";
		form1.submit();
	}
</script>
</head>

<body>
<form method="post" name="form1" target="_self" >
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
   <td height="20" colspan="8" class="t-b-midautumn"> 
		报废记录查询
	</td>
  </tr>
 <tr>
    <td height="20">料号</td>
    <td height="20" colspan="6">
	<input type="text" name="txtMaterialPartNumber" id="txtMaterialPartNumber" value="<%=MaterialPartNumber%>">
	</td>
  </tr>
  
   <tr>
    <td height="20">工号</td>
    <td height="20" colspan="6">
	<input type="text" name="txtOpcode" id="txtOpcode" value="<%=OpCode%>">
	</td>
  </tr> 
   <tr>
    <td height="20">线别</td>
    <td height="20" colspan="6">
	<input type="text" name="txtLineName" id="txtLineName" value="<%=LineName%>">
	</td>
  </tr>
   
  <Tr>
  	 <td>报废时间 从:</td> 
	 <td align="left"><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
	
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
	<select name="fromtime" id="fromtime">
   			 <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
</td>
<td align="left">到:</td>
<td colspan="4" align="left">
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
	<select name="totime" id="totime">
   			 <option value="14:30:00" <% if totime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if totime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
&nbsp;
<td>
<tr>
<td height="20">
	报表类型    
</td>
    <td height="20" colspan="6"><label>
      <select name="ReportType" id="ReportType">
	   <option value="1" <%if ReportType="1" then response.write "selected" end if %>>未送审批</option>
	   <option value="0" <%if ReportType="0" then response.write "selected" end if %>>所有</option>
	   <option value="2" <%if ReportType="2" then response.write "selected" end if %>>已送审批</option>
      </select>
	  <br><span id="errorinsertERPreason" class="red"></span>
    </label></td>
  </tr>
  
  
    <tr>
    <td height="20" colspan="7">
	<input type="button" name="btnSearch" id="btnSearch" value="查询" onclick="SearchData()"  class="t-b-Yellow">
	</td>
  </tr>
</table>

<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr>
		<td class="t-t-Borrow">原材料料号</td>
		<td class="t-t-Borrow">报废人</td>
		<td class="t-t-Borrow">报废数量</td>
		<td class="t-t-Borrow">线别</td>
		<td class="t-t-Borrow">报废时间</td>
		<td class="t-t-Borrow">是否送审批</td>
	</tr>
	<%if(request.QueryString("Action")="2") then 
		while not rs.eof
	%>
		<tr>
			<td><%=rs("MATERIAL_PART_NUMBER")%></td>
			<td><%=rs("SCRAP_OP_CODE")%></td>
			<td><%=rs("QTY")%></td>
			<td><%=rs("LINE_NAME")%></td>
			<td><%=rs("SCRAP_DATETIME")%></td>
			<td><% if rs("IS_SENT_APPROVE")="0" then response.write "未送审批" else response.write "已送审批" end if %></td>
		</tr>
	<%
			rs.movenext
		wend
		end if %>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->