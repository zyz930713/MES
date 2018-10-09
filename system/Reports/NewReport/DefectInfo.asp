<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->


<%
	Family= request.querystring("Family") 
	Series= request.querystring("Series") 
	SubSeries= request.querystring("SubSeries") 
	Model= request.querystring("Model") 
	JobNumber=request.querystring("Job_Number")
	SheetNumber=request.querystring("Sheet_Number")
	
	SQL= "select FA.FAMILY_NAME,SN.SERIES_NAME" 
	where=""
	groupBy=",SN.SERIES_NAME"
	if Series<>"" then		
		where=where + " and SN.SERIES_NAME='"&Series&"'"		
	end if
	
	if SubSeries<>"" then
		SQL=SQL+" ,SS.SUBSERIES_NAME"
		where=where + " and SS.SUBSERIES_NAME='"&SubSeries&"'"
		groupBy=groupBy+" ,SS.SUBSERIES_NAME"
	end if
	
	if Model<>"" then
		SQL=SQL+" ,TEMP.part_number_tag"
		where=where + " and TEMP.part_number_tag='"&Model&"'"
		groupBy=groupBy+" ,TEMP.part_number_tag"
	end if
	if JobNumber<>"" and SheetNumber<>"" then
		SQL=SQL+" ,TEMP.job_number,TEMP.sheet_number "
		where=where + " and TEMP.job_number='"&JobNumber&"' and TEMP.sheet_number='"&SheetNumber&"'"
		groupBy=groupBy+" ,TEMP.job_number,TEMP.sheet_number"
	end if
	
	SQL=SQL+"  ,temp.defect_name,sum(temp.defect_quantity) as Quantity"
	
	SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM,( "
	SQL=SQL+ "SELECT j.part_number_tag,df.defect_code||' - '||df.defect_name||'('||df.defect_chinese_name||')' as defect_name,jd.defect_quantity,j.job_number,j.sheet_number"
	SQL=SQL+" FROM job_defectcodes jd,job j,defectcode df "
	SQL=SQL+"  WHERE jd.job_number = j.job_number AND jd.sheet_number = j.sheet_number AND jd.defect_code_id = df.nid(+) "
	SQL=SQL+" AND j.close_time > to_date('"& session("FromDateTime") &"','yyyy-mm-dd hh24:mi:ss') "
	SQL=SQL+" AND j.close_time <= to_date('"& session("ToDateTime") &"','yyyy-mm-dd hh24:mi:ss') "
	SQL=SQL+" AND J.JOB_TYPE='N'"
	SQL=SQL+" ) temp "
	SQL=SQL+" where temp.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID "
	SQL=SQL+" AND FA.FAMILY_NAME='"&Family&"'"
	SQL=SQL+ where +" GROUP BY FAMILY_NAME " + groupBy	
	SQL=SQL+" ,TEMP.defect_name"
	SQL=SQL+" order BY FAMILY_NAME,sum(temp.defect_quantity) desc "		
	
	'RESPONSE.WRITE SQL
	'response.end
	rs.open SQL,conn,1,3
				
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script>
	function LoadNextItem()
	{
		form1.action="DefectReport.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}

		form1.action="DefectReport.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>
<body onLoad="language(<%=session("language")%>);">
<table width="100%" align="center"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="18">&nbsp; </td>
  </tr>
<%if(rs.State > 0 ) then %>  
  <tr>
  	<td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Family"></span></div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Series"></span></div></td>
	<%for i=0 to rs.Fields.count-1
	  spanId=""
	  if rs.Fields(i).name ="SUBSERIES_NAME" then
		spanId="td_SubSeries"
	  elseif rs.Fields(i).name ="PART_NUMBER_TAG" then
		spanId="td_Model"
	  elseif rs.Fields(i).name ="JOB_NUMBER" then
		spanId="td_JobNumber"
	  elseif rs.Fields(i).name ="SHEET_NUMBER" then
		spanId="td_SheetNumber"					
	  end if
	  if spanId <> "" then
	%> 
		<td height="20" class="t-t-Borrow"><div align="center"><span id='<%=spanId%>'></span></div></td>
  	<%end if
	next%>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_DefectCode"></span></div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Quantity"></span></div></td>	
</tr>
	<% for j=0 to rs.recordcount-1%>
	<tr align="center"><td><%=(j+1)%></td>
	<%			
				for i=0 to rs.Fields.count-1
	%>
				 <td height="20" ><div align="center"><%=rs(i).value%> </div></td>
				 
  	<%			
				next
				rs.movenext
			next 
	%>
		</tr>
	<%
		end if
	%>
	
  </tr>
 
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->