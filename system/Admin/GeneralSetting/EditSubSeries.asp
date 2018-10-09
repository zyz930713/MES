<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
id=request.QueryString("id")

'SQL="SELECT * FROM GENERAL_SETTING WHERE SUBSERIESNAME='"&id&"' and( MODELNAME is not null or MODELNAME<>'')"
SQL="select * from "
SQL=SQL+" ( "
SQL=SQL+" SELECT * FROM GENERAL_SETTING WHERE SUBSERIESNAME='"+id+"' and( MODELNAME is not null or MODELNAME<>'')"
SQL=SQL+" union "
SQL=SQL+" select '-100','"+id+"',item_name,'','','','','','','','','','','','','','','','','',sysdate,sysdate,'','',sysdate,'','','','',''"
SQL=SQL+" from product_model where series_id=(select nid from subseries where subseries_name='"+id+"') "
SQL=SQL+" and item_name not in (SELECT modelname FROM GENERAL_SETTING WHERE SUBSERIESNAME='"+id+"' and( MODELNAME is not null or MODELNAME<>'') )  "
SQL=SQL+" ) a"
SQL=SQL+" order by gnid"


'response.Write SQL
'response.End()
rs.open SQL,conn,1,3	

if request.QueryString("action")="1" then
	id=request.Form("seriesName")
	counts=request.Form("counts")
	'delete old date
	SQL="Delete from GENERAL_SETTING where SUBSERIESNAME='"&id&"' and ( MODELNAME is not null or MODELNAME<>'') "
	set rsD=server.createobject("adodb.recordset")
	rsD.open SQL,conn,1,3
	'insert new date
	for i=1 to counts
		SQL="SELECT * FROM GENERAL_SETTING WHERE 1=2 "
	 
		rsD.open SQL,conn,1,3
		rsD.addnew
	
		ModelName=request("Model"&i)
		RetestYield=request("RETESTYIELD"&i)
		'RetestRatio=request("RETESTRATIO"&i)
		RetestRatio=0
		Remark1=request("RemarkY"&i)
		SlappingHoldDays=request("SlappingHoldDays"&i)
		Remark2=request("RemarkSlapping"&i)
		SpecialExamine1=request("SEA"&i)
		AQL1=request("QA"&i)
		SpecialExamine2=request("SEB"&i)
		AQL2=request("QB"&i)
		SpecialExamine3=request("SEC"&i)
		AQL3=request("QC"&i)
		SpecialExamine4=request("SED"&i)
		AQL4=request("QD"&i)
		Suffix=request("Suffix")
		CurrentAQL=request("CurrentAQL"&i)
		MaxAQL=request("MaxAQL"&i)
		MinAQL=request("MinAQL"&i)
		Owner=request("Owner"&i)
		Suffix=request("Suffix"&i)
		
		Appearance=request("APPEARANCE"&i)
		Performance=request("PERFORMANCE"&i)
		MoreOQC=request("MOREOQC"&i)
		
		rsD("GNID")="GS"&NID_SEQ("GENERAL")
		rsD("SubSeriesName")=id
		rsD("MODELNAME")=ModelName
		rsD("RETESTYIELD")=RetestYield
		rsD("RETESTRATIO")=RetestRatio
		rsD("REMARK1")=REMARK1
		rsD("SLAPPINGHOLDDAYS")=SLAPPINGHOLDDAYS
		rsD("REMARK2")=REMARK2
		rsD("SPECIALEXAMINE1")=SPECIALEXAMINE1
		rsD("AQL1")=AQL1
		rsD("SPECIALEXAMINE2")=SPECIALEXAMINE2
		rsD("AQL2")=AQL2
		rsD("SPECIALEXAMINE3")=SPECIALEXAMINE3
		rsD("AQL3")=AQL3
		rsD("SPECIALEXAMINE4")=SPECIALEXAMINE4
		rsD("AQL4")=AQL4
		rsD("SUFFIX")=Suffix
		rsD("CURRENTAQL")=CURRENTAQL
		rsD("MAXAQL")=MAXAQL
		rsD("MINAQL")=MinAQL
		rsD("VALIDSTARTTIME")=now()
		rsD("ISEXPIRED")="0"
		rsD("LASTUPDATECODE")=session("code")
		rsD("Owner")=Owner
		rsD("APPEARANCE")=Appearance
		rsD("PERFORMANCE")=Performance
		rsD("MOREOQC")=MoreOQC
		rsD.update
		rsD.close()
	next
	response.write "<script>window.alert('Update SubSeries Successfully!');window.close();</script>"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
<base target="_self"/>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
	function saveDate()
	{
		document.form1.action="EditSubSeries.asp?action=1";
		document.form1.submit();
	}
</script>
</head>
<body onload="language(<%=session("language")%>);" >

<form method="post" name="form1" action="">
<table width="90%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">

<tr>
  <td height="24" class="t-t-Borrow" ><div align="center"><span id="inner_NO"></span></div></td> 
  <td class="t-t-Borrow" ><div align="center"><span id="td_SubSeries"></span></div></td>
  <td class="t-t-Borrow" ><div align="center"><span id="inner_PartNumber"></span></div></td>
  <!--
  <td class="t-t-Borrow" ><div align="center">retest 良率</div></td>
  <td class="t-t-Borrow"><div align="center">retest 比率</div></td> 
  -->
   <td class="t-t-Borrow"><div align="center"><span id="td_ChangeReason"></span></div></td>
   <!--
   <td class="t-t-Borrow"><div align="center">Slapping放置天数</div></td>
   <td class="t-t-Borrow"><div align="center">Slapping 良率</div></td>   
   <td class="t-t-Borrow"><div align="center">修改原因</div></td>   
   <td class="t-t-Borrow"><div align="center">特殊检验相目 1</div></td>
   <td class="t-t-Borrow"><div align="center">检验水准</div></td>
   <td class="t-t-Borrow"><div align="center">特殊检验相目 2</div></td>
   <td class="t-t-Borrow"><div align="center">检验水准</div></td>
   <td class="t-t-Borrow"><div align="center">特殊检验相目 3</div></td>
   <td class="t-t-Borrow"><div align="center">检验水准</div></td>
   <td class="t-t-Borrow"><div align="center">特殊检验相目 4</div></td>
   <td class="t-t-Borrow"><div align="center">检验水准</div></td>
  <td class="t-t-Borrow"><div align="center">Suffix</div></td>
  -->
  <td class="t-t-Borrow"><div align="center"><span id="td_CurrentAQL"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_MaxAQL"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_MinAQL"></span></div></td>
  <!--
  <td class="t-t-Borrow"><div align="center">Owner </div></td>
  <td class="t-t-Borrow"><div align="center">外观</div></td>
  <td class="t-t-Borrow"><div align="center">性能</div></td>
  <td class="t-t-Borrow"><div align="center">是否2次OQC</div></td>
  -->
</tr>
<%
	i=0
	while not rs.eof
		i=i+1
%>
<tr>
  <td height="24" align="center" ><%=i%></td>  
  <td align="center"><%=rs("SUBSERIESNAME")%> &nbsp;</td>
  <td><input  type="hidden" id="Model<%=i%>" name="Model<%=i%>" value="<%=rs("MODELNAME")%>"><%=rs("MODELNAME")%></td>
  <!--
  <td align="center"><input  type="text" id="RETESTYIELD<%=i%>" name="RETESTYIELD<%=i%>" value="<%=rs("RETESTYIELD")%>" size="8"></td>
  <td align="center"><input  type="text" id="RETESTRATIO<%=i%>" name="RETESTRATIO<%=i%>" value="<%=rs("RETESTRATIO")%>" size="8"></td>
  -->
  <td align="center"><input  type="text" id="RemarkY<%=i%>" name="RemarkY<%=i%>" value="<%=rs("Remark1")%>"></td>
  <!--
  <td align="center"><input  type="text" id="SlappingHoldDays<%=i%>" name="SlappingHoldDays<%=i%>" value="<%=rs("SlappingHoldDays")%>" size="10"></td>
  <td align="center"><input  type="text" id="SlappingYield<%=i%>" name="SlappingYield<%=i%>" value="<%=rs("SLAPPINGYIELD")%>" size="8"></td>
  <td align="center"><input  type="text" id="RemarkSlapping<%=i%>" name="RemarkSlapping<%=i%>" value="<%=rs("Remark2")%>"></td>   
   <td align="center"><input  type="text" id="SEA<%=i%>" name="SEA<%=i%>" value="<%=rs("SPECIALEXAMINE1")%>" size="8"></td>
   <td align="center"><input  type="text" id="QA<%=i%>" name="QA<%=i%>" value="<%=rs("AQL1")%>" size="8"></td>
   <td align="center"><input  type="text" id="SEB<%=i%>" name="SEB<%=i%>" value="<%=rs("SPECIALEXAMINE2")%>" size="8"></td>
   <td align="center"><input  type="text" id="QB<%=i%>" name="QB<%=i%>" value="<%=rs("AQL2")%>" size="8"></td>
   <td align="center"><input  type="text" id="SEC<%=i%>" name="SEC<%=i%>" value="<%=rs("SPECIALEXAMINE3")%>" size="8"></td>
   <td align="center"><input  type="text" id="QC<%=i%>" name="QC<%=i%>" value="<%=rs("AQL3")%>" size="8"></td>
   <td align="center"><input  type="text" id="SED<%=i%>" name="SED<%=i%>" value="<%=rs("SPECIALEXAMINE4")%>" size="8"></td>
   <td align="center"><input  type="text" id="QD<%=i%>" name="QD<%=i%>" value="<%=rs("AQL4")%>" size="8"></td>   
  <td align="center"><input  type="text" id="Suffix<%=i%>" name="Suffix<%=i%>" value="<%=rs("Suffix")%>" size="8"></td>
  -->
  <td align="center"><input  type="text" id="CurrentAQL<%=i%>" name="CurrentAQL<%=i%>" value="<%=rs("CurrentAQL")%>" size="8"></td>
  <td align="center"><input  type="text" id="MAXAQL<%=i%>" name="MAXAQL<%=i%>" value="<%=rs("MAXAQL")%>" size="8"></td>
  <td align="center"><input  type="text" id="MINAQL<%=i%>" name="MINAQL<%=i%>" value="<%=rs("MINAQL")%>" size="8"></td>
  <!--
   <td align="center"><input  type="text" id="Owner<%=i%>" name="Owner<%=i%>" value="<%=rs("Owner")%>" size="8"></td>
   
    <td align="center"><input  type="text" id="APPEARANCE<%=i%>" name="APPEARANCE<%=i%>" value="<%=rs("APPEARANCE")%>" size="8"></td>
	 <td align="center"><input  type="text" id="PERFORMANCE<%=i%>" name="PERFORMANCE<%=i%>" value="<%=rs("PERFORMANCE")%>" size="8"></td>
	  <td align="center"><input  type="text" id="MOREOQC<%=i%>" name="MOREOQC<%=i%>" value="<%=rs("MOREOQC")%>" size="8"></td>
   -->


</tr>
<%
	rs.movenext
	wend
%>
  <tr>
  	<td colspan="25" align="center">
	<input type="hidden" name="counts" value="<%=i%>">
	<input  type="hidden" name="seriesName" value="<%=id%>">
	<input  name="btnOK" type="button"  value=" OK " onClick="javascript:document.form1.action='EditSubSeries.asp?action=1';document.form1.submit()">
		&nbsp;
	<input type="reset" name="btnReset" value="Reset">	 
	</td>
  </tr>
</table>
</form>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->