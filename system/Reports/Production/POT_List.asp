<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>


<!--#include virtual="/WOCF/BOCF_Open.asp" -->





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->

<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>

<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<style type="text/css">
.t-t-DarkBlue div {
	font-size: 18px;
}
</style>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">



<%
material_part_number=request("material_part_number")
material_lot_number=request("material_lot_number")
  
if  material_part_number="430307703141"  or  material_part_number="430307703241"  then

sql="select * from emr_pack_detail where box_id in  ((select batchnoa from emr_pack_detail where box_id='"&material_lot_number&"'))"

'else

'set rsE=server.createobject("adodb.recordset")
'sqlE="select * from emr_pack_detail where box_id='"&LABELIDN&"'"
'sql=" select * from emr_pack_detail where box_id  in('"&replace(material_lot_number,"/","','")&"')"
'end if
rs.open sql,conn,1,3

'while not rs.eof  
%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    	
    <tr>
      <td height="20" colspan="6" class="t-c-greenCopy"  align="center">(Outer Pot Assy)</td>
    </tr>
	
    <tr>
      <td width="147" rowspan="2" >Operator Code </td>
      <td width="313" rowspan="2" ><%=rs("PACK_USER")%></td>
      <td width="117" height="30" >Inspector</td>
      <td colspan="3" ><%=rs("INSPECTOR")%>&nbsp;</td>
    </tr>
    <tr>
      <td height="25" >Description</td>
      <td width="128" >Supplier </td>
      <td width="178" >Batch No</td>
      <td width="258" >Date NO</td>
    </tr>
    <tr>
      <td >Production</td>
      <td ><%=rs("PRODUCTION_NAME")%>&nbsp;</td>
      <td >POT</td>
      <td ><%=rs("SUPPLIERA")%>&nbsp;</td>
      <td ><%=rs("BATCHNOA")%>&nbsp;</td>
      <td ><%=rs("PRODTIMEA")%>&nbsp;</td>
    </tr>
    <tr>
      <td >12NC</td>
      <td ><%=rs("PART_NUMBER")%>&nbsp;</td>
      <td >TOP-Ring</td>
      <td ><%=rs("SUPPLIERB")%>&nbsp;</td>
      <td ><%=rs("BATCHNOB")%>&nbsp;</td>
      <td ><%=rs("PRODTIMEB")%>&nbsp;</td>
    </tr>
    <tr>
      <td >Shift</td>
      <td ><%=rs("SHIFT")%>&nbsp;</td>
      <td >Magent-L</td>
      <td ><%=rs("SUPPLIERC")%>&nbsp;</td>
      <td ><%=rs("BATCHNOC")%>&nbsp;</td>
      <td ><%=rs("PRODTIMEC")%>&nbsp;</td>
    </tr>
     <tr>
      <td >Date</td>
      <td ><%=rs("PACK_TIME")%>&nbsp;</td>
      <td >Magent-S</td>
      <td ><%=rs("SUPPLIERD")%>&nbsp;</td>
      <td ><%=rs("BATCHNOD")%>&nbsp;</td>
      <td ><%=rs("PRODTIMED")%>&nbsp;</td>
    </tr>
    <tr>
      <td >Qty</td>
      <td ><%=rs("PACK_QTY")%></td>
      <td >Equipment NO</td>
      <td colspan="3" ><%=rs("EQUIPMENT_NO")%>&nbsp;</td>
    </tr>
	<tr>
      <td >Remark</td>
      <td ><%=rs("REMARK")%>&nbsp;</td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    	<tr>
      <td >JOB Number</td>
      <td ><%=rs("JOB_NUMBER")%></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" align="center">&nbsp;</td>
    </tr>
    
   
 
</table>
<%

rs.close

sql="select * from emr_pack_detail where box_id in  ((select batchnob from emr_pack_detail where box_id='"&material_lot_number&"'))"
rs.open sql,conn,1,3
%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
 
    <tr>
      <td height="20" colspan="6" class="t-c-greenCopy"  align="center">(inner  Pot Assy )</td>
    </tr>
	
    <tr>
      <td width="147" rowspan="2" >Operator Code </td>
      <td width="313" rowspan="2" ><%=rs("PACK_USER")%></td>
      <td width="117" height="26" >Inspector</td>
      <td colspan="3" ><%=rs("INSPECTOR")%>&nbsp;</td>
    </tr>
    <tr>
      <td height="25" >Description</td>
      <td width="128" >Supplier </td>
      <td width="178" >Batch No</td>
      <td width="258" >Date NO</td>
    </tr>
    <tr>
      <td >Production</td>
      <td ><%=rs("PRODUCTION_NAME")%>&nbsp;</td>
      <td >Top-Plate</td>
      <td ><%=rs("SUPPLIERA")%>&nbsp;</td>
      <td ><%=rs("BATCHNOA")%>&nbsp;</td>
      <td >&nbsp;</td>
    </tr>
    <tr>
      <td >12NC</td>
      <td ><%=rs("PART_NUMBER")%>&nbsp;</td>
      <td >Magent</td>
      <td ><%=rs("SUPPLIERB")%>&nbsp;</td>
      <td ><%=rs("BATCHNOB")%>&nbsp;</td>
      <td >&nbsp;</td>
    </tr>
    <tr>
      <td >Shift</td>
      <td ><%=rs("SHIFT")%>&nbsp;</td>
      <td >Equipment NO</td>
      <td ><%=rs("EQUIPMENT_NO")%>&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
    </tr>
     <tr>
      <td >Date</td>
      <td ><%=rs("PACK_TIME")%>&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
    </tr>
    <tr>
      <td >Qty</td>
      <td ><%=rs("PACK_QTY")%></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
	<tr>
      <td >Remark</td>
      <td ><%=rs("REMARK")%>&nbsp;</td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    	<tr>
      <td >JOB Number</td>
      <td ><%=rs("JOB_NUMBER")%></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" align="center">&nbsp;</td>
    </tr>
    
   
  
</table>

<%
rs.close
else
sql=" select * from emr_pack_detail where box_id  in('"&replace(material_lot_number,"/","','")&"')"
rs.open sql,conn,1,3


%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    	
    <tr>
      <td height="20" colspan="6" class="t-c-greenCopy"  align="center">(Outer Pot Assy )</td>
    </tr>
	
    <tr>
      <td width="147" rowspan="2" >Operator Code </td>
      <td width="313" rowspan="2" ><%=rs("PACK_USER")%></td>
      <td width="117" height="30" >Inspector</td>
      <td colspan="3" ><%=rs("INSPECTOR")%>&nbsp;</td>
    </tr>
    <tr>
      <td height="25" >Description</td>
      <td width="128" >Supplier </td>
      <td width="178" >Batch No</td>
      <td width="258" >Date NO</td>
    </tr>
    <tr>
      <td >Production</td>
      <td ><%=rs("PRODUCTION_NAME")%>&nbsp;</td>
      <td >POT</td>
      <td ><%=rs("SUPPLIERA")%>&nbsp;</td>
      <td ><%=rs("BATCHNOA")%>&nbsp;</td>
      <td ><%=rs("PRODTIMEA")%>&nbsp;</td>
    </tr>
    <tr>
      <td >12NC</td>
      <td ><%=rs("PART_NUMBER")%>&nbsp;</td>
      <td >TOP-Ring</td>
      <td ><%=rs("SUPPLIERB")%>&nbsp;</td>
      <td ><%=rs("BATCHNOB")%>&nbsp;</td>
      <td ><%=rs("PRODTIMEB")%>&nbsp;</td>
    </tr>
    <tr>
      <td >Shift</td>
      <td ><%=rs("SHIFT")%>&nbsp;</td>
      <td >Magent-L</td>
      <td ><%=rs("SUPPLIERC")%>&nbsp;</td>
      <td ><%=rs("BATCHNOC")%>&nbsp;</td>
      <td ><%=rs("PRODTIMEC")%>&nbsp;</td>
    </tr>
     <tr>
      <td >Date</td>
      <td ><%=rs("PACK_TIME")%>&nbsp;</td>
      <td >Magent-S</td>
      <td ><%=rs("SUPPLIERD")%>&nbsp;</td>
      <td ><%=rs("BATCHNOD")%>&nbsp;</td>
      <td ><%=rs("PRODTIMED")%>&nbsp;</td>
    </tr>
    <tr>
      <td >Qty</td>
      <td ><%=rs("PACK_QTY")%></td>
      <td >Equipment NO</td>
      <td colspan="3" ><%=rs("EQUIPMENT_NO")%>&nbsp;</td>
    </tr>
	<tr>
      <td >Remark</td>
      <td ><%=rs("REMARK")%>&nbsp;</td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    	<tr>
      <td >JOB Number</td>
      <td ><%=rs("JOB_NUMBER")%></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" align="center">&nbsp;</td>
    </tr>
    
   
 
</table>

<%
rs.close
end if%>



<table width="1154" height="26">
<tr><td align="center"><input name="btnClose2" type="button"  id="btnClose2"  value="Back" onClick="javascript:history.go(-1);"></td>
</tr></table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->

