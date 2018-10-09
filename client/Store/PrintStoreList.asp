<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Store/PrintStoreList.asp"
factory=request.QueryString("factory")
section=request.QueryString("section")
ids=""
idcount=0
SQL="select FACTORY_NAME from FACTORY where NID='"&factory&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_name=rs("FACTORY_NAME")
end if
rs.close
'is there any unprinted store records
SQL="select JMS.NID from JOB_MASTER_STORE_PRE JMS left join LINE L on JMS.LINE_NAME=L.LINE_NAME where JMS.PRINT_TIMES=0 and JMS.FACTORY_ID='"&factory&"' and L.SECTION_ID='"&section&"' and rownum<=100 and store_code like '%0000' and exists (select job_number From job_master_scrap_pre where job_number = JMS.job_number and scrap_code like '%0000')"
'response.Write(SQL)
'response.End()
rs.open SQL,conn,1,3
if not rs.eof  then
idcount=rs.recordcount
	while not rs.eof
	ids=ids&rs("NID")&","
	rs.movenext
	wend
	ids=left(ids,len(ids)-1)
end if
rs.close

if idcount>0 then'there are new unprinted store records.
	'if there is no unprint store list, to create a new print sequence, otherwise get last unprinted sequence
	SQL="select * from STORE_PRINT where PRINT_TIME is null and FACTORY_ID='"&factory&"' order by NID desc"
	rs.open SQL,conn,1,3
	if rs.eof then
		rs.addnew
		NID="WP"&NID_SEQ("STOCK")
		rs("NID")=NID
		rs("FACTORY_ID")=factory
		rs.update
	else
		NID=rs("NID")
	end if
	rs.close
	
	'save or update new store records into print history table
	SQL="select * from STORE_PRINT where NID='"&NID&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		rs("PRINT_MEMBERS")=ids
		rs("NOTE")="Auto"
		rs.update
	end if
	rs.close
end if

SQL="select JOB_NUMBER,INSPECT_QUANTITY,PART_NUMBER_TAG,LINE_NAME,NID,MRB,MRB_RFP,STORE_CODE,STORE_TIME,nvl(COMPLETIONSUBINVENTORY,'无') as COMPLETIONSUBINVENTORY  from JOB_MASTER_STORE_PRE JMS left join BAR_PROD_KEB.TBL_MES_LOTMASTER TML on JMS.JOB_NUMBER = TML.WIPENTITYNAME where NID in ('"&replace(ids,",","','")&"') order by COMPLETIONSUBINVENTORY, STORE_TIME"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=factory_name%> 产品入库清单<%=NID%> （<%= date() %>）</title>
<link href="/CSS/Print.css" rel="stylesheet" type="text/css">
<script language="javascript">
function myprint()
{
	window.print();
	return true;
}
function window.onbeforeprint() 
{
	document.all.printspan.style.visibility="hidden";
	return true;
}
function window.onafterprint() 
{
	document.all.printspan.style.visibility="visible";
	return true;
}
function myclose()
{
	document.all.SetPage.src="/Store/PrintStoreListSet.asp?NID=<%=NID%>";
}
function repeatprint()
{
	if(document.all.printid.value!='')
	{
		window.open('/Store/PrintStoreRepeatList.asp?printid='+document.all.printid.value,'');
	}
	else
	{
	alert('打印编号不得为空！');
	}
}
</script>
<style type="text/css">
<!--
.Barcode3of9 {	font-family: "3 of 9 Barcode";
	font-size: 16px;
}
-->
</style>
</head>

<body>
<div id="SetDIV" style="visibility:hidden; position:absolute">
<iframe src="" id="SetPage"></iframe>
</div>
<table width="540" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td>
	<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="30" colspan="13"><%=factory_name%>&nbsp;<%=line%>产品入库清单 （<%= date() %>） </td>
  </tr>
  <tr>
    <td height="30" colspan="13">
      &nbsp;&nbsp;<span class="Barcode3of9"><%= "*"&trim(NID)&"*" %></span>&nbsp;<br>
      &nbsp;&nbsp;<%= NID %></td>
  </tr>
  <tr class="today">
    <td height="20"><div align="center">序列</div></td>
    <td><div align="center">型号</div></td>
    <td><div align="center">工单号</div></td>
    <td><div align="center">线别</div></td>
    <td><div align="center">入库数量 </div></td>
    <td><div align="center">合计数</div></td>
    <td><div align="center">报废数</div></td>
    <td><div align="center">MRB-PEN</div></td>
    <td><div align="center">MRB-RFP</div></td>
    <td><div align="center">入库人</div></td>
    <td><div align="center">入库时间</div></td>
    <td><div align="center">备注</div></td>
	<td><div align="center">入库库位</div></td>
  </tr>
	<%
	i=1
	j=1
	quantity=0
	scrap=0
	while not rs.eof
	%>
  	<tr>
    <td height="24" class="Item"><div align="center">
      <input name="job_number<%=i%>" type="hidden" id="job_number<%=i%>" value="<%=rs("JOB_NUMBER")%>">
      <input name="store_quantity<%=i%>" type="hidden" id="store_quantity<%=i%>" value="<%=rs("inspect_quantity")%>">
      <%= i %>&nbsp;</div></td>
    <td height="24" class="Item"><div align="center"><%= rs("PART_NUMBER_TAG") %></div></td>
    <td height="24" class="Item"><div align="center"><span class="Barcode3of9"><%=rs("JOB_NUMBER")%></span><br><%= rs("JOB_NUMBER") %></div></td>
    <td height="24" class="Item"><div align="center"><%= rs("LINE_NAME") %></div></td>
    <td height="24" class="Item"><div align="center"><%=rs("inspect_quantity") %></div></td>
    <td height="24" class="ItemBold"><div align="center"><span id="gather<%=i%>"></span></div></td>
    <td class="Item"><div align="center">
	<%
	set scraprs=server.createobject("adodb.recordset")
	ScrapSQL="Select nvl(sum(SCRAP_QUANTITY),0) as SCRAP_QUANTITY from JOB_MASTER_SCRAP_PRE where STORE_NID='"&rs("NID")&"'"
	scraprs.open ScrapSQL,conn,1,3
	if not scraprs.eof then
	scrap=scrap+csng(scraprs("SCRAP_QUANTITY"))
	%>
	<%=scraprs("SCRAP_QUANTITY")%>
    <%else%>
    &nbsp;
	<%
	end if
	scraprs.close
	%>
    </div></td>
    <td class="Item"><div align="center">
      <%if rs("MRB")="1" then%>
      √
	  <%else%>
	  &nbsp;
      <%end if%>
    </div></td>
    <td class="Item"><div align="center">
      <div align="center">
        <%if rs("MRB_RFP")="1" then%>
        √
  <%else%>
  &nbsp;
  <%end if%>
      </div>
    </div></td>
    <td height="24" class="Item"><div align="center"><%= rs("STORE_CODE") %></div></td>
    <td height="24" class="Item"><div align="center"><%= rs("STORE_TIME") %></div></td>
    <td width="100" height="24"><div align="center">&nbsp;<span class="Barcode3of9"><%= "*"&rs("NID")&"*" %></span>&nbsp;<br>
        <%= rs("NID") %></div></td>
		<td><%=rs("COMPLETIONSUBINVENTORY")%></td>
  	</tr>
	<%
	quantity=quantity+cint(rs("inspect_quantity"))
	i=i+1
	rs.movenext
	wend
	rs.close%>
	<tr>
    <td><div align="center">总计</div></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><%=quantity%></td>
    <td><%=scrap%></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
    </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" valign="bottom" class="t-c-GrayLight">交货人</td>
    <td height="30" valign="bottom" class="t-c-GrayLight"><div align="left">_____________</div></td>
    <td height="30" valign="bottom" class="t-c-GrayLight">收货人</td>
    <td height="30" valign="bottom" class="t-c-GrayLight"><div align="left">_____________</div></td>
    </tr>
  <tr>
    <td height="30" valign="bottom" class="t-c-GrayLight">打印时间：</td>
    <td height="30" colspan="3" valign="bottom" class="t-c-GrayLight"><%= now() %></td>
    </tr>
    <tr>
      <td height="50" colspan="4" valign="bottom"><div align="center"><span id="printspan" style="visibility:visible">
        <%if i>1 then%>
        先确认，后打印
        <input name="Print" type="button" id="Print" value="打印" onClick="return myprint()" disabled>
        &nbsp;
        <input name="OK" type="button" id="OK" onClick="myclose()" value="确认">
        &nbsp;
        <input name="Close" type="button" disabled id="Close" onClick="javascript:window.close()" value="关闭">
        <%else%>
        <input name="Close" type="button" id="Close" onClick="javascript:window.close()" value="关闭">
        <%end if%>
        &nbsp;
        <input name="printid" type="text" id="printid" size="10">
        <input name="Repeat" type="button" id="Repeat" value="重复打印" onClick="repeatprint()">
        </span></div>      </td>
    </tr>
</table>
	</td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/Store/RowQuantity.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->