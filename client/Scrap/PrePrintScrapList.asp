<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request("factory")
line=request("line")
printid=request("printid")
factory_name=request("factory_name")
ids=""
idcount=-1

'reprint
if printid <>"" then
	SQL="SELECT PRINT_MEMBERS,B.FACTORY_NAME FROM SCRAP_PRINT A,FACTORY B WHERE A.FACTORY_ID = B.NID AND A.NID='"&printid&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		ids=rs("PRINT_MEMBERS")
		factory_name=rs("FACTORY_NAME")
		NID=printid
	else
		word="This print id does not exist. 该打印编号不存在。"
	end if
	rs.close
end if	
if word="" and (line <> "" or printid <>"") then
	SQL="SELECT JS.NID,JS.PART_NUMBER_TAG,JS.JOB_NUMBER,JS.LINE_NAME,SUM(JS.SCRAP_QUANTITY) AS SCRAP_QUANTITY,MAX(JS.SCRAP_CODE) AS SCRAP_CODE ,MAX(JS.SCRAP_TIME) AS SCRAP_TIME "
	SQL=SQL+" FROM JOB_MASTER_SCRAP_PRE JS  "
	if ids <> "" then
		SQL=SQL+" WHERE NID IN ('"&replace(ids,",","','")&"')"
	else
		SQL=SQL+" INNER JOIN LINE L ON JS.LINE_NAME=L.LINE_NAME WHERE JS.PRINT_TIMES=0 AND JS.SCRAP_QUANTITY<>0 AND JS.FACTORY_ID='"&factory&"' AND L.NID='"&line&"'"
	end if
	SQL=SQL+" GROUP BY JS.NID,JS.PART_NUMBER_TAG,JS.JOB_NUMBER,JS.LINE_NAME"
	rs.open SQL,conn,1,3
	idcount=rs.recordcount
	if not rs.eof  then
		line_name=rs("LINE_NAME")		
		while not rs.eof
			ids=ids&rs("NID")&","
			rs.movenext
		wend
		ids=left(ids,len(ids)-1)
	rs.movefirst	
	end if
	if idcount>0 and printid="" then'there are new unprinted scrap records.
		set rsPrint = server.createobject("adodb.recordset")
		'if there is no unprint scrap list, to create a new print sequence, otherwise get last unprinted sequence
		SQL="SELECT NID,PRINT_TIME,PRINT_MEMBERS,FACTORY_ID,LINE_ID FROM SCRAP_PRINT WHERE PRINT_TIME IS NULL ORDER BY NID DESC"
		rsPrint.open SQL,conn,1,3
		if rsPrint.eof then
			rsPrint.addnew
			NID="SP"&NID_SEQ("KE_SCRAP")
			rsPrint("NID")=NID
			rsPrint("FACTORY_ID")=factory
			rsPrint("LINE_ID")=line
		else
			NID=rsPrint("NID")
		end if
		rsPrint("PRINT_MEMBERS")=ids
		rsPrint.update
		rsPrint.close
		set rsPrint=nothing			
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">

<script language="javascript">
<%if word<>"" then%>
	alert("<%=word%>")
<%end if%>
function myprint()
{
	window.print();
	<%if printid = "" then%>
	window.showModalDialog("UpdatePrintResult.asp?NID=<%=NID%>",window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
	<%end if%>
	return true;
}
function window.onbeforeprint() 
{
	document.all.printspan.style.visibility="hidden";
	document.all.SetDIV.style.visibility="hidden";
	return true;
}
function window.onafterprint() 
{
	document.all.printspan.style.visibility="visible";
	document.all.SetDIV.style.visibility="visible";
	return true;
}

function repeatprint()
{
	if(document.all.printid.value!='')
	{
		window.open('/Scrap/PrintScrapRepeatList.asp?printid='+document.all.printid.value,'');
	}
	else
	{
		alert('打印编号不得为空！');
	}
}
function getFactoryName(){
	var selIndex=document.getElementById("factory").selectedIndex;
	document.getElementById("factory_name").value=document.getElementById("factory").options[selIndex].text;
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

<body bgcolor="#339966" onLoad="getFactoryName();">
<div id="SetDIV">
<table border="1" width="800" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<form id="form1" method="post" name="form1">
	<tr>
		<Td height="20" colspan="8" class="t-t-DarkBlue"  align="center">Print Scrap List 打印报废清单 （<%= date() %>）</Td>
	</tr>	
	<tr><td>
		<table border="1" width="100%" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
		<tr>
		<td align="center">Factory 工厂</td>
		<td ><select id="factory"  name="factory" onChange="getFactoryName();" >
			<%=getFactory("OPTION",factory)%>			
			</select>
		</td>
		<Td align="center">Line Name 线别</Td>
		<Td ><select id="line"  name="line" >
			<%=getLine("OPTION",line,"","","")%>					
			</select>
		</Td>
		<Td align="center">Print Id 打印编号</Td>
		<Td ><input type="text" id="printid" name="printid" value="<%=printid%>"> 
		</Td>
		<Td align="center">	
			<input type="hidden" name="factory_name" id="factory_name">
			<input name="btnQuery" type="button"  id="btnQuery" onClick="form1.submit();" value="Query 查询">
			&nbsp;
			<input name="btnSearchHis" type="button"  id="btnSearchHis" onClick="location.href='SearchZeroPrintList.asp'" value="Print HIS 打印历史">
		</Td>
		</tr></table></td>
	</tr>	
<%if idcount =0 then%>
	<tr><td align="center">No Records 没有记录</td></tr>	
<%end if%>
	</form>
</table>
</div>
<br>
<%if idcount >0 then%>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td>
	<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="30" colspan="9"><%=factory_name%>&nbsp;<%=line_name%>&nbsp;产品报废清单 （<%= date() %>） </td>
  </tr>
  <tr>
    <td height="30" colspan="9">
      &nbsp;&nbsp;<span class="Barcode3of9"><%= "*"&trim(NID)&"*" %></span>&nbsp;<br>
      &nbsp;&nbsp;<%= NID %></td>
  </tr>
  <tr class="today">
    <td height="20"><div align="center">序列</div></td>
    <td><div align="center">型号</div></td>
    <td><div align="center">工单号</div></td>
    <td><div align="center">线别</div></td>
    <td><div align="center">报废数量 </div></td>
    <td><div align="center">合计数</div></td>
    <td><div align="center">报废人</div></td>
    <td><div align="center">报废时间</div></td>
    <td><div align="center">备注</div></td>
  </tr>
	<%
	i=1
	j=1
	quantity=0
	while not rs.eof
	%>
  	<tr>
	<td height="24" class="Item"><div align="center">
      <input name="job_number<%=i%>" type="hidden" id="job_number<%=i%>" value="<%=rs("JOB_NUMBER")%>">
      <input name="scrap_quantity<%=i%>" type="hidden" id="scrap_quantity<%=i%>" value="<%=rs("SCRAP_QUANTITY")%>">
      <%= i %>&nbsp;</div></td>
    <td height="24" class="Item"><div align="center"><%= rs("PART_NUMBER_TAG") %></div></td>
    <td height="24" class="Item"><div align="center"><%= rs("JOB_NUMBER") %></div></td>
    <td height="24" class="Item"><div align="center"><%= rs("LINE_NAME") %></div></td>
    <td height="24" class="Item"><div align="center"><%=rs("SCRAP_QUANTITY") %></div></td>
    <td height="24" class="ItemBold"><div align="center"><span id="gather<%=i%>"></span></div></td>
    <td height="24" class="Item"><div align="center"><%= rs("SCRAP_CODE") %></div></td>
    <td height="24" class="Item"><div align="center"><%= rs("SCRAP_TIME") %></div></td>
    <td width="100" height="24"><div align="center">&nbsp;<span class="Barcode3of9"><%= "*"&rs("NID")&"*" %></span>&nbsp;<br>
        <%= rs("NID") %></div></td>
  	</tr>
	<%
	quantity=quantity+cint(rs("SCRAP_QUANTITY"))
	i=i+1
	rs.movenext
	wend
	rs.close%>
	<tr>
    <td><div align="center">总计</div></td>
    <td colspan="8"><%=quantity%>&nbsp;</td>
    </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" valign="bottom" class="t-c-GrayLight" >&nbsp;交货人：&nbsp;_____________</td>
    <td height="30" valign="bottom" class="t-c-GrayLight" >&nbsp;收货人：&nbsp; _____________</td>
    </tr>
  <tr>
    <td height="30" colspan="2" valign="bottom" class="t-c-GrayLight">&nbsp;打印时间：<%= now() %></td>    
    </tr>
    <tr>
      <td height="50" colspan="4" valign="bottom"><div align="center"><span id="printspan" style="visibility:visible">
        <%if quantity>0 then%>
        <input name="Print" type="button" id="Print" value="打印" onClick="return myprint()">                
        <%end if%> 
        &nbsp;<input name="Close" type="button" id="Close" onClick="javascript:window.close()" value="关闭">              
        </span></div>
		</td>
    </tr>
</table>
	</td>
  </tr>
</table>
<!--#include virtual="/Scrap/RowQuantity.asp" -->
<%end if%>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->