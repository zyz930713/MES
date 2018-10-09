<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
server.ScriptTimeout=99999
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
scantype=trim(request.Form("scantype"))
code=trim(request.Form("code"))
TicketID=ucase(trim(request.Form("TicketID")))
ID=ucase(trim(request.Form("ID")))
set rss=server.CreateObject("adodb.recordset")

if scantype="0" then
	total_store=0
	SQL="select * from STORE_PRINT where NID='"&TicketID&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	ids=rs("PRINT_MEMBERS")
	array_ids=split(ids,",")
		for i=0 to ubound(array_ids)
			errormsg=errormsg&eachstore(array_ids(i),code,store_quantity,scrap_quantity)
			total_store=total_store+csng(store_quantity)
			total_scrap=total_scrap+csng(scrap_quantity)
			new_ids=new_ids&array_ids(i)&"<br>"
		next
	end if
	rs.close
else
	errormsg=eachstore(ID,code,store_quantity,scrap_quantity)
end if
'check whether ID is comfirmed
'check have confirmed or not
function eachstore(thisID,thisCode,store_quantity,scrap_quantity)
store_quantity=0
	SQL="select * from JOB_MASTER_STORE_PRE where NID='"&thisID&"'"
	rss.open SQL,conn,1,3
	if rss.eof then
	eachstore=thisID&"没有入过库，数据错误！<br>"
	end if
	rss.close
	if eachstore<>"" then
		exit function
	end if

	SQL="select * from JOB_MASTER_STORE where NID='"&thisID&"'"
	rss.open SQL,conn,1,3
	if not rss.eof then
		eachstore=thisID&"已经由"&rss("STORE_CODE")&"在"&rss("STORE_TIME")&"做过入库确认！请重新选择批号。<br>"
	end if
	rss.close
	if eachstore<>"" then
		exit function
	end if
	'get store quantity in job_master
	SQL="select * from JOB_MASTER_STORE_PRE where NID='"&thisID&"'"
	rss.open SQL,conn,1,3
	if not rss.eof then
		jobnumber=rss("JOB_NUMBER")
		store_quantity=csng(rss("inspect_quantity"))
	end if
	rss.close
	'save store info in job_master_store产品入库，JOB_MASTER_STORE_PRE中真正接受的数量是inspect_quantity，STORE_QUANITY是理论应该入库数量。
	SQL="insert into JOB_MASTER_STORE select JOB_NUMBER,'"&thisCode&"',sysdate,INPUT_QUANTITY,INSPECT_QUANTITY,YIELD,STORE_TYPE,SUB_JOB_NUMBERS,PART_NUMBER_TAG,FACTORY_ID,LINE_NAME,NID,RETEST_CHECK_STATUS,RETEST_CHECK_CODE,RETEST_CHECK_TIME,PRINT_TIMES,NOTE,NO_YIELD,MRB,INSPECT_QUANTITY,MRB_RFP,TRANSACTION_ID,MESSAGE_STATUS,ERP_JOB_CLOSE_STATUS,ERP_ERROR_CODE,ERP_ERROR_EXPLANATION,ALLOW_ERP_RESUBMIT,ERP_CREATE_TIME,ERP_CREATE_BY,ERP_LAST_UPDATE_TIME,ERP_LAST_UPDATE_BY,INSTORE_GUID_ID from JOB_MASTER_STORE_PRE where NID='"&thisID&"'"
	rss.open SQL,conn,1,3
	
	'save store info in job_master
	'SQL="select * from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
'	rss.open SQL,conn,1,3
'	if not rss.eof then
'	rss("CONFIRM_GOOD_QUANTITY")=csng(rss("CONFIRM_GOOD_QUANTITY"))+store_quantity
'	rss.update
'	end if
'	rss.close
	
scrap_quantity=0
	SQL="select * from JOB_MASTER_SCRAP_PRE where STORE_NID='"&thisID&"'"
	rss.open SQL,conn,1,3
	if rss.eof then
	eachscrap=thisID&"没有入报废，数据错误！<br>"
	end if
	rss.close
	if eachscrap<>"" then
		exit function
	end if

	SQL="select * from JOB_MASTER_SCRAP where STORE_NID='"&thisID&"'"
	rss.open SQL,conn,1,3
	if not rss.eof then
		eachscrap=thisID&"已经由"&rss("SCRAP_CODE")&"在"&rss("SCRAP_TIME")&"做过报废确认！请重新选择批号。<br>"
	end if
	rss.close
	if eachscrap<>"" then
		exit function
	end if


	'get store quantity in job_master
	SQL="select sum(SCRAP_QUANTITY) as SCRAP_QUANTITY from JOB_MASTER_SCRAP_PRE where STORE_NID='"&thisID&"'"
	rss.open SQL,conn,1,3
	if not rss.eof then
		scrap_quantity=csng(rss("SCRAP_QUANTITY"))
	end if
	rss.close
	
	'save scrap info in job_master_scrap产品入库。
	if scrap_quantity<>0 then
		SQL="insert into JOB_MASTER_SCRAP select JOB_NUMBER,'"&thisCode&"',sysdate,SCRAP_QUANTITY,PART_NUMBER_TAG,FACTORY_ID,LINE_NAME,NID,PROFILE_FORM_ID,PRINT_TIMES,NOTE,REASON,TRANSACTION_ID,SCRAP_ACCOUNT,SCRAP_REASON_ID,MESSAGE_STATUS,ERP_JOB_CLOSE_STATUS,ERP_ERROR_CODE,ERP_ERROR_EXPLANATION,ALLOW_ERP_RESUBMIT,ERP_CREATE_TIME,ERP_CREATE_BY,ERP_LAST_UPDATE_TIME,ERP_LAST_UPDATE_BY,STORE_NID,SCRAP_REFERENCE,INSTORE_GUID_ID from JOB_MASTER_SCRAP_PRE where STORE_NID='"&thisID&"'"
	rss.open SQL,conn,1,3
	end if
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/Store.css" rel="stylesheet" type="text/css">
<script language="javascript">
timePopup=5;
adCount=0;
function showPopup()
{
adCount+=1
	if(adCount<timePopup)
	{
	setTimeout("showPopup()",1000);
	document.all.countinsert.innerText="("+(timePopup-adCount)+")";
	}
  else
	{
	closePopup()
	}
}
function closePopup()
{
location.href="Confirm1.asp"
}
</script>
</head>

<body <%if errormsg="" then%>onLoad="showPopup();"<%end if%>>
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2"><div align="center" class="strongred">System will close window in 10 seconds.&nbsp;<br>
      系统将在10秒钟后关闭窗口。 <span id="countinsert"></span></div></td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-c-green">Confirmation of storing 入库确认</td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-t-Borrow">Operator 操作员:<% =code %></td>
  </tr>
  <tr>
    <td height="20">Ticket Number 入库单编号</td>
    <td height="20"><% =TicketID %>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="37%" height="20">ID Number 单次入库序号</td>
    <td width="63%" height="20"><%if scantype="0" then%><%=new_ids%><%else%><% =ID %><%end if%>&nbsp;</td>
  </tr>
  <tr>
    <td height="20">Confirmed Store Quantity 确认入库数量</td>
    <td height="20"><%if scantype="0" then%>
      <%=total_store%>
      <%else%>
      <% =store_quantity %>
    <%end if%>&nbsp;</td>
  </tr>
  <tr>
    <td height="20">Confirmed Scrap Quantity 确认报废数量</td>
    <td height="20"><%if scantype="0" then%>
      <%=total_scrap%>
      <%else%>
      <% =scrap_quantity %>
      <%end if%>
      &nbsp;</td>
  </tr>
  <tr>
    <td height="20">Confirmed Person 确认人</td>
    <td height="20"><% =code %>&nbsp;</td>
  </tr>
  <tr>
    <td height="20">错误信息</td>
    <td height="20"><%=errormsg%>&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center"><input type="button" name="button" id="button" value="返回" onClick="history.back()"></div></td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
