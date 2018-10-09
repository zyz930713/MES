<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="JobStock.asp"
'response.Write(trim(request.QueryString("mainjobnumber")))
'response.End()
if instr(trim(request.QueryString("mainjobnumber")),"-")=0 then
jobnumber=trim(request.QueryString("mainjobnumber"))
else
	ajobnumber=split(ucase(trim(request.QueryString("mainjobnumber"))),"-")
	for w=0 to ubound(ajobnumber)
		if w<>ubound(ajobnumber) then
		jobnumber=jobnumber&ajobnumber(w)&"-"
		else'get sheet number
			if instr(lcase(ajobnumber(w)),"c")=0 and instr(lcase(ajobnumber(w)),"e")=0 and instr(lcase(ajobnumber(w)),"o")=0 and instr(lcase(ajobnumber(w)),"r")=0 then
				if instr(ajobnumber(w),"R")=0 then'is not rework job
					if isnumeric(ajobnumber(w)) then
					sheetnumber=cint(ajobnumber(w))
					else
					response.Redirect("Redirect.asp?factory="&factory&"&errorstring=Job Number is error, please re-scan sub job！<br>工单错误（非数字），请重新扫描子工单！")
					end if
				else
					if isnumeric(right(ajobnumber(w),len(ajobnumber(w))-1)) then
					sheetnumber=cint(right(ajobnumber(w),len(ajobnumber(w))-1))
					else
					response.Redirect("Redirect.asp?factory="&factory&"&errorstring=Job Number is error, please re-scan sub job！<br>工单错误，请重新扫描子工单！")
					end if
				end if
			else
			jobnumber=jobnumber&ajobnumber(w)&"-"
			end if
		end if
	next
	jobnumber=left(jobnumber,len(jobnumber)-1)
end if


is_new=true
store_finished=false
SQL="select * from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
if not rs.eof then
is_new=false
part_number_tag=rs("PART_NUMBER_TAG")
start_quantity=rs("START_QUANTITY")
SQL="SELECT * FROM label_print_history WHERE JOB_NUMBER='"&jobnumber&"'"
set rsPrintHistory=server.createobject("adodb.recordset")
rsPrintHistory.open SQL,conn,1,3
if(rsPrintHistory.recordcount>0) then
	final_good_quantity=CLng(rs("CONFIRM_GOOD_QUANTITY"))
else
	final_good_quantity=CLng(rs("FINAL_GOOD_QUANTITY"))
end if 
erp_scrap_quantity=rs("FINAL_SCRAP_QUANTITY")
	if rs("STORE_STATUS")="1" then
		store_finished=true
	end if
else
is_new=true
end if
rs.close
SQL="select * from JOB_MASTER_STORE_PRE where JOB_NUMBER='"&jobnumber&"' order by store_time desc"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<%if store_finished=true then%>
<script language="javascript">
parent.document.form1.Next.disabled=true;
</script>
<%end if%>
<script language="javascript">
function updatequantity(ob_changequantity,ob_changereason,ob_changecode,nid)
{
	if(isNumberString(ob_changequantity.value,'1234567890')!=1)
	{
	alert("数字格式错误！");
	location.reload();
	}
	else
	{
		if(ob_changereason.selectedIndex==0)
		{
		alert("修改理由不得为空！");
		}
		else
		{
			if(ob_changecode.value=="")
			{
			alert("修改人不得为空！");
			}
			else
			{
				if(confirm("确定要更新数量吗？"))
				{
				location.href="UpdateQuantity.asp?nid="+nid+"&thisvalue="+ob_changequantity.value+"&changereason="+ob_changereason.options[ob_changereason.selectedIndex].value+"&changecode="+ob_changecode.value+"&path=<%=path%>&query=<%=query%>"
				}
				else
				{
				location.reload();
				}
			}
		}
	}
}
function deleterecord(ob_d_changereason,ob_d_changecode,nid)
{
	if(ob_d_changereason.selectedIndex==0)
	{
		alert("Delete reason cannot be blank! 删除理由不得为空!");
	}
	else
	{
		if(ob_d_changecode.value=="")
		{
		alert("Op Code cannot be blank! 删除人不得为空!");
		}
		else
		{
			if(confirm("Are you sure to delete this record? 确定要删除该记录吗？"))
			{
			location.href="DeleteStore.asp?nid="+nid+"&changereason="+ob_d_changereason.options[ob_d_changereason.selectedIndex].value+"&changecode="+ob_d_changecode.value+"&path=<%=path%>&query=<%=query%>"
			}

		}
	}
}
</script>
</head>

<body>
<form action="" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="10" class="today" align="right"><%= jobnumber %><%if is_new=false then%> [<%=part_number_tag%>] Start Qty 开始数量[<%=start_quantity%>] Store Qty 入库数量[<%=final_good_quantity%>]  Scrap Qty 报废数量[<%=erp_scrap_quantity%>] <%end if%><%if store_finished=true then%>
      <span class="bigred">Store Finished 入库已结束</span>
      <%end if%>&nbsp;</td>
    </tr>
  <tr class="today">
    <td height="20"><div align="center">Seq 序列</div></td>
    <td>
      <div align="center">Cancel 取消</div></td>
    <td><div align="center">Type 类型</div></td>
    <td height="20"><div align="center">Store Code 入库人</div></td>
    <td><div align="center">Input Qty 生产数量 </div></td>
    <td><div align="center">Inspect Qty 复检数量</div></td>
    <td><div align="center">Store Qty 入库数量 </div></td>
    <td><div align="center">Store Time 入库时间</div></td>
    <td><div align="center">Yeild 即时良率</div></td>
    <td><div align="center">Sub Job Number 子工单号</div></td>
    </tr>
	<%i=1
	totalinput=0
	totalinspect=0
	totalstore=0
	while not rs.eof%>
  <tr>
    <td><div align="center"><%= i %>&nbsp;</div></td>
    <td><div align="center">
		<span style="cursor:hand" onClick="javascript:if(document.all.d_changecontrol<%=i%>.style.visibility=='hidden'){document.all.d_changecontrol<%=i%>.style.visibility='visible';document.all.d_changecontrol<%=i%>.style.position='static';}else{document.all.d_changecontrol<%=i%>.style.visibility='hidden';document.all.d_changecontrol<%=i%>.style.position='absolute'}">
			<%if rs("PRINT_TIMES")<>"1" then%><img src="/Images/IconDelete.gif" alt="" width="23" height="20"><%else%>&nbsp;<%end if%> 
		</span>
	<span id="d_changecontrol<%=i%>" style="visibility:hidden;position:absolute">
	<select name="d_changereason<%=i%>" id="d_changereason<%=i%>">
	<option value=""></option>
	<option value="输入错误">输入错误</option>
	<option value="产线未输入系统">产线未输入系统</option>
	<option value="产线数量溢出">产线数量溢出</option>
	<option value="线别显示错误">线别显示错误</option>
	</select>
	<input name="d_changecode<%=i%>" type="text" id="d_changecode<%=i%>" value="Op工号" size="4" onFocus="this.value=''">
	<input name="d_changeOK" type="button" id="d_changeOK" onClick="javascript:if (document.all.d_changecode<%=i%>.value=='Op工号'){document.all.d_changecode<%=i%>.value='';};deleterecord(document.all.d_changereason<%=i%>,document.all.d_changecode<%=i%>,'<%=rs("NID")%>');" value="OK 确定">
	</span>
	  </div></td>
    <td><div align="center"><%= rs("STORE_TYPE") %></div></td>
    <td><div align="center"><%= rs("STORE_CODE") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("INPUT_QUANTITY") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("INSPECT_QUANTITY") %></div></td>
    <td><div align="center"><input name="old_quantity<%=i%>" type="hidden" value="<%=rs("STORE_QUANTITY") %>"><%=rs("STORE_QUANTITY") %><!--<input name="quantity<%'=i%>" readonly=true type="text" id="quantity<%'=i%>" value="<%'=rs("STORE_QUANTITY") %>" size="4" onChange="javascript:if (document.all.old_quantity<%'=i%>.value!=this.value){document.all.changecontrol<%'=i%>.style.visibility='visible';document.all.changecontrol<%'=i%>.style.position='static';}else{document.all.changecontrol<%'=i%>.style.visibility='hidden';document.all.changecontrol<%'=i%>.style.position='absolute';}">-->
      <span id="changecontrol<%=i%>" style="visibility:hidden;position:absolute">
	  <select name="changereason<%=i%>" id="changereason<%=i%>">
	  <option value="">选择理由</option>
	  <option value="输入错误">输入错误</option>
	  <option value="产线未输入系统">产线未输入系统</option>
	  <option value="产线数量溢出">产线数量溢出</option>
	  <option value="线别显示错误">线别显示错误</option>
	  <option value="维修工单的型号错误">维修工单的型号错误</option>
      </select>
      <input name="changecode<%=i%>" type="text" id="changecode<%=i%>" value="Op工号" size="4" onFocus="this.value=''">
      <input name="changeOK" type="button" id="changeOK" onClick="if (document.all.changecode<%=i%>.value=='Op工号'){document.all.changecode<%=i%>.value='';};updatequantity(document.all.quantity<%=i%>,document.all.changereason<%=i%>,document.all.changecode<%=i%>,'<%=rs("NID")%>')" value="OK 确定">
	  </span>
    </div></td>
    <td><div align="center"><%= rs("STORE_TIME") %></div></td>
    <td><div align="center"><%if rs("YIELD")<>"" then%><%= formatpercent(csng(rs("YIELD")),2,-1) %><%end if%>&nbsp;</div></td>
    <td><%= rs("SUB_JOB_NUMBERS")%>&nbsp;</td>
    </tr>	<%
	totalinput=totalinput+csng(rs("INPUT_QUANTITY"))
	totalinspect=totalinspect+csng(rs("INSPECT_QUANTITY"))
	totalstore=totalstore+csng(rs("STORE_QUANTITY"))
	i=i+1
	rs.movenext
	wend
	rs.close
	SQL="SELECT START_QUANTITY,CONFIRM_GOOD_QUANTITY,FINAL_SCRAP_QUANTITY from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		master_start_quantity=csng(rs("START_QUANTITY"))
		master_store_quantity=csng(rs("CONFIRM_GOOD_QUANTITY"))
		master_scrap_quantity=csng(rs("FINAL_SCRAP_QUANTITY"))
	end if
	rs.close
	if master_store_quantity<>totalinspect then
		SQL="update JOB_MASTER set CONFIRM_GOOD_QUANTITY="&totalinspect&" where JOB_NUMBER='"&jobnumber&"'"
		rs.open SQL,conn,1,3
	end if
	if master_store_quantity+master_scrap_quantity<>master_start_quantity then
		SQL="update JOB_MASTER set STORE_STATUS=0 where JOB_NUMBER='"&jobnumber&"'"
		rs.open SQL,conn,1,3
	else
		SQL="update JOB_MASTER set STORE_STATUS=1 where JOB_NUMBER='"&jobnumber&"'"
		rs.open SQL,conn,1,3
	end if
	%>
  <tr>
    <td>Total 合计</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><%= totalinput %></td>
    <td><%= totalinspect %></td>
    <td><%= totalstore %></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>

</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->