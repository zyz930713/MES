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
			if instr(lcase(ajobnumber(w)),"c")=0 and instr(lcase(ajobnumber(w)),"e")=0 and instr(lcase(ajobnumber(w)),"r")=0 then
				if instr(ajobnumber(w),"R")=0 then'is not rework job
					if isnumeric(ajobnumber(w)) then
					sheetnumber=cint(ajobnumber(w))
					else
					response.Redirect("Redirect.asp?factory="&factory&"&errorstring=Job Number is error, please re-scan sub job��<br>�������󣨷����֣���������ɨ���ӹ�����")
					end if
				else
					if isnumeric(right(ajobnumber(w),len(ajobnumber(w))-1)) then
					sheetnumber=cint(right(ajobnumber(w),len(ajobnumber(w))-1))
					else
					response.Redirect("Redirect.asp?factory="&factory&"&errorstring=Job Number is error, please re-scan sub job��<br>��������������ɨ���ӹ�����")
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
final_good_quantity=rs("FINAL_GOOD_QUANTITY")
erp_scrap_quantity=rs("ERP_SCRAP_QUANTITY")
	if rs("STORE_STATUS")="1" then
	store_finished=true
	end if
else
is_new=true
end if
rs.close
SQL="select * from JOB_MASTER_STORE where JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/Store.css" rel="stylesheet" type="text/css">
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
	alert("���ָ�ʽ����");
	location.reload();
	}
	else
	{
		if(ob_changereason.selectedIndex==0)
		{
		alert("�޸����ɲ���Ϊ�գ�");
		}
		else
		{
			if(ob_changecode.value=="")
			{
			alert("�޸��˲���Ϊ�գ�");
			}
			else
			{
				if(confirm("ȷ��Ҫ����������"))
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
	alert("ɾ�����ɲ���Ϊ�գ�");
	}
	else
	{
		if(ob_d_changecode.value=="")
		{
		alert("ɾ���˲���Ϊ�գ�");
		}
		else
		{
			if(confirm("ȷ��Ҫɾ���ü�¼��"))
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
    <td height="20" colspan="9" class="today"><%= jobnumber %><%if is_new=false then%> [<%=part_number_tag%>] &lt;Start Quantity ��<%=start_quantity%>/Stocked Quantity��<%=final_good_quantity%>/Scraped Quantity��<%=erp_scrap_quantity%>&gt;<%end if%><%if store_finished=true then%>
      <span class="bigred">����ѽ���</span>
      <%end if%>&nbsp;</td>
    </tr>
  <tr class="today">
    <td height="20"><div align="center">NO</div></td>
    <td>
      <div align="center">Cancel</div></td>
    <td><div align="center">Type</div></td>
    <td height="20"><div align="center">Operator</div></td>
    <td><div align="center">Assembly Quantity  </div></td>
    <td><div align="center">Stock Quantity  </div></td>
    <td><div align="center">Stock Time </div></td>
    <td><div align="center">On-time Yield </div></td>
    <td><div align="center">Sub Jobs Number </div></td>
    </tr>
	<%i=1
	while not rs.eof%>
  <tr>
    <td><div align="center"><%= i %>&nbsp;</div></td>
    <td><div align="center"><span style="cursor:hand" onClick="javascript:if(document.all.d_changecontrol<%=i%>.style.visibility=='hidden'){document.all.d_changecontrol<%=i%>.style.visibility='visible';document.all.d_changecontrol<%=i%>.style.position='static';}else{document.all.d_changecontrol<%=i%>.style.visibility='hidden';document.all.d_changecontrol<%=i%>.style.position='absolute'}"><img src="/Images/IconDelete.gif" alt="���ͼ��ɾ���ü�¼����������⡣" width="23" height="20"></span>
	<span id="d_changecontrol<%=i%>" style="visibility:hidden;position:absolute">
	<select name="d_changereason<%=i%>" id="d_changereason<%=i%>">
	<option value="">ѡ������</option>
	<option value="�������">�������</option>
	<option value="����δ����ϵͳ">����δ����ϵͳ</option>
	<option value="�����������">�����������</option>
	<option value="�߱���ʾ����">�߱���ʾ����</option>
	<option value="ά�޹������ͺŴ���">ά�޹������ͺŴ���</option>
	</select>
	<input name="d_changecode<%=i%>" type="text" id="d_changecode<%=i%>" value="ɾ����" size="4" onFocus="this.value=''">
	<input name="d_changeOK" type="button" id="d_changeOK" onClick="javascript:if (document.all.d_changecode<%=i%>.value=='ɾ����'){document.all.d_changecode<%=i%>.value='';};deleterecord(document.all.d_changereason<%=i%>,document.all.d_changecode<%=i%>,'<%=rs("NID")%>');" value="ɾ��">
	</span>
	  </div></td>
    <td><div align="center"><%= rs("STORE_TYPE") %></div></td>
    <td><div align="center"><%= rs("STORE_CODE") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("INPUT_QUANTITY") %>&nbsp;</div></td>
    <td><div align="center"><input name="old_quantity<%=i%>" type="hidden" value="<%=rs("STORE_QUANTITY") %>"><input name="quantity<%=i%>" type="text" id="quantity<%=i%>" value="<%=rs("STORE_QUANTITY") %>" size="4" onChange="javascript:if (document.all.old_quantity<%=i%>.value!=this.value){document.all.changecontrol<%=i%>.style.visibility='visible';document.all.changecontrol<%=i%>.style.position='static';}else{document.all.changecontrol<%=i%>.style.visibility='hidden';document.all.changecontrol<%=i%>.style.position='absolute';}">
      <span id="changecontrol<%=i%>" style="visibility:hidden;position:absolute">
	  <select name="changereason<%=i%>" id="changereason<%=i%>">
	  <option value="">ѡ������</option>
	  <option value="�������">�������</option>
	  <option value="����δ����ϵͳ">����δ����ϵͳ</option>
	  <option value="�����������">�����������</option>
	  <option value="�߱���ʾ����">�߱���ʾ����</option>
	  <option value="ά�޹������ͺŴ���">ά�޹������ͺŴ���</option>
      </select>
      <input name="changecode<%=i%>" type="text" id="changecode<%=i%>" value="�޸���" size="4" onFocus="this.value=''">
      <input name="changeOK" type="button" id="changeOK" onClick="if (document.all.changecode<%=i%>.value=='�޸���'){document.all.changecode<%=i%>.value='';};updatequantity(document.all.quantity<%=i%>,document.all.changereason<%=i%>,document.all.changecode<%=i%>,'<%=rs("NID")%>')" value="�޸�">
	  </span>
    </div></td>
    <td><div align="center"><%= rs("STORE_TIME") %></div></td>
    <td><div align="center"><%if rs("YIELD")<>"" then%><%= formatpercent(csng(rs("YIELD")),2,-1) %><%end if%>&nbsp;</div></td>
    <td><%= rs("SUB_JOB_NUMBERS")%>&nbsp;</td>
    </tr>
	<%
	i=i+1
	rs.movenext
	wend
	rs.close%>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->