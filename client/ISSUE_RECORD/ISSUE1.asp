<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
code=trim(request.Form("LM_USER"))
Box_id=trim(request("Box_id"))


      
        set rsGetBoxid=server.CreateObject("adodb.recordset")
		SQL="SELECT * FROM job_pack_DETAIL WHERE Box_id='"+Box_id+"'"
		rsGetBoxid.open SQL,conn,1,3
	 	if rsGetBoxid.eof and rsGetBoxid.bof then
			rsGetBoxid=""
		word="<span align='center' style='color:red;'>����Ų����ڣ��뻻һ���µ����</span>"
	    response.Redirect("ISSUE.asp?word="&word)	
		else 
		WHREC_TIME=rsGetBoxid("WHREC_TIME")
		
		GET_TIME=rsGetBoxid("GET_TIME")
		
		
			if WHREC_TIME<>"" and  isnull(GET_TIME) then
			word="<span align='center' style='color:red;'>������ڴ�ⷿ��������ȡ</span>"
	   		response.Redirect("ISSUE.asp?word="&word)	
			end if 
		
		end if 




					   if ISSUE_ID="" then
							'����box id						
							countType="ISSUE"
							countCondition=countType&formatdate(Now,"ymmdd")
							sql="select count_value,lm_time,count_type,count_condition from serial_count "
							sql=sql+"where count_type='"&countType&"' and count_condition = '"&countCondition&"'"
							rs.open sql,conn,1,3
							if rs.eof then
								ISSUE_ID=countCondition&"001"
								rs.addNew
								rs("count_type")=countType
								rs("count_condition")=countCondition
								rs("count_value")=1
								rs("lm_time")=now()
							else
							    rs("count_value")=clng(rs("count_value"))+1
								ISSUE_ID=countCondition&repeatstring(rs("count_value"),"0",3)
								
								rs("lm_time")=now()
							end if
							rs.update
							rs.close													
						end if


if SNNOA<>"" then
     if trim(PTCSTATE)<>"�ȴ�ȷ�Ͻ��" then
word="<span align='center' style='color:red;'>�˵����Ѿ�ʹ�ù����뻻һ���µĵ���</span>"
	response.Redirect("SendProd.asp?word="&word)
	end if
end if


	
SQL="select * from USERS where USER_CODE='"&code&"'"
set rs4=server.createobject("adodb.recordset")
rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" �����ڣ�����ϵ����ʦ��</span>"
	response.Redirect("ISSUE.asp?word="&word)
else		
OperatorCode=request.Form("txtOperatorCode")	
end if 
%>
<%
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript"  type="text/javascript">
	function CheckForm()
{
    
   if (document.form1.txtComments.value=="")
  {
    alert("��ע����Ϊ�գ�");
	document.form1.txtComments.focus();
	return false;
  }
  if (confirm("ȷ��Ҫ�ύ�����¼�쳣�ţ�"))
   { 
			document.form1.action="ISSUE2.asp";
			document.form1.submit();
	}
}
	 
	
	 
</script>

</head>
<body bgcolor="#339966">
<table width="80%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1" action="SendProd_Save.asp"  >
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(�쳣������¼)</td>
    </tr>
    
	 <tr>
      <td colspan="2"><BR>
	  
	  
	  </td>
    </tr>
	<tr>
		<td>
		<table width="100%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			<tr>
			  <td>����</td>
				<td>���</td>
				<td>�쳣��</td>
				<td>Comments(��ע)</td>
			</tr>
				<tr>
				<td width="8%"><input name="LM_USER" type="text" id="LM_USER"  value="<%=request("LM_USER")%>"  readonly="True" size="15"></td>
				<td><input name="BOX_ID" type="text" id="BOX_ID"  value="<%=request("BOX_ID")%>"  readonly="True" size="17"></td>
				<td width="15%"><input name="ISSUE_ID" type="text" id="ISSUE_ID" value="<%=ISSUE_ID%>" readonly size="20" ></td>
				<td><input name="txtComments" type="text" id="txtComments" value="<%=txtComments%>" size="50"  maxlength="100"></td>
			</tr>
		  </table>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<Td colspan="2" align="center"><label>
		  <input name="btnSubmit" type="button" id="btnSubmit"  value="�ύ" onClick="CheckForm();">
		</label>
		  &nbsp;
			<input name="btnClose2" type="button"  id="btnClose2"  value="������һҳ��" onClick="javascript:history.go(-1);">
			&nbsp;
			<input name="btnClose" type="button"  id="btnClose"  value="Close(�ر�)" onClick="window.close();">
	  </Td>
	</tr>
	 <input type="hidden" id="txtSubJobList" name="txtSubJobList" value=<%=SubJob%>>
	 <input type="hidden" id="txtOperatorCode" name="txtOperatorCode" value=<%=OperatorCode%>>
	   <input type="hidden" id="computername" name="computername" value=<%=ComputerName%>>
  </form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->