<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->

<%
code=request.Form("txtOperatorCode")
JOB_number=trim(request("JOB_number"))




'if isnull(district) or instr(district,",PTC_Packup")=0 then
'	word="<span align='center' style='color:red;'>Please contact engiener to set the printer for this machine.<br>请联系工程师,此机器不能查询。</span>"
'	response.Redirect("Packup.asp?word="&word)
'end if 
	
SQL="select * from Users where USER_CODE='"&code&"'"
set rs4=server.createobject("adodb.recordset")

rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not permissions, please contact engineer.<br>"&code&" 不存在，请联系工程师。</span>"
	response.Redirect("RU1.asp?word="&word)

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

<script language="javascript" type="text/javascript">


	function SaveData()
	{
	
	     if(confirm("确定要删除吗？"))
   { 
			document.form1.action="RU_Save.asp?Action=EDIT";
			document.form1.submit();
	}
	
	
	
	}

</script>



</head>
<body bgcolor="#339966">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1" action=""  >
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(物料查询操作)</td>
    </tr>
    
	 <tr>
      <td colspan="2"><BR></td>
    </tr>
	
	
	
	<%
	 set rs=server.CreateObject("adodb.recordset")
	 
JOB_number=request("JOB_number")
	  
	 
			sql=" select FINAL_SCRAP_QUANTITY,REWORK_GOOD_QUANTITY from JOB_master where JOB_number='"&JOB_number&"'"
	 
	 rs.open sql,conn,1,1
	 if rs.bof then
     Response.Write "<p align='center' class='contents'>工单号不对</p>"
   	 else
	 
	
	 
	%>
	
	<tr>
		<td><table width="95%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
          <tr>
            <td width="43%">工单号</td>
            <td>报费数量</td>
            <td>反工数量</td>
            <td width="2%">&nbsp;</td>
            <td width="11%">(操作)</td>
          </tr>
          <%do while not rs.eof%>
          <tr>
            <td><%=request("JOB_number")%><input type="hidden" id="JOB_number" name="JOB_number" value="<%=request("JOB_number")%>"></td>
            <td width="26%"><input type="text" name="FINAL_SCRAP_QUANTITY" value="<%=rs("FINAL_SCRAP_QUANTITY")%>"></td>
            <td width="18%"><input type="text" name="REWORK_GOOD_QUANTITY" value="<%=rs("REWORK_GOOD_QUANTITY")%>"></td>
            <td>&nbsp;</td>
            <td><BR><input name="btnclose2" type="button"  id="btnclose2"  value="确定" onClick="SaveData();"></td>
		  </tr>            
          <% 
			rs.movenext
		    loop
		    %>
        </table>
	  <%end if%>	  </td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<Td colspan="2" align="center"><label>
		  
		</label>
		  
		  &nbsp;
			
			&nbsp;
			<input name="btnClose" type="button"  id="btnClose"  value="Close(关闭)" onClick="window.close();">	  </Td>
	</tr>
	 <input type="hidden" id="txtSubJobList" name="txtSubJobList" value=<%=SubJob%>>
	 <input type="hidden" id="txtOperatorCode" name="txtOperatorCode" value=<%=OperatorCode%>>
	   <input type="hidden" id="computername" name="computername" value=<%=ComputerName%>>
  </form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->