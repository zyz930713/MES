<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->

<%
code=request.Form("txtOperatorCode")
PTCstate=request("PTCstate")
SNNO=request("SNNO")
ComputerName=request("computername")
district=","&GetdistrictName(ComputerName)




'if isnull(district) or instr(district,",PTC_Packup")=0 then
'	word="<span align='center' style='color:red;'>Please contact engiener to set the printer for this machine.<br>请联系工程师,此机器不能查询。</span>"
'	response.Redirect("Packup.asp?word="&word)
'end if 
	
SQL="select * from Users where USER_CODE='"&code&"'"
set rs4=server.createobject("adodb.recordset")

rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not permissions, please contact engineer.<br>"&code&" 不存在，请联系工程师。</span>"
	response.Redirect("Packup1.asp?word="&word)
else
      session("role")=","&rs4("ROLES_ID")
      if instr(session("role"),",PACKING_ADMINISTRATOR")<>0  then
	
	OperatorCode=request.Form("txtOperatorCode")
	
	else
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" 没有权限，请联系工程师。</span>"
	response.Redirect("Packup1.asp?word="&word)
		
    end if
end if 
%>
<%
	
	
	
	function GetdistrictName(ComputerName)
		set rsdistrictName=server.CreateObject("adodb.recordset")
		SQL="SELECT * FROM COMPUTER_PRINTER_MAPPING WHERE UPPER(COMPUTER_NAME)='"+UCase(ComputerName)+"'"
	 
		rsdistrictName.open SQL,conn,1,3
	 	if(rsdistrictName.recordcount=0) then
			GetdistrictName=""
			exit function
		end if 
		    
		if (rsdistrictName(1)="") then
			GetdistrictName=""
			exit function
		end if 
		
		'if (rsdistrictName(1)="") then
		
		
		'end if
		GetdistrictName=rsdistrictName(1)
	
	end function
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
			document.form1.action="Packup_Save.asp?Action=DEL";
			document.form1.submit();
	}
	
	
	
	}

</script>



</head>
<body bgcolor="#339966">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1" action="Packup_Save.asp"  >
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(物料查询操作)</td>
    </tr>
    
	 <tr>
      <td colspan="2"><BR></td>
    </tr>
	
	
	
	<%
	 set rs=server.CreateObject("adodb.recordset")
	 
box_id=request("box_id")
	  
	 
			sql=" select COUNT(*) sn from job_2d_code where box_id='"&box_id&"'"
	 
	 rs.open sql,conn,1,1
	 if rs("SN")= "0"  then
     Response.Write "<p align='center' class='contents'> 数据库中无数据！</p>"
   	 else
	 
	
	 
	%>
	
	<tr>
		<td><table width="95%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
          <tr>
            <td>箱号</td>
            <td>打包数量</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td width="6%">&nbsp;</td>
            <td width="7%">&nbsp;</td>
            <td>(操作)</td>
          </tr>
          <%do while not rs.eof%>
          <tr>
            <td><%=request("box_id")%><input type="hidden" id="box_id" name="box_id" value="<%=request("box_id")%>"></td>
            <td><%=rs("sn")%></td>
            <td width="4%">&nbsp;</td>
            <td width="8%">&nbsp;</td>
            <td width="6%">&nbsp;</td>
            <td width="5%">&nbsp;</td>
            <td width="9%">&nbsp;</td>
            <td>&nbsp;</td>
			<td>&nbsp;</td>
			<td><BR><input name="btnclose2" type="button"  id="btnclose2"  value="确定删除箱号" onClick="SaveData();"></td>
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