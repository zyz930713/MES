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




if isnull(district) or instr(district,",PTC_QUERY")=0 then
	word="<span align='center' style='color:red;'>Please contact engiener to set the printer for this machine.<br>����ϵ����ʦ,�˻������ܲ�ѯ��</span>"
	response.Redirect("Sanprod.asp?word="&word)
end if 
	
SQL="select * from Users where USER_CODE='"&code&"'"
set rs4=server.createobject("adodb.recordset")

rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not permissions, please contact engineer.<br>"&code&" �����ڣ�����ϵ����ʦ��</span>"
	response.Redirect("Sanprod.asp?word="&word)
else
      session("role")=","&rs4("ROLES_ID")
      if instr(session("role"),",PTC_QUERY")<>0  then
	
	OperatorCode=request.Form("txtOperatorCode")
	
	else
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" û��Ȩ�ޣ�����ϵ����ʦ��</span>"
	response.Redirect("Sanprod.asp?word="&word)
		
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
<script language="javascript"  type="text/javascript">
	function CheckForm()
{
 
  if (document.form1.BoxNo.value=="")
  {
    alert("��Ų���Ϊ�գ�");
	document.form1.BoxNo.focus();
	return false;
  }  
  if (document.form1.LendNO.value=="")
  {
    alert("��������Ϊ�գ�");
	document.form1.LendNO.focus();
	return false;
  }
  
  if (document.form1.acceptCode.value=="")
  {
    alert("�����˲���Ϊ�գ�");
	document.form1.acceptCode.focus();
	return false;
  }
  
  if (document.form1.txtComments.value=="")
  {
    alert("��ע����Ϊ�գ�");
	document.form1.txtComments.focus();
	return false;
  }
  
  
  
  
  return true;  
}
	 
	 function checkNumber() {    //�ж������ַ���keyCode��������48��57֮�䣬�������ַ���false   
	  if ((event.keyCode >= 48) && (event.keyCode <= 57))
	   {      
	     event.returnValue = true;    
		 } 
		 else 
		 {        event.returnValue = false;    }
		 
		 }
	 
	 
</script>




</head>
<body bgcolor="#339966">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1" action="acceptProd_Save.asp"  onSubmit="return CheckForm();">
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(���ϲ�ѯ����)</td>
    </tr>
    
	 <tr>
      <td colspan="2"><BR></td>
    </tr>
	
	
	
	<%
	 set rs=server.CreateObject("adodb.recordset")
	 
	 if PTCstate="����δ�黹" then
	  sql="select * from PTC_SN where PTCstate<>'�ѹ黹'"
	 elseif PTCstate="ͨ�����Ų�ѯ" and SNNO <>"" then
	 
	  sql="select * from PTC_SN where SNNO='"&SNNO&"'"
	 elseif PTCstate="HKTNPI" then
	 
	 sql="select * from PTC_SN where SENDAREA='"&PTCstate&"'"
	 elseif PTCstate="WSTNPI" then
	  sql="select * from PTC_SN where SENDAREA='"&PTCstate&"'"
	 elseif PTCstate="DBP" then
	 sql="select * from PTC_SN where SENDAREA='"&PTCstate&"'"
	 
	 elseif PTCstate="ACCEPTCODE" then
	  
	 sql="select * from PTC_SN where acceptcode='"&SNNO&"'"
	 
	 else
	  
	 
			 sql="select * from PTC_SN where PTCstate='"&PTCstate&"'"
	 end if
	 rs.open sql,conn,1,1
	 if rs.eof And rs.bof then
     Response.Write "<p align='center' class='contents'> ���ݿ��������ݣ�</p>"
   	 else
	 
	
	 
	%>
	
	<tr>
		<td><table width="95%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
          <tr>
            <td>(����)</td>
            <td>���</td>
            <td>�������</td>
            <td>δ������</td>
            <td>�ѻ�����</td>
            <td>����Ա����</td>
            <td>������</td>
            <td>�����˹���</td>
            <td>����</td>
            <td>����</td>
            <td>״̬ </td>
            <td width="6%">���ʱ��</td>
            <td width="7%">����ʱ��</td>
            <td width="6%">�黹ʱ��</td>
            <td width="7%">NPI����ʱ��</td>
            <td width="2%">(��ע)</td>
          </tr>
          <%do while not rs.eof%>
          <tr>
            <td width="6%"><a href="2DCodeList.asp?SNNO=<%=rs("SNNO")%>&Codestate=����"><%=rs("SNNO")%></a></td>
            <td width="7%"><%=rs("BoxNo")%></td>
            <td width="7%"><%=rs("LendNO")%></td>
            <td width="7%"><%=rs("LendNO")-rs("ReturnNO")%></td>
            <td width="8%"><%=rs("ReturnNO")%></td>
            <td width="5%"><a href="http://keb-dt056/admin/KEB_DisplayB.asp?EmployeeNo=<%=rs("OperatorCode")%>" target="_blank"><%=rs("OperatorCode")%></a></td>
            <td width="4%"><a href="http://keb-dt056/admin/KEB_DisplayB.asp?EmployeeNo=<%=rs("getCode")%>" target="_blank"><%=rs("getCode")%></a></td>
            <td width="8%"><a href="http://keb-dt056/admin/KEB_DisplayB.asp?EmployeeNo=<%=rs("acceptCode")%>" target="_blank"><%=rs("acceptCode")%></a></td>
            <td width="6%"><%=rs("SelectType")%></td>
            <td width="5%"><%=rs("area")%></td>
            <td width="9%"><%=rs("PTCstate")%></td>
            <td><%=rs("SendDate")%></td>
			<td><%=rs("acceptDate")%></td>
			<td><%=rs("returnDate")%></td>
			<td><%=rs("EndDate")%></td>
			<td><%=rs("txtComments")%><BR></td>
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
			<input name="btnClose" type="button"  id="btnClose"  value="Close(�ر�)" onClick="window.close();">	  </Td>
	</tr>
	 <input type="hidden" id="txtSubJobList" name="txtSubJobList" value=<%=SubJob%>>
	 <input type="hidden" id="txtOperatorCode" name="txtOperatorCode" value=<%=OperatorCode%>>
	   <input type="hidden" id="computername" name="computername" value=<%=ComputerName%>>
  </form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->