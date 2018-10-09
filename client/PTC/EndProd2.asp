<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->

<%
code=request.Form("txtOperatorCode")
SNNO=request("SNNO")
ComputerName=request("computername")
district=GetdistrictName(ComputerName)
Sendarea=request("Sendarea")



if isnull(district) or instr(district,Sendarea)=0 then
	word="<span align='center' style='color:red;'>Please contact engiener to set the printer for this machine.<br>请联系工程师为此机器不能结束。</span>"
	response.Redirect("Endprod.asp?word="&word)
end if 
SQL="select OPERATOR_NAME,AUTHORIZED_STATIONS_ID,AUTHORIZED_PARTS_ID,LOCKED,PRACTISED,PRACTISE_START_TIME,PRACTISE_END_TIME from OPERATORS where code='"&code&"'"
set rs4=server.createobject("adodb.recordset")
rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" 不存在，请联系工程师。</span>"
	response.Redirect("Endprod.asp?word="&word)
else
	SubJob=request.Form("txtSubJobList")
	NewJob=request.Form("newJob")
	ManJobNumber=request.Form("batchNo")
	Action=request.querystring("Action")
	TotalCount=request.Form("TotalCount")	
	OperatorCode=request.Form("txtOperatorCode")	
	Set TypeLib = CreateObject("Scriptlet.TypeLib")
    Guid = TypeLib.Guid	
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
    alert("箱号不能为空！");
	document.form1.BoxNo.focus();
	return false;
  }  
  if (document.form1.LendNO.value=="")
  {
    alert("数量不能为空！");
	document.form1.LendNO.focus();
	return false;
  }
  
  if (document.form1.acceptCode.value=="")
  {
    alert("接收人不能为空！");
	document.form1.acceptCode.focus();
	return false;
  }
  
  if (document.form1.txtComments.value=="")
  {
    alert("备注不能为空！");
	document.form1.txtComments.focus();
	return false;
  }
  
  
  
  
  return true;  
}
	 
	 function checkNumber() {    //判断输入字符的keyCode，数字在48到57之间，超出部分返回false   
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
  <form id="form1" method="post" name="form1" action="EndProd_Save.asp?PTCstate=End&SNNO=<%=SNNO%>"  onSubmit="return CheckForm();">
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(物料接受操作)</td>
    </tr>
    
	 <tr>
      <td colspan="2"><BR></td>
    </tr>
	
	
	
	<%
	 set rs=server.CreateObject("adodb.recordset")
			 sql="select * from PTC_SN where SNNO='"&SNNO&"'"
	
	 rs.open sql,conn,1,1
	 if rs.eof And rs.bof then
     Response.Write "<p align='center' class='contents'> 数据库中无数据！</p>"
   	 else
	 
	
	 
	%>
	
	<tr>
		<td><table width="96%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
          <tr align="center">
            <td>单号</td>
            <td>箱号</td>
            <td>借出数量</td>
            <td>未还数量</td>
            <td>已还数量</td>
            <td>正在归还</td>
            <td>操作员工号</td>
            <td>接收人工号</td>
            <td>类型</td>
            <td>区域</td>
            <td>状态 </td>
            <td width="13%">借出时间</td>
            <td width="9%">(备注)</td>
            <td width="5%">操作</td>
          </tr>
          <%do while not rs.eof%>
          <tr align="center">
            <td width="6%"><a href="2DCodeList.asp?SNNO=<%=rs("SNNO")%>&Codestate=所有"><%=rs("SNNO")%></a></td>
            <td width="7%"><%=rs("BoxNo")%></td>
            <td width="7%"><%=rs("LendNO")%></td>
            <td width="7%"><%=rs("LendNO")-rs("ReturnNO")%></td>
            <td width="6%"><%=rs("ReturnNO")%></td>
            <td width="7%"><a href="2DCodeList.asp?SNNO=<%=rs("SNNO")%>&Codestate=正在归还"><%=rs("ReturnNOJSQ")%></a></td>
            <td width="6%"><%=rs("OperatorCode")%></td>
            <td width="7%"><%=rs("acceptCode")%></td>
            <td width="6%"><%=rs("SelectType")%></td>
            <td width="5%"><%=rs("area")%></td>
            <td width="9%"><%=rs("PTCstate")%></td>
            <td><%=rs("SendDate")%></td>
			<td><%=rs("txtComments")%></td>
            <td><br>
              <a href="Barcode_END.asp?SNNO=<%=rs("SNNO")%>">检查</a>
			<br><br>
			<%if rs("LendNO")=rs("ReturnNO") then %>
              <a href="EndProd_Save.asp"></a>
              <label><a href="EndProd_Save.asp?PTCstate=End&SNNO=<%=rs("SNNO")%>">结束</a></label>
           <%end if%><BR></td>
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