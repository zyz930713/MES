<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->

<%
code=request.Form("txtOperatorCode")
Barcode=request("Barcode")
BarcodeL=len(Barcode)
ComputerName=request("computername")
district=","&GetdistrictName(ComputerName)




	
	
	
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
	

if isnull(district) or instr(district,",PTC_QUERY")=0 then
	word="<span align='center' style='color:red;'>Please contact engiener to set the printer for this machine.<br>请联系工程师,此机器不能接收。</span>"
	response.Redirect("Sanprod.asp?word="&word)
end if 
	
SQL="select * from Users where USER_CODE='"&code&"'"
set rs4=server.createobject("adodb.recordset")
rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not permissions, please contact engineer.<br>"&code&" 不存在，请联系工程师。</span>"
	response.Redirect("Sanprod.asp?word="&word)
else
      session("role")=","&rs4("ROLES_ID")
      if instr(session("role"),",PTC_QUERY")<>0  then
	
	OperatorCode=request.Form("txtOperatorCode")
	
	else
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" 没有权限，请联系工程师。</span>"
	response.Redirect("Sanprod.asp?word="&word)
		
    end if
end if 
%>

<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">





</head>
<body bgcolor="#339966">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1" action="acceptProd_Save.asp"  onSubmit="return CheckForm();">
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(物料查询操作)</td>
    </tr>
    
	 <tr>
      <td colspan="2"><BR></td>
    </tr>
	
	
	
	<%
	 set rs=server.CreateObject("adodb.recordset")
	  if BarcodeL="17" then
	  sql="select * from PTC_BarcodeNO where Barcode='"&Barcode&"'"
	  else
	
	  sql="select * from PTC_BarcodeNO where SNNO='"&Barcode&"'"
	  end if
	
	 rs.open sql,conn,1,1
	 if rs.eof And rs.bof then
     Response.Write "<p align='center' class='contents'> 数据库中无数据！</p>"
   	 else
	 
	
	 
	%>
	
	<tr>
		<td><table width="74%" height="43" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" colspan="2">
          <tr align="center">
            <td>(单号)</td>
            <td width="21%">2D码</td>
            <td>状态 </td>
            <td width="16%">借出时间</td>
            <td width="11%">接收时间</td>
            <td width="13%">归还时间</td>
            <td width="12%">NPI接收时间</td>
            <td width="6%">&nbsp;</td>
          </tr>
          <%do while not rs.eof%>
          <tr align="center">
            <td width="9%"><a href="2DCodeList.asp?SNNO=<%=rs("SNNO")%>&Codestate=所有"><%=rs("SNNO")%></a></td>
            <td><%=rs("Barcode")%></td>
            <td width="12%"><%=rs("Codestate")%></td>
            <td><%=rs("CreatDate")%></td>
			<td><%=rs("acceptDate")%></td>
			<td><%=rs("ReturnDate")%></td>
			<td><%=rs("EndDate")%></td>
			<td><BR></td>
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