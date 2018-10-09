<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->

<%
code=request.Form("txtOperatorCode")
Codestate=request("Codestate")
SNNO=request("SNNO")
BarcodeL=len(Barcode)
	
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
	 if CodeState="所有" then
	 sql="select * from PTC_BarcodeNO where SNNO='"&SNNO&"'"
	 elseif  codeState="已接收" then
	 sql="select * from PTC_BarcodeNO where SNNO='"&SNNO&"' and  Codestate='已接收'"
     else
	 sql="select * from PTC_BarcodeNO where Codestate='"&Codestate&"' and SNNO='"&SNNO&"'"
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
            <td width="9%"><%=rs("SNNO")%></td>
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
		  <input name="btnClose2" type="button"  id="btnClose2"  value="返回上一页面" onClick="javascript:history.go(-1);">&nbsp;
			
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