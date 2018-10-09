<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript"  type="text/javascript">

	function SaveData()
	{
		//document.form1.computername.value="KES-IT-05";
	 	if(document.form1.computername.value=="")
		{
			window.alert("Can not get computer name of this machine,Please contact with IT. \n没有能获取到机器名，请联系IT！");
			return;
		}
		  
		
	   if(!document.form1.txtOperatorCode.value){
			alert("Operator Code cannot be blank!\n工号不能为空！");
			document.form1.txtOperatorCode.focus();
			return false;
		}		
		
		
		else{
			
			document.form1.action="Admincode2.asp";
			document.form1.submit();
		}
	}
	
	 
</script>
</head>
<body onLoad="form1.BoxNo.focus();" bgcolor="#339966">

<table border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1">  	
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(物料接收)</td>
    </tr>
	<%if request.QueryString("word") <> "" then%>
	<Tr>
		<td height="20" colspan="2"><%=request.QueryString("word")%></td>
	</Tr>
	<%end if%>
    <tr>
      <td >状态</td>
      <td ><select name="PTCstate">
	   <option value="通过单号查询">通过单号查询</option>
	   <option value="所有未归还">所有未归还</option>
	   <option value="等待接收">等待接收</option>		
        <option value="已接收">已接收</option>
		<option value="正在归还">正在归还</option>
		<option value="已归还">已归还</option>
      </select></td>
    </tr>
	<tr>
      <td >单号</td>
      <td ><input name="SNNO" type="text" id="SNNO" value=<%=SNNO%> ></td>
    </tr>
	<tr>
      <td >Operator Code 工号</td>
      <td ><input name="txtOperatorCode" type="text" id="txtOperatorCode" value=<%=OperatorCode%> ></td>
    </tr>
	<tr><td colspan="2">&nbsp;</td></tr>
    <tr>
      <td colspan="2" align="center">
	  <input type="hidden" id="newJob" name="newJob">
      <input type="hidden" id="batchNo" name="batchNo">
	  <input type="hidden" id="computername" name="computername">
	  <input name="btnSubmit" type="button" id="btnSubmit"  value="Next 下一步" onClick="SaveData();">  
	  &nbsp;
	  <input name="btnclose" type="button"  id="btnclose"  value="Close 关闭" onClick="javascript:window.close();"> 
          
	  </td>
    </tr>
  </form>
</table>
</body>
<script>
<%if Request.Cookies("computer_name") = "" then %>
	var wsh=new ActiveXObject("WScript.Network"); 
	document.form1.computername.value=wsh.ComputerName; 
<%else%>
	document.form1.computername.value="<%=Request.Cookies("computer_name")%>"; 
<%end if%>
</script>
</html>