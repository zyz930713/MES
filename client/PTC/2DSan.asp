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
			window.alert("Can not get computer name of this machine,Please contact with IT. \nû���ܻ�ȡ��������������ϵIT��");
			return;
		}
		  
		
	if(!document.form1.Barcode.value){
			alert("Please input Barcode  Number!\n������2D�룡");
			document.form1.Barcode.focus();
			return false;
		
		}else if(!document.form1.txtOperatorCode.value){
			alert("Operator Code cannot be blank!\n���Ų���Ϊ�գ�");
			document.form1.txtOperatorCode.focus();
			return false;
		}		
		
		
		else{
			
			document.form1.action="2DSan2.asp";
			document.form1.submit();
		}
	}
	
	 
</script>
</head>
<body onLoad="form1.BoxNo.focus();" bgcolor="#339966">

<table width="500" height="130" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1">  	
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">2D���ѯ</td>
    </tr>
	<%if request.QueryString("word") <> "" then%>
	<Tr>
		<td height="20" colspan="2"><%=request.QueryString("word")%></td>
	</Tr>
	<%end if%>
    <tr>
      <td >2D��򵥺�</td>
      <td ><input name="Barcode" size="25" type="text" id="Barcode" value=<%=Barcode%>  ></td>
    </tr>
	
	<tr>
      <td >Operator Code ����</td>
      <td ><input name="txtOperatorCode" size="25" type="text" id="txtOperatorCode" value=<%=OperatorCode%> ></td>
    </tr>
	<tr><td colspan="2">&nbsp;</td></tr>
    <tr>
      <td colspan="2" align="center">
	  <input type="hidden" id="newJob" name="newJob">
      <input type="hidden" id="batchNo" name="batchNo">
	  <input type="hidden" id="computername" name="computername">
	  <input name="btnSubmit" type="button" id="btnSubmit"  value="Next ��һ��" onClick="SaveData();">  
	  &nbsp;
	  <input name="btnclose" type="button"  id="btnclose"  value="Close �ر�" onClick="javascript:window.close();"> 
          
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