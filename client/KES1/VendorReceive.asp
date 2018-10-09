<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript" type="text/javascript" src="/Components/My97DatePicker/WdatePicker.js"></script>

<script type="text/javascript">

</script>
</head>

<body bgcolor="#339966">
<form id="form1" method="post">
<table id="table1" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Vendor Receive 客户签收</td>
      </tr>
      <tr >
        <td align="right">Op Code 工号<span class="red"> *</span></td>
        <td><input type="text" id="txt_opcode" name="txt_opcode" /></td>
	  </tr>
	  <tr >
	  	<td align="right">Delivery No 提货单号</td>
        <td><input type="text" id="txt_delivery" name="txt_delivery" /></td>        
      </tr>
	  <tr >
        <td align="right">Receive Time 签收时间</td>
        <td><input type="text" id="txt_rcvTime" name="txt_rcvTime" readonly="true"/>
		<img onclick="WdatePicker({el:'txt_rcvTime',dateFmt:'yyyy-MM-dd HH:mm:ss',lang:'en'})" src="/Images/dynCalendar.gif" align="absmiddle" style="cursor:pointer"/>		
		</td>
	  </tr>	  		
	<tfoot>	
       <tr>
        <td colspan="4" align="center" ><br>		
			<input type="button" id="btn_print" value="OK 确定" onClick="SubmitForm()"/>  
			&nbsp;			
			<input name="Close" type="button" id="Close" onClick="window.close()" value="Close 关闭">
		</td>		
      </tr>
    </tfoot>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->