<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<link rel="Stylesheet" type="text/css" href="../Scripts/jquery-ui-1.9.2/css/ui-lightness/jquery-ui-1.9.2.custom.min.css" />

<body bgcolor="#339966">
<%

%>
<form id="form1" method="post"  action="PackingOperationCheckABCD.asp">
<input type="hidden" id="txt_pack_type"  name="txt_pack_type" value="<%=request("rad_pack_type")%>" />
  <table id="table1" width="800px" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Packing 输入包装箱号
			&nbsp;&nbsp;&nbsp;<a style="color:#6495ED;" href="SelectPackType.asp"></a></td>
      </tr>
      <Tr>
		<td height="20" colspan="2"><%=request.QueryString("word")%></td>
	</Tr>
      <tr >
        <td align="right">Box Id<span class="red"> *</span></td>
        <td><input type="text" id="txt_boxid" size="22" name="txt_boxid" value="" /></td>
		<td align="right">&nbsp;</td>
        <td>&nbsp;</td>		        
      </tr>	  
    </thead>
    <tbody align="center">
    </tbody>
    <tfoot>	
      <tr>
        <td colspan="4" align="center" >&nbsp;</td>		
      </tr>
    </tfoot>
  </table>

  <center>
    <label>
    <input type="submit" name="Submit" value="提交" />
    </label>
  </center>
  <input type="hidden" id="computername" name="computername" />
</form>
</body>
</html>
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>