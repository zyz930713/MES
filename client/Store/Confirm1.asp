<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
session.Abandon()
if request.QueryString("type")="" then
response.Redirect("/SelectStoreConfirmationType.asp")
end if
pagename="Confirm1.asp"
errorstring=request.QueryString("errorstring")
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/Store.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCMD.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="javascript">
function pageload()
{
	document.form1.scantype[0].checked=true;
	document.form1.TicketID.disabled=false;
	document.form1.TicketID.focus();
}	
function formcheck()
{
	with(document.form1)
	{
		if (scantype[0].checked&&TicketID.value=="")
		{
			alert("Ticket ID Number cannot be blank!\n��ⵥ��Ų���Ϊ��");
			ID.focus();
			return;
		}
		
		if (scantype[1].checked&&ID.value=="")
		{
			alert("ID Number cannot be blank!\n������ⲻ��Ϊ��");
			ID.focus();
			return;
		}
		
		if(code.value=="")
		{
			alert("Operator Code cannot be blank!\n����Ա���Ų���Ϊ�գ�");
			code.focus();
			return false;
		}
	}
}
function typechange()
{
	with(document.form1)
	if (scantype[0].checked==true)
	{
		TicketID.disabled=false;
		TicketID.focus();
		Next.disabled=true;
		ID.disabled=true;
	}
	else
	{
		TicketID.disabled=true;
		ID.disabled=false;
		Next.disabled=false;
		ID.focus();
	}
	
}
function getTicket()
{
	with(document.form1)
	{
		if (TicketID.value!="")
		{
		var jfunction=queryTicket;
		var url="/Functions/GetAssemblyStoreTicket.asp?ticketID="+TicketID.value;
		getXMLResponse(url,true,jfunction);
		}
		else
		{
		document.all.inner_ticket.innerText="";
		}
	}
}

function queryTicket()
{
	if(http_request.readyState==4)
	{	
		if(http_request.status==200)
		{
			if(http_request.responseText!="")
			{
				//Response.Charset = "GB2312";
				document.all.inner_ticket.innerHTML=http_request.responseText;
				if (http_request.responseText!="ȷ��ҳ��ѡ�����")
				{
					document.form1.Next.disabled=false;
				}
				else
				{
					document.form1.Next.disabled=true;
				}
			}
		}
	}
}
</script>
</head>

<body onLoad="pageload();" onKeyPress="keyhandler();">
<span id="erroralarm"></span>
<form action="/Store/Confirm2.asp"  method="post" name="form1" target="_self" >
<table width="98%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-green">�Զ� Storing Confirming �Զ����ȷ�� </td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-c-green">Scan ID info 
      ɨ��������</td>
    </tr>
	<%if errorstring<>"" then%>
  <tr>
    <td height="20" colspan="2" class="strongred"><div align="center"><%=errorstring%></div></td>
    </tr>
	<%end if%>
    <tr>
      <td rowspan="2"><input type="radio" name="scantype" id="scantype" value="0" onClick="typechange()">
      ��ⵥ��� <span class="red">*</span></td>
      <td><input name="TicketID" type="text" disabled id="TicketID" accesskey="j" tabindex="1" onFocus="focushandler(this)" onBlur="blurhandler(this)" onKeyPress="tabhandler(1,this,document.form1.code)" size="50" autocomplete="off" onChange="getTicket()"></td>
    </tr>
    <tr>
      <td><span id="inner_ticket"></span>&nbsp;</td>
    </tr>
    <tr>
    <td><!--<span class="d_link">J</span>ob Number<br>-->
      <input type="radio" name="scantype" id="scantype" value="1" onClick="typechange()">
      ���������� <span class="red">*</span></td>
    <td><input name="ID" type="text" disabled id="ID" accesskey="j" tabindex="1" onFocus="focushandler(this)" onBlur="blurhandler(this)" onKeyPress="tabhandler(1,this,document.form1.code)" size="50" autocomplete="off"><br></td>
  </tr>
  <tr>
    <td width="17%"><!--<span class="d_link">O</span>perator Code<br>-->
      ���� <span class="red">*</span> </td>
    <td>
      <input name="code" type="text" id="code" accesskey="o" tabindex="2" size="10"  autocomplete="off" onFocus="focushandler(this)" onBlur="blurhandler(this)"><br></td>
  </tr>
  <tr>
    <td height="20"><!--Note<br>-->
      ע��</td>
    <td height="20"><input name="note" type="text" id="note" onFocus="focushandler(this)" onBlur="blurhandler(this)"></td>
  </tr>
  
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="Next" type="submit" accesskey="n" tabindex="5" value="ȷ ��" onclick=" return formcheck()" disabled>
      &nbsp;
      <input name="Reset" type="reset" accesskey="r" tabindex="6" value="��  ��">
      &nbsp;
      <input name="Close" type="button" id="Close" tabindex="7" onClick="javascript:window.close()" value="��  ��">
      &nbsp;&nbsp;
  <input name="Search" type="button" id="Search" value="ȷ�ϲ�ѯ" onClick="javascript:window.open('ConfirmRecords.asp')">
      &nbsp;</div></td>
  </tr>
  </table>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  
  <tr>
    <td colspan="2" class="darkbule">֧�ֵ绰��18910736955 (24Сʱ);</td>
    </tr>
  <tr>
    <td><div align="left"><span class="darkbule style1">IP Address: <%=request.ServerVariables("REMOTE_HOST")%></span></div>      </td>
    <td><div align="right"><span class="darkbule style1">System Name: <%=request.ServerVariables("HTTP_HOST")%></span></div></td>
  </tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->