<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
	response.Expires=0
	response.CacheControl="no-cache"
	pagename="Station.asp"
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
</head>
<%
SQL = "select st.nid,st.station_name,st.station_chinese_name, sc.section_name from station_new st,section sc "
SQL = SQL+"where st.section_id = sc.nid and st.is_delete = 0 and st.transaction_type=0 and sc.status = 1 "
SQL = SQL+"order by sc.nid,st.wip_sequency"
rs.open SQL,conn,1,3
section = ""
%>

<body bgcolor="#339966">
<form action="Station1.asp" method="post" name="form1">
<table border="0" align="center"><tr>
  <td><table>	
    <tr>
      <td height="20" ><div align="center">&nbsp;&nbsp;&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20" ><div align="center">
        &nbsp;
        <input type="button" class="t-b-midautumn" value="Out Pot Assy" onClick="javascript:window.open('Outer_Pot_Print.asp')">
        &nbsp;	         
        <input type="button" class="t-b-midautumn" value="Inner Pot Assy" onClick="javascript:window.open('Inner_Pot_Print.asp')">
        &nbsp;</div></td>
    </tr>
     <tr>
      <td height="20" ><div align="center">&nbsp;&nbsp;&nbsp;</div></td>
    </tr>
	<tr>
      <td height="20" ><div align="center">
        <input type="button" class="t-b-midautumn" value="Hektor&amp;Maple Pot assy" onClick="javascript:window.open('ALL_Pot_Print.asp')">
        &nbsp;
<input type="button" class="t-b-midautumn" value="WinsletMarlgold Pot assy" onClick="javascript:window.open('Pot_Print.asp')">
		  &nbsp;</div></td>
    </tr>
    <tr>
      <td height="20" ><div align="center">&nbsp;&nbsp;</div></td>
    </tr>
	<!--
	<tr>
      <td height="20" ><div align="center">
		  <input type="button" class="t-b-midautumn" value="Print Scrap List 打印报废清单" onClick="javascript:window.open('/Scrap/PrePrintScrapList.asp')">&nbsp;
          <input type="button" class="t-b-midautumn" value="Print Pack List 打印包装清单" onClick="javascript:window.open('/Store/PrintPackList.asp')">&nbsp;          
        </div></td>
    </tr>
	-->
  </table>
</td></tr></table>
<input type="hidden" id="computername" name="computername">
</form>	
</body>
<script language="javascript">
function formCheck()
{
	var isSltStation = false;
	for(var i=1;i<<%=stationCount%>;i++){
		if(form1.radStationId[i-1].checked){
			isSltStation = true;
			break;
		}
	}
	if(!isSltStation){
		alert("Please select one station! \n请选择一个工作站！");
	}
	return isSltStation;
}
</script>
<script>
	var wsh=new ActiveXObject("WScript.Network"); 
	document.form1.computername.value=wsh.ComputerName; 
</script>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->