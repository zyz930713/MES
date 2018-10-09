<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<%
BOX_ID = trim(request("Box_ID"))
PART_NUMBER=trim(request("PART_NUMBER"))
Job_number =trim(request("Job"))	
Qty=cint(trim(request("Qty")))
rtnValue=""

'get job number by tray id
if BOX_ID = "" or PART_NUMBER="" or Job_number="" or Qty="" then 
	rtnValue = "信息不能为空。"
else
    
	Packed_time=mid(BOX_ID,7,7)
	Packed_time=201&Packed_time
	'response.Write(Packed_time)
	'response.End()
	sql="select to_date('"&Packed_time&"','yyyy-mm-dd hh24:mi:ss') ndate from dual;"
	rs.open sql,conn,1,3
	if not rs.eof then	
	PACKED_TIME=rs("ndate")	
	end if
	rs.close
	
	if rtnValue="" then
    sql= "insert into JOB_PACK_DETAIL (BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,PACKED_TIME) values ('"&BOX_ID&"','"&Job_number&"','"&PART_NUMBER&"',"&Qty&",'"&PACKED_TIME&"')"
    conn.execute(sql)
	end if

	
end if
%>
<script type="text/javascript">
	window.returnValue="<%=rtnValue%>";
	window.close();
</script>
<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
