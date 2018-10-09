<!--#include file= "conn.asp"--> 
<!--#include file= "../include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>LPB</title>
</head>
<body>
<%
	AddMachine = request("AddMachine")
	ProductName = request("ProductName")
	LineName = request("LineName")
	DDN = request("DDN")
	HShift = request("HShift")
	Ddate = request("Ddate")
	
	MdId = request("MdId")
	LineNo = ProductName & LineName
	XlType = request("XlType")
	Dhtime=request("Dhtime")
	Detime=request("Detime")
	Dequipment=request("Dequipment")
	DPosition=request("DPosition")
	Dproblem=request("Dproblem")
	DReason=request("DReason")
	Dsolution=request("Dsolution")
	Dtime=request("Dtime")
	PcSum = request("PcSum")
	PcQty = request("PcQty")
	PlanStop = request("PlanStop")

	'增加设备维护信息'
	if AddMachine = "增加" then
		if ProductName = ""  or LineName = "" or DDN = "" or Ddate = "" or HShift = "" or XlType = "" or Dhtime = "" or Detime = "" or Dequipment = "" or DPosition = "" or Dproblem = "" or DReason = "" or Dsolution = "" or Dtime = "" or (NOT IsNumeric(Dtime)) then
			call errorHistoryBack("有空数据或停机时间不为空,请重新填写!")
		else
			Sqlstr = "INSERT INTO Ministat.dbo.BJLPBD2(Ddate,Udate,DLineno,DD_N,HShift,Dlossn,DlossSN,Dhtime,Detime,Dequipment,DPosition,Dproblem,DReason,Dsolution,Dtime,Dqty)"
			Sqlstr = Sqlstr & " VALUES('"&Ddate&"',getdate(),'"&LineNo&"','"&DDN&"','"&HShift&"','"&Xltype&"','1','"&Dhtime&"','"&Detime&"','"&Dequipment&"','"&DPosition&"','"&Dproblem&"','"&DReason&"','"&Dsolution&"','"&Dtime&"','0')"
			conn.execute(Sqlstr)
		end if
		call sussLoctionHref("添加成功","LpbMachine.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN&"&HShift="&HShift)
	end if
	'增加设备维护信息'
	
	'增加主载体信息'
	if AddMachine = "增 加" then
		if ProductName = ""  or LineName = "" or DDN = "" or Ddate = "" or HShift = ""  then
			call errorHistoryBack("有空数据,请重新填写!")
		else
			set rs = Server.CreateObject("adodb.recordset")
			sql = "SELECT top 1 * FROM Ministat.dbo.BJPC where Ddate = '"&Ddate&"' and DD_N = '"&DDN&"' and Dlineno = '"&LineNo&"' and HShift = '"&HShift&"'"
			rs.open sql,conn,1,1
			if rs.eof then
				Sqlstr = "INSERT INTO Ministat.dbo.BJPC(Ddate,Udate,DLineno,DD_N,HShift,PlanStop,PcSum,PcQty)"
				Sqlstr = Sqlstr & " VALUES('"&Ddate&"',getdate(),'"&LineNo&"','"&DDN&"','"&HShift&"','"&PlanStop&"','"&PcSum&"','"&PcQty&"')"
				conn.execute(Sqlstr)
			end if
			rs.close
			set rs = nothing
		end if
		call sussLoctionHref("添加成功","LpbMachine.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN&"&HShift="&HShift)
	end if
	'增加主载体信息'
	
	'修改设备维护信息'
	if AddMachine = "修改" then
		Sqlstr = "update Ministat.dbo.BJLPBD2 set Dlossn='"&Xltype&"',Dhtime='"&Dhtime&"',Detime='"&Detime&"',Dequipment='"&Dequipment&"',DPosition='"&DPosition&"',Dproblem='"&Dproblem&"',DReason='"&DReason&"',Dsolution='"&Dsolution&"',Dtime='"&Dtime&"' where id = '"&MdId&"'"
		conn.execute(Sqlstr)
		call sussLoctionHref("修改成功","LpbMachine.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN&"&HShift="&HShift)
	end if
	'修改设备维护信息'
	
	'修改主载体信息'
	if AddMachine = "修 改" then
		Sqlstr = "update Ministat.dbo.BJPC set PcSum='"&PcSum&"',PcQty='"&PcQty&"',PlanStop='"&PlanStop&"' where Ddate = '"&Ddate&"' and DD_N = '"&DDN&"' and Dlineno = '"&LineNo&"' and HShift = '"&HShift&"'"
		conn.execute(Sqlstr)
		call sussLoctionHref("修改成功","LpbMachine.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN&"&HShift="&HShift)
	end if
	'修改主载体信息'
	
	'提取Reamrk'
	if AddMachine = "remark" or AddMachine="确认写入到Remark" then
		dim RemarkStr,RemarkStrSum
		RemarkStr = ""
		RemarkStrSum = ""
		Dlineno = ProductName & LineName
		
			set rs = Server.CreateObject("adodb.recordset")
			sql = "SELECT * FROM [Ministat].[dbo].[BJPC] where Dlineno = '"&Dlineno&"' and DD_N = '"&DDN&"' AND Hshift = '"&HShift&"' and Ddate = '"&Ddate&"'"
			rs.open sql,conn,1,1
			if not rs.eof then
				PcSum = rs("PcSum")
				PcQty = rs("PcQty")
				PlanStop = rs("PlanStop")
			end if
			rs.close
			set rs = nothing
		
		set rs = Server.CreateObject("adodb.recordset")
		sql = "SELECT * FROM Ministat.dbo.BJLPBD2 where Dlineno = '"&Dlineno&"' and DD_N = '"&DDN&"' AND Hshift = '"&HShift&"' and Ddate = '"&Ddate&"' order by id"
		rs.open sql,conn,1,1
		if not rs.eof then
			RemarkStrSum = Ddate & "　" & Dlineno & "　Shift: " & rs("HShift") & "　D/N: " & rs("DD_N") & " </BR></BR>"
			RemarkStrSum = RemarkStrSum & "计划停机: " & PlanStop & " 小时</br></br>"
			do while not rs.eof
				HShift = rs("HShift")
				Dlossn = rs("Dlossn")
				Ddate = FormatTime(rs("Ddate"),2)
				Dhtime = rs("Dhtime")
				Detime = rs("Detime")
				Dtime = rs("Dtime")
				Dequipment = rs("Dequipment")
				DPosition = rs("DPosition")
				Dproblem = rs("Dproblem")
				DReason = rs("DReason")
				Dsolution = rs("Dsolution")
					
					RemarkStr = Dhtime & " - " & Detime & "　停机: " & Dtime & " 分钟</br>"
					RemarkStr = RemarkStr & "位置:" & Dequipment & " " & DPosition & "</br>"
					RemarkStr = RemarkStr & "问题:" & Dproblem & "</br>"
					RemarkStr = RemarkStr & "原因:" & DReason & "</br>"
					RemarkStr = RemarkStr & "解决:" & Dsolution & "</br>"
					RemarkStr = RemarkStr & "</br>"
			
				RemarkStrSum = RemarkStrSum & RemarkStr
			rs.movenext
			loop
			
			
			'提取封存数据'
			BQTY = 0
			RQTY = 0
			set rsq = Server.CreateObject("adodb.recordset")
			SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.QABlock where CType = 'Block' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
			rsq.open SqlStr,conn,1,1
			if not rsq.eof then BQTY = rsq("CQTY")
			rsq.close

			SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.QABlock where CType = 'Release' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
			rsq.open SqlStr,conn,1,1
			if not rsq.eof then RQTY = rsq("CQTY")
			rsq.close
			set rsq = nothing
			
			if BQTY <> 0 then RemarkStrSum = RemarkStrSum & "封存数量: " & BQTY & "</br></br>"
			if RQTY <> 0 then RemarkStrSum = RemarkStrSum & "放行数量: " & RQTY & "</br></br>"
			'提取封存数据'

			RemarkStrSum = RemarkStrSum & "在线主载体数量: " & PcSum & "　待修主载体数量: " & PcQty & "</br>"
			rs.close
			set rs = nothing

			' if AddMachine="确认写入到Remark" then
				' Sqlstr = "update Ministat.dbo.BJPC set MRemark='"&RemarkStrSum&"' where Ddate = '"&Ddate&"' and DD_N = '"&DDN&"' and Dlineno = '"&Dlineno&"' and HShift = '"&HShift&"'"
				' conn.execute(Sqlstr)
				' call sussLoctionHref("写入成功","LpbMachine.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN&"&HShift="&HShift)
			' end if

			'response.write "<textarea cols='100' id='biao1'>"&RemarkStrSum&"</textarea>"
			'response.write "<a href='javascript:;' onclick='copyUrl2()'>点击复制</a>"
			
		end if

		' response.write "<form id='form' name='form1' method='post' action='AddMachine.asp'>"
		' response.write "<input type='hidden' name='Ddate' value='"&Ddate&"'>"
		' response.write "<input type='hidden' name='ProductName' value='"&ProductName&"'>"
		' response.write "<input type='hidden' name='LineName' value='"&LineName&"'>"
		' response.write "<input type='hidden' name='DDN' value='"&DDN&"'>"
		' response.write "<input type='submit' name='AddMachine' value='确认写入到Remark'>"
		' response.write "</form>"

		response.write "<input type='button' value='返回' onClick='history.back(-1)' />　"
		
		response.write "<input type=button value='复制到剪切板' onclick='cycode(ip);'>"
		response.write "<div id='ip'>" & RemarkStrSum & "</div>"

	end if
%>
		<script>
		function cycode(obj)
		{
			var rng = document.body.createTextRange();
			rng.moveToElementText(obj);
			rng.select();
			rng.execCommand('Copy');
		}
		</script>

</body>
</html>