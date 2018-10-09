<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
<!--#include virtual="/Components/JSON_2.0.4.asp" -->
<!--#include virtual="/Components/JSON_UTIL_0.1.1.asp" -->
<%
'On Error Resume Next
'ajax获取数据
response.Expires=0
response.CacheControl="no-cache"
Response.Charset="gb2312"
asynid=request.QueryString("asynid")
if	len(asynid)=0 then
	asynid=request.Form("asynid")
end if
if len(asynid)>0 then
	set rsAsyn=server.CreateObject("adodb.recordset")
	set rsAsyn1=server.CreateObject("adodb.recordset")
	dim member
	set member=jsObject()
	select case asynid
		case "1"'输入jobnumber后获取相关资料
			jobno=request.QueryString("txt_jobno")
			wipentityname=split(jobno,"-")(0)
			sheetnumber=split(jobno,"-")(1)
			sheetnumber=CInt(sheetnumber)
			'获取partno
			sql="select PrimaryItemDesc from tbl_MES_LotMaster where wipentityname='"&wipentityname&"'"
			rsAsyn.open sql,connTicket,1,3
			if not rsAsyn.eof then
				partno=rsAsyn("PrimaryItemDesc")
				sql="select quantity from tbl_mes_lotmastersub where wipentityname='"&wipentityname&"' and sheetnumber='"&sheetnumber&"'"
				rsAsyn1.open sql,connTicket,1,3
				if not rsAsyn1.eof then 
					startqty=rsAsyn1("quantity")
					member("partno")=partno
					member("startqty")=startqty
				else
					member("error")="This sub job number is not printed.该子工单没有打印."
				end if
				rsAsyn1.close
			else
				member("error")="This job number does not exist.此工单不存在."
			end if
			rsAsyn.close
			'检查是否有op code保存过
			sql="select distinct link_user from job_tray_mapping where is_disable=0 and JOB_NUMBER='"&wipentityname&"' and SHEET_NUMBER='"&sheetnumber&"'"
			rsAsyn.open sql,conn,1,3
			if not rsAsyn.eof then
				member("opcode")=rsAsyn("link_user")
			else
				member("opcode")=""
			end if
			rsAsyn.close
			response.Write toJSON(member)
		case "2"'根据partnumber获取工站明细
			partno=request.QueryString("partno")
			sql="select a.STATION_ID,a.STATION_SEQ,a.TRAY_TYPE,a.tray_size,b.station_name,b.station_chinese_name from part_tray a left join station_new b on a.station_id=b.nid where a.part_number='"&partno&"' order by b.wip_sequency, a.tray_type"
			response.Write QueryToJSON(conn, sql).Flush
			'rsAsyn.open sql,conn,1,3
			'if not rsAsyn.eof then
			'	Response.Write toJSON(rsAsyn)
			'end if
			'rsAsyn.close
		case "3"'Save，数据量可以会大，采用post传输方式
			jobnumber=split(request.Form("txt_jobno"),"-")(0)
			sheetnumber=cint(split(request.Form("txt_jobno"),"-")(1))
			opcode=request.Form("txt_opcode")
			startqty=request.Form("txt_qty")
			partno=request.Form("txt_partno")
			trayid=request.Form("txt_trayid")
			trayid=replace(trayid," ","")
			'lotno=request.Form("txt_lotno")
			'qty=request.Form("txt_qty")			
			'检查输入Tray ID是否与其他sub job number绑定，若无，则disabled该sub job number的绑定关系并将输入数据存入table JOB_TRAY_MAPPING；否则提示link关系已经存在。
			'判断是否有IS_DISABLE=0的tray id
			sql="select tray_id from JOB_TRAY_MAPPING where tray_id in ('"&replace(trayid,",","','")&"') and is_disable=0 and (job_number<>'"&jobnumber&"' or sheet_number<>'"&sheetnumber&"')"
			'response.Write(sql)
			rsAsyn.open sql,conn,1,3
			While Not (rsAsyn.EOF Or rsAsyn.BOF)
				tray_id=rsAsyn("tray_id")&","
			rsAsyn.MoveNext
			Wend		
			rsAsyn.close
			if len(tray_id)>0 then
				member("error")="These tray ID("&tray_id&") are link with other job numbers.|这些料盘已经与其他工单绑定."
			else 		
				sql="update JOB_TRAY_MAPPING set is_disable=1 where job_number='"&jobnumber&"' and sheet_number='"&sheetnumber&"' and is_disable=0"
				rsAsyn.open sql,conn,1,3
				for i=1 to request.Form("txt_trayid").count
					if len(trim(request.Form("txt_trayid")(i)))>0 then
						sql="insert into JOB_TRAY_MAPPING (JOB_NUMBER,SHEET_NUMBER,TRAY_ID,LINK_DATETIME,IS_DISABLE,STATION_ID,LINK_USER,MATERIAL_LOT_NO,   MATERIAL_QTY,DISPLAY_SEQ) values ('"&jobnumber&"','"&sheetnumber&"','"&request.Form("txt_trayid")(i)&"',sysdate,'0','"&request.Form("txt_stationid")(i)&"','"&opcode&"','"&request.Form("txt_lotno")(i)&"','"&request.Form("txt_mqty")(i)&"',"&request.Form("txt_displayseq")(i)&")"
						rsAsyn.open sql,conn,1,3
					end if
				next
				'rsAsyn.close
				member("message")="Save successfully.保存成功."
			end if			
			response.Write toJSON(member)
		case "4"'获取已经保存的数据
			jobnumber=split(request.QueryString("txt_jobno"),"-")(0)
			sheetnumber=cint(split(request.QueryString("txt_jobno"),"-")(1))
			'sql="SELECT  TRAY_ID, STATION_ID,  MATERIAL_LOT_NO,  MATERIAL_QTY,  DISPLAY_SEQ FROM JOB_TRAY_MAPPING where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and IS_DISABLE='0'"
			sql="select a.TRAY_ID,a.STATION_ID,a.MATERIAL_LOT_NO,a.MATERIAL_QTY,a.DISPLAY_SEQ from JOB_TRAY_MAPPING a left join part_tray b on a.station_id=b.station_id left join station_new c on b.station_id=c.nid  where b.part_number=(select part_number_tag from job_master where job_number='"&jobnumber&"') and a.JOB_NUMBER='"&jobnumber&"' and a.SHEET_NUMBER='"&sheetnumber&"' and a.IS_DISABLE='0' order by c.wip_sequency,b.tray_type,a.DISPLAY_SEQ"'连接part_tray和station_new保证顺序的一致性
			response.Write QueryToJSON(conn, sql).Flush
	end select 	
	'rsAsyn.close
	'rsAsyn1.close
end if
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->
