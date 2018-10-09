<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
 <%
 	SQL="SELECT station_name,station_chinese_name,actions_index from STATION"
	rs.open SQL,conn,1,3
	
 %>
 <table>
 <%
 	for i=0 to rs.recordcount-1
 %>
 	<Tr>
 	<td><%=rs(0)%></td>
 	<td><%=rs(1)%></td>
	<td>
		<%
			Action_Str=""
			Action_Chinese_Str=""
			if rs(2)<>"" then
				actionArr=split(rs(2),",")
				for j=0 to ubound(actionarr)
					SQL="SELECT ACTION_NAME,action_chinese_name FROM ACTION WHERE NID='"+actionArr(J)+"'"
					set rsAction=server.createobject("adodb.recordset")
					rsAction.open SQL,conn,1,3
					if(rsAction.recordcount<>0) then
						Action_Str=Action_Str+rsAction(0)+","
						Action_Chinese_Str=Action_Chinese_Str+rsAction(1)+","
					end if 
					rsAction.close()
					set rsAction=nothing
				next 
			end if
		%>
		<%=Action_Str%>
	</td>
	<td><%=Action_Chinese_Str%></td>
	</Tr>
 <%
 	rs.movenext
 	next 
 %>
  </table>

</Tr>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->