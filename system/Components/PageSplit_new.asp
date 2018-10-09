<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<%
totalSql = cstr(session("totalSql"))
set rsTotal=server.createobject("adodb.recordset")
rsTotal.open totalSql,conn,1,3
if not rsTotal.eof then
	recordCount=csng(rsTotal(0))	
	pageCount=fix(recordCount/recordsize)
	modValue=recordCount mod recordsize
	if modValue>0 then
		pageCount=pageCount+1
	end if
end if
rsTotal.close
set rsTotal=nothing
%>
<table width="100%" height="42" border="0" align="center">
<form name="FormPageRecord" method="post" action="<%=pagename&"?none="&pagepara%>">
  <tr> 
    <td height="16" colspan="2">
<span id="inner_PageTotal"></span>&nbsp;<font class="red"><%=recordCount%></font>&nbsp;<span id="inner_PageTotalRecords"></span> <% =pagenum %>/<%=pageCount%> <span id="inner_PageTotalPages"></span>&nbsp; 
      <%if cint(pagenum)<>1 then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=firstpage%>&recordsize=<%=recordsize%><%=pagepara%><%=pageorder%>"><span id="inner_PageFirstPage"></span></a>
	  <%else%>
	  <span id="inner_PageFirstPage2" class="d_link"></span>
      <%end if%>&nbsp;
      <%if cint(pagenum)>1 then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=cint(pagenum)-1%>&recordsize=<%=recordsize%><%=pagepara%><%=pageorder%>"><span id="inner_PagePreviousPage"></span></a> 
	  <%else%>
	  <span id="inner_PagePreviousPage2" class="d_link"></span>
      <%end if%>&nbsp;
      <%if cint(pagenum)<pageCount then %>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=cint(pagenum)+1%>&recordsize=<%=recordsize%><%=pagepara%><%=pageorder%>"><span id="inner_PageNextPage"></span></a> 
	  <%else%>
	  <span id="inner_PageNextPage2" class="d_link"></span>
      <%end if %>&nbsp;
      <%if cint(pagenum)<>pageCount then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=pageCount%>&recordsize=<%=recordsize%><%=pagepara%><%=pageorder%>"><span id="inner_PageLastPage"></span></a> 
	  <%else%>
	  <span id="inner_PageLastPage2" class="d_link"></span>
      <%end if%>     </td>
  </tr>
  <tr> 
      <td width="74%" height="18"><span id="inner_PageGoto"></span>
	  <%if cint(session("pagespan"))-20>=0 then%>
	  <a href="<%=pagename%>?pagespan=<%=cint(session("pagespan"))-20%>&pagenum=<%=cint(session("pagespan"))-20+1%>&recordsize=<%=recordsize%><%=pagepara%><%=pageorder%>" title="Before 20 Pages"><</a>
	  <%end if%>&nbsp; 
		<% for i=cint(session("pagespan"))+1 to cint(session("pagespan"))+20
	  	if i<=pageCount then
		  if cint(pagenum)<>i then%>
		  <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=i%>&recordsize=<%=recordsize%><%=pagepara%><%=pageorder%>" title="Page of <%=i%>"><%=i%></a>&nbsp; 
		  <%else%>
		  <%=i%>
		  <%
		  end if
		  else
		  exit for
		end if
	  next
	  if i<pageCount then%>
	  <a href="<%=pagename%>?pagespan=<%=cint(session("pagespan"))+20%>&pagenum=<%=cint(session("pagespan"))+21%>&recordsize=<%=recordsize%><%=pagepara%><%=pageorder%>" title="Next 20 Pages">></a>
	  <%end if%>&nbsp;	  </td>
      <td width="26%" height="18">
          <div align="right">
          <select name="recordsize" id="recordsize" onchange="javascript:document.FormPageRecord.submit()">
            <option value="10" <%if recordsize=10 then%>selected<%end if%>>10</option>
            <option value="15" <%if recordsize=15 then%>selected<%end if%>>15</option>
            <option value="20" <%if recordsize=20 then%>selected<%end if%>>20</option>
            <option value="25" <%if recordsize=25 then%>selected<%end if%>>25</option>
            <option value="30" <%if recordsize=30 then%>selected<%end if%>>30</option>
            <option value="1000" <%if recordsize=1000 then%>selected<%end if%>>1000</option>
          </select> 
          <span id="inner_PagePerPage"></span>
      </div></td>
  </tr>
</form>
</table>