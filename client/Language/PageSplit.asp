<table width="100%" height="42" border="0" align="center">
<form name="FormPageRecord" method="post" action="<%=pagename&"?"&pagepara%>">
  <tr> 
    <td height="16" colspan="2">
<span id="inner_PageTotal"></span>&nbsp;<font class="red"><%=rs.RecordCount%></font>&nbsp;<span id="inner_PageTotalRecords"></span> <% =session("strpagenum") %>/<%=rs.pagecount%> <span id="inner_PageTotalPages"></span>&nbsp; 
      <%if clng(session("strpagenum"))<>1 then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=firstpage%>&recordsize=<%=recordsize%><%=pagepara%>"><span id="inner_PageFirstPage"></span></a>
	  <%else%>
	  <span id="inner_PageFirstPage2" class="d_link"></span>
      <%end if%>&nbsp;
      <%if clng(session("strpagenum"))>1 then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=clng(session("strpagenum"))-1%>&recordsize=<%=recordsize%><%=pagepara%>"><span id="inner_PagePreviousPage"></span></a> 
	  <%else%>
	  <span id="inner_PagePreviousPage2" class="d_link"></span>
      <%end if%>&nbsp;
      <%if clng(session("strpagenum"))<rs.pagecount then %>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=clng(session("strpagenum"))+1%>&recordsize=<%=recordsize%><%=pagepara%>"><span id="inner_PageNextPage"></span></a> 
	  <%else%>
	  <span id="inner_PageNextPage2" class="d_link"></span>
      <%end if %>&nbsp;
      <%if clng(session("strpagenum"))<>rs.pagecount then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=rs.pagecount%>&recordsize=<%=recordsize%><%=pagepara%>"><span id="inner_PageLastPage"></span></a> 
	  <%else%>
	  <span id="inner_PageLastPage2" class="d_link"></span>
      <%end if%>     </td>
  </tr>
  <tr> 
      <td width="74%" height="18"><span id="inner_PageGoto"></span>
	  <%if session("pagespan")-20>=0 then%>
	  <a href="<%=pagename%>?pagespan=<%=session("pagespan")-20%>&pagenum=<%=session("pagespan")-20+1%>&recordsize=<%=recordsize%><%=pagepara%>" title="Before 20 Pages"><</a>
	  <%end if%>&nbsp; 
		<% for i=session("pagespan")+1 to session("pagespan")+20
	  	if i<=rs.pagecount then
		  if clng(session("strpagenum"))<>i then%>
		  <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=i%>&recordsize=<%=recordsize%><%=pagepara%>" title="Page of <%=i%>"><%=i%></a>&nbsp; 
		  <%else%>
		  <%=i%>
		  <%
		  end if
		  else
		  exit for
		end if
	  next
	  if i<rs.pagecount then%>
	  <a href="<%=pagename%>?pagespan=<%=session("pagespan")+20%>&pagenum=<%=session("pagespan")+21%>&recordsize=<%=recordsize%><%=pagepara%>" title="Next 20 Pages">></a>
	  <%end if%>&nbsp;	  </td>
      <td width="26%" height="18">
          <div align="right">
          <select name="recordsize" id="recordsize" onchange="javascript:document.FormPageRecord.submit()">
            <option value="10" <%if recordsize=10 then%>selected<%end if%>>10</option>
            <option value="15" <%if recordsize=15 then%>selected<%end if%>>15</option>
            <option value="20" <%if recordsize=20 then%>selected<%end if%>>20</option>
            <option value="25" <%if recordsize=25 then%>selected<%end if%>>25</option>
            <option value="30" <%if recordsize=30 then%>selected<%end if%>>30</option>
            <option value="all" <%if recordsize="all" then%>selected<%end if%>>All</option>
          </select> 
          <span id="inner_PagePerPage"></span>
      </div></td>
  </tr>
</form>
</table>