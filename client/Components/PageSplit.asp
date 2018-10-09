<table width="100%" height="42" border="0" align="center">
<form name="FormPageRecord" method="post" action="<%=pagename&"?none="&pagepara%>">
  <tr> 
    <td height="16" colspan="2">
<span id="inner_PageTotal">总计</span>&nbsp;<font class="red"><%=rs.RecordCount%></font>&nbsp;<span id="inner_PageTotalRecords">记录在</span> <% =session("strpagenum") %>/<%=rs.pagecount%> <span id="inner_PageTotalPages">页</span>&nbsp; 
      <%if csng(session("strpagenum"))<>1 then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>"><span id="inner_PageFirstPage">首页</span></a>
	  <%else%>
	  <span id="inner_PageFirstPage" class="d_link">首页</span>
      <%end if%>&nbsp;
      <%if csng(session("strpagenum"))>1 then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=clng(session("strpagenum"))-1%>&pagesize_a=<%=pagesize_s%><%=pagepara%>"><span id="inner_PagePreviousPage">上一页</span></a> 
	  <%else%>
	  <span id="inner_PagePreviousPage" class="d_link">上一页</span>
      <%end if%>&nbsp;
      <%if csng(session("strpagenum"))<rs.pagecount then %>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=clng(session("strpagenum"))+1%>&pagesize_a=<%=pagesize_s%><%=pagepara%>"><span id="inner_PageNextPage">下一页</span></a> 
	  <%else%>
	  <span id="inner_PageNextPage" class="d_link">下一页</span>
      <%end if %>&nbsp;
      <%if csng(session("strpagenum"))<>rs.pagecount then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=rs.pagecount%>&pagesize_a=<%=pagesize_s%><%=pagepara%>"><span id="inner_PageLastPage">尾页</span></a> 
	  <%else%>
	  <span id="inner_PageLastPage" class="d_link">尾页</span>
      <%end if%>     </td>
  </tr>
  <tr> 
      <td width="74%" height="18"><span id="inner_PageGoto">转到</span>
	  <%if csng(session("pagespan"))-20>=0 then%>
	  <a href="<%=pagename%>?pagespan=<%=csng(session("pagespan"))-20%>&pagenum=<%=csng(session("pagespan"))-20+1%>&pagesize_a=<%=pagesize_s%><%=pagepara%>" title="Before 20 Pages"><</a>
	  <%end if%>&nbsp; 
		<% for i=session("pagespan")+1 to session("pagespan")+20
	  	if i<=rs.pagecount then
		  if csng(session("strpagenum"))<>i then%>
		  <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=i%>&pagesize_a=<%=pagesize_s%><%=pagepara%>" title="Page of <%=i%>"><%=i%></a>&nbsp; 
		  <%else%>
		  <%=i%>
		  <%
		  end if
		  else
		  exit for
		end if
	  next
	  if i<rs.pagecount then%>
	  <a href="<%=pagename%>?pagespan=<%=csng(session("pagespan"))+20%>&pagenum=<%=csng(session("pagespan"))+21%>&pagesize_a=<%=pagesize_s%><%=pagepara%>" title="Next 20 Pages">></a>
	  <%end if%>&nbsp;	  </td>
      <td width="26%" height="18">
          <div align="right">
          <select name="select" onchange="javascript:document.FormPageRecord.submit()">
            <option value="10" <%if pagesize_s=10 then%>selected<%end if%>>10</option>
            <option value="15" <%if pagesize_s=15 then%>selected<%end if%>>15</option>
            <option value="20" <%if pagesize_s=10 then%>selected<%end if%>>20</option>
            <option value="25" <%if pagesize_s=25 then%>selected<%end if%>>25</option>
            <option value="30" <%if pagesize_s=30 then%>selected<%end if%>>30</option>
            <option value="all" <%if pagesize_s="all" then%>selected<%end if%>>All</option>
          </select> 
          <span id="inner_PagePerPage">每页</span>
      </div></td>
  </tr>
</form>
</table>