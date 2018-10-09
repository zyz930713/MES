<form name="FormPageRecord" method="post" action="<%=pagename&"?none="&pagepara%>">
<table width="100%" height="42" border="0" align="center">
  <tr> 
    <td height="16" colspan="2">
Total <font class="red"><%=rsPR.RecordCount%></font> records on <% =session("strpagenum") %>/<%=rsPR.pagecount%> pages&nbsp; 
      <%
if clng(session("strpagenum"))<>1 then
%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=firsPRtpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>">First Page</a>
	  <%else%>
	  <span class="d_link">First Page</span>
      <%
end if
%>
      <%
if clng(session("strpagenum"))>1 then
%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=clng(session("strpagenum"))-1%>&pagesize_a=<%=pagesize_s%><%=pagepara%>">Previous Page</a> 
	  <%else%>
	  <span class="d_link">Previous Page</span>
      <%
end if
%>
      <%if clng(session("strpagenum"))<rsPR.pagecount then %>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=clng(session("strpagenum"))+1%>&pagesize_a=<%=pagesize_s%><%=pagepara%>">Next Page</a> 
	  <%else%>
	  <span class="d_link">Next Page</span>
      <%
end if 
%>
      <%
if clng(session("strpagenum"))<>rsPR.pagecount then
%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=rsPR.pagecount%>&pagesize_a=<%=pagesize_s%><%=pagepara%>">Last Page</a> 
	  <%else%>
	  <span class="d_link">Last Page</span>
      <%
end if
%>     </td>
  </tr>
  <tr> 
      <td width="58%" height="18"> <p align="left">Go to
	  <%if session("pagespan")-20>=0 then%>
	  <a href="<%=pagename%>?pagespan=<%=session("pagespan")-20%>&pagenum=<%=session("pagespan")-20+1%>&pagesize_a=<%=pagesize_s%><%=pagepara%>" title="Before 20 Pages"><</a>
	  <%end if%>&nbsp; 
		<% for i=session("pagespan")+1 to session("pagespan")+20
	  	if i<=rsPR.pagecount then
		  if clng(session("strpagenum"))<>i then%>
		  <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=i%>&pagesize_a=<%=pagesize_s%><%=pagepara%>" title="Page of <%=i%>"><%=i%></a>&nbsp; 
		  <%else%>
		  <%=i%>
		  <%
		  end if
		  else
		  exit for
		end if
	  next
	  if i<rsPR.pagecount then%>
	  <a href="<%=pagename%>?pagespan=<%=session("pagespan")+20%>&pagenum=<%=session("pagespan")+21%>&pagesize_a=<%=pagesize_s%><%=pagepara%>" title="Next 20 Pages">></a>
	  <%end if%>&nbsp; 
	  </td>
      <td width="42%" height="18"> <p align="right">
          <%=pagesize_s%> Records per page now
          <select name="select">
            <option value="10">10</option>
            <option value="15">15</option>
            <option value="20">20</option>
            <option value="25">25</option>
            <option value="30">30</option>
			<option value="all">All</option>
          </select>
          Records per page
          <img src="/Images/Reflash.gif" width="52" height="18" align="absmiddle" onClick="javascrip:document.FormPageRecord.submit()">
      </td>
  </tr>
</table>
</form>