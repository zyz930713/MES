<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<table width="100%" height="42" border="0" align="center">
<form name="FormPageRecord" method="post" action="<%=pagename&"?none="&pagepara%>">
  <tr> 
    <td height="16" colspan="2">
<span id="inner_PageTotal"></span>&nbsp;<font class="red"><%=rsPR.RecordCount%></font>&nbsp;<span id="inner_PageTotalRecords"></span> <% =session("strpagenum") %>/<%=rsPR.pagecount%> <span id="inner_PageTotalPages"></span>&nbsp; 
      <%if clng(session("strpagenum"))<>1 then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=firsPRtpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>"><span id="inner_PageFirsPRtPage"></span></a>
	  <%else%>
	  <span id="inner_PageFirsPRtPage2" class="d_link"></span>
      <%end if%>&nbsp;
      <%if clng(session("strpagenum"))>1 then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=clng(session("strpagenum"))-1%>&pagesize_a=<%=pagesize_s%><%=pagepara%>"><span id="inner_PagePreviousPage"></span></a> 
	  <%else%>
	  <span id="inner_PagePreviousPage2" class="d_link"></span>
      <%end if%>&nbsp;
      <%if clng(session("strpagenum"))<rsPR.pagecount then %>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=clng(session("strpagenum"))+1%>&pagesize_a=<%=pagesize_s%><%=pagepara%>"><span id="inner_PageNextPage"></span></a> 
	  <%else%>
	  <span id="inner_PageNextPage2" class="d_link"></span>
      <%end if %>&nbsp;
      <%if clng(session("strpagenum"))<>rsPR.pagecount then%>
      <a href="<%=pagename%>?pagespan=<%=session("pagespan")%>&pagenum=<%=rsPR.pagecount%>&pagesize_a=<%=pagesize_s%><%=pagepara%>"><span id="inner_PageLastPage"></span></a> 
	  <%else%>
	  <span id="inner_PageLastPage2" class="d_link"></span>
      <%end if%>     </td>
  </tr>
  <tr> 
      <td width="74%" height="18"><span id="inner_PageGoto"></span>
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
          <span id="inner_PagePerPage"></span>
      </div></td>
  </tr>
</form>
</table>