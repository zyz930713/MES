<script language="javascript">
var strPageTotal="Total|总计"
var arrPageTotal=strPageTotal.split("|")
var strPageTotalRecords="records on|记录在"
var arrPageTotalRecords=strPageTotalRecords.split("|")
var strPageTotalPages="pages|页"
var arrPageTotalPages=strPageTotalPages.split("|")
var strPageFirstPage="First Page|首页"
var arrPageFirstPage=strPageFirstPage.split("|")
var strPagePreviousPage="Previous Page|上一页"
var arrPagePreviousPage=strPagePreviousPage.split("|")
var strPageNextPage="Next Page|下一页"
var arrPageNextPage=strPageNextPage.split("|")
var strPageLastPage="Last Page|尾页"
var arrPageLastPage=strPageLastPage.split("|")
var strPageGoto="Go to|到"
var arrPageGoto=strPageGoto.split("|")
var strPagePerPage="records per page|条记录/每页"
var arrPagePerPage=strPagePerPage.split("|")
function language_page()
{
try{inner_PageTotal.innerText=arrPageTotal[<%=session("language")%>]}catch(e){} 
try{inner_PageTotalRecords.innerText=arrPageTotalRecords[<%=session("language")%>]}catch(e){} 
try{inner_PageTotalPages.innerText=arrPageTotalPages[<%=session("language")%>]}catch(e){} 
try{inner_PageFirstPage.innerText=arrPageFirstPage[<%=session("language")%>]}catch(e){} 
try{inner_PageFirstPage2.innerText=arrPageFirstPage[<%=session("language")%>]}catch(e){}
try{inner_PagePreviousPage.innerText=arrPagePreviousPage[<%=session("language")%>]}catch(e){}
try{inner_PagePreviousPage2.innerText=arrSPagePreviousPage[<%=session("language")%>]}catch(e){}
try{inner_PageNextPage.innerText=arrPageNextPage[<%=session("language")%>]}catch(e){}
try{inner_PageNextPage2.innerText=arrPageNextPage[<%=session("language")%>]}catch(e){}
try{inner_PageLastPage.innerText=arrPageLastPage[<%=session("language")%>]}catch(e){}
try{inner_PageLastPage2.innerText=arrPageLastPage[<%=session("language")%>]}catch(e){}
try{inner_PageGoto.innerText=arrPageGoto[<%=session("language")%>]}catch(e){}
try{inner_PagePerPage.innerText=arrPagePerPage[<%=session("language")%>]}catch(e){}
}
</script>
