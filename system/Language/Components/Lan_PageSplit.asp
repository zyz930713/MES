<script language="javascript">
var strPageTotal="Total|�ܼ�"
var arrPageTotal=strPageTotal.split("|")
var strPageTotalRecords="records on|��¼��"
var arrPageTotalRecords=strPageTotalRecords.split("|")
var strPageTotalPages="pages|ҳ"
var arrPageTotalPages=strPageTotalPages.split("|")
var strPageFirstPage="First Page|��ҳ"
var arrPageFirstPage=strPageFirstPage.split("|")
var strPagePreviousPage="Previous Page|��һҳ"
var arrPagePreviousPage=strPagePreviousPage.split("|")
var strPageNextPage="Next Page|��һҳ"
var arrPageNextPage=strPageNextPage.split("|")
var strPageLastPage="Last Page|βҳ"
var arrPageLastPage=strPageLastPage.split("|")
var strPageGoto="Go to|��"
var arrPageGoto=strPageGoto.split("|")
var strPagePerPage="records per page|����¼/ÿҳ"
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
