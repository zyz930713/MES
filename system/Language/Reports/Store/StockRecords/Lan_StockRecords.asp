<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchPrintTime="Print Time|打印时间" 
var arrSearchPrintTime=strSearchPrintTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")

var strBrowse="Browse Store Records|浏览入库记录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strID="ID|编号" 
var arrID=strID.split("|")
var strPrintTime="Print Time|打印时间" 
var arrPrintTime=strPrintTime.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchPrintTime.innerText=arrSearchPrintTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_ID.innerText=arrID[<%=session("language")%>]}catch(e){}
try{inner_PrintTime.innerText=arrPrintTime[<%=session("language")%>]}catch(e){}
}
</script>