<script language="javascript">
var strSearch="Search Records|������¼" 
var arrSearch=strSearch.split("|")
var strSearchPrintTime="Print Time|��ӡʱ��" 
var arrSearchPrintTime=strSearchPrintTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")

var strBrowse="Browse Store Records|�������¼" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strID="ID|���" 
var arrID=strID.split("|")
var strPrintTime="Print Time|��ӡʱ��" 
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