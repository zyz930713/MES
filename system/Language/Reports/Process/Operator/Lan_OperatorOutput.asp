<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchLine="Line|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strSearchProgressTypeSelect="-- Select Type --|-- 选择类型 --" 
var arrSearchProgressTypeSelect=strSearchProgressTypeSelect.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchCreateTime.innerText=arrSearchCreateTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_LineLost.innerText=arrLineLost[<%=session("language")%>]}catch(e){}
try{inner_InputTime.innerText=arrInputTime[<%=session("language")%>]}catch(e){}
try{inner_LastTime.innerText=arrLastTime[<%=session("language")%>]}catch(e){}
}
</script>
