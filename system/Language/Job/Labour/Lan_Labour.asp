<script language="javascript">
var strSearch="Search Line|查询线别"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|线别" 
var arrSearchLine=strSearchLine.split("|")
var strBrowse="Browse Labour|浏览劳动力" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strYearIndex="Year Index|年度" 
var arrYearIndex=strYearIndex.split("|")
var strWeekIndex="Week Index|周次" 
var arrWeekIndex=strWeekIndex.split("|")
var strLineName="Line Name|线别" 
var arrLineName=strLineName.split("|")
var strRealWork="Real Work|考勤工时" 
var arrRealWork=strRealWork.split("|")
var strTransferredWork="Transferred Work|转移工时" 
var arrTransferredWork=strTransferredWork.split("|")
var strIndirectWork="Indirect Work|非直接工时" 
var arrIndirectWork=strIndirectWork.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_YearIndex.innerText=arrYearIndex[<%=session("language")%>]}catch(e){}
try{inner_WeekIndex.innerText=arrWeekIndex[<%=session("language")%>]}catch(e){}
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{inner_RealWork.innerText=arrRealWork[<%=session("language")%>]}catch(e){}
try{inner_TransferredWork.innerText=arrTransferredWork[<%=session("language")%>]}catch(e){}
try{inner_IndirectWork.innerText=arrIndirectWork[<%=session("language")%>]}catch(e){}
}
</script>
