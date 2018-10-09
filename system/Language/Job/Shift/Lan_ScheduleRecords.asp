<script language="javascript">
var strSearch="Search Line|查询线别"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchResult="Result|结果" 
var arrSearchResult=strSearchResult.split("|")
var strBrowse="Browse Schedule Records|浏览计划任务运行记录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strbyPart="Browse by Part|按型号" 
var arrbyPart=strbyPart.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strStatus="Status|状态" 
var arrStatus=strStatus.split("|")
var strLineName="Line Name|线别名称" 
var arrLineName=strLineName.split("|")
var strPerson="Person|创建人" 
var arrPerson=strPerson.split("|")
var strShiftType="Shift Type|班次类型" 
var arrShiftType=strShiftType.split("|")
var strRuntime="Runtime|运行时间" 
var arrRuntime=strRuntime.split("|")
var strRunResult="Run Result|运行结果" 
var arrRunResult=strRunResult.split("|")
var strErrorInfo="Error Info|错误信息" 
var arrErrorInfo=strErrorInfo.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchResult.innerText=arrSearchResult[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{inner_Person.innerText=arrPerson[<%=session("language")%>]}catch(e){}
try{inner_ShiftType.innerText=arrShiftType[<%=session("language")%>]}catch(e){}
try{inner_Runtime.innerText=arrRuntime[<%=session("language")%>]}catch(e){}
try{inner_RunResult.innerText=arrRunResult[<%=session("language")%>]}catch(e){}
try{inner_ErrorInfo.innerText=arrErrorInfo[<%=session("language")%>]}catch(e){}
}
</script>
