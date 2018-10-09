<script language="javascript">
var strBrowse="Browse Shift Status|浏览班次状态" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strAdd="Add New DB Job|新增数据库计划" 
var arrAdd=strAdd.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strStatus="Status|状态" 
var arrStatus=strStatus.split("|")
var strJobID="Job ID|任务编号" 
var arrJobID=strJobID.split("|")
var strWhat="What to Do|执行内容" 
var arrWhat=strWhat.split("|")
var strRunTime="Run Time|计划时间" 
var arrRunTime=strRunTime.split("|")
var strTotalTime="Total Times|执行次数" 
var arrTotalTime=strTotalTime.split("|")
var strInterval="Interval|间隔时间" 
var arrInterval=strInterval.split("|")
var strFailures="Failures|失败次数" 
var arrFailures=strFailures.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_Add.innerText=arrAdd[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_JobID.innerText=arrJobID[<%=session("language")%>]}catch(e){}
try{inner_What.innerText=arrWhat[<%=session("language")%>]}catch(e){}
try{inner_RunTime.innerText=arrRunTime[<%=session("language")%>]}catch(e){}
try{inner_TotalTime.innerText=arrTotalTime[<%=session("language")%>]}catch(e){}
try{inner_Interval.innerText=arrInterval[<%=session("language")%>]}catch(e){}
try{inner_Failures.innerText=arrFailures[<%=session("language")%>]}catch(e){}
}
</script>
