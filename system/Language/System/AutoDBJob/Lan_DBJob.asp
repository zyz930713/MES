<script language="javascript">
var strBrowse="Browse Shift Status|������״̬" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strAdd="Add New DB Job|�������ݿ�ƻ�" 
var arrAdd=strAdd.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strStatus="Status|״̬" 
var arrStatus=strStatus.split("|")
var strJobID="Job ID|������" 
var arrJobID=strJobID.split("|")
var strWhat="What to Do|ִ������" 
var arrWhat=strWhat.split("|")
var strRunTime="Run Time|�ƻ�ʱ��" 
var arrRunTime=strRunTime.split("|")
var strTotalTime="Total Times|ִ�д���" 
var arrTotalTime=strTotalTime.split("|")
var strInterval="Interval|���ʱ��" 
var arrInterval=strInterval.split("|")
var strFailures="Failures|ʧ�ܴ���" 
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
