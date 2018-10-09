<script language="javascript">
var strBrowse="My Form List|我的表单列表"
var arrBrowse=strBrowse.split("|")
var strBrowseUser="User|用户"
var arrBrowseUser=strBrowseUser.split("|")
var strAdd="Add a New Form|新增表单" 
var arrAdd=strAdd.split("|")
var strNO="No|编号" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strFormName="Form Name|表单名称" 
var arrFormName=strFormName.split("|")
var strFormStatus="Form Status|表单状态" 
var arrFormStatus=strFormStatus.split("|")
var strParam1="Param 1|参数 1" 
var arrParam1=strParam1.split("|")
var strParam2="Param 2|参数 2" 
var arrParam2=strParam2.split("|")
var strParam3="Param 3|参数 3" 
var arrParam3=strParam3.split("|")
var strApproveName="Approve Names|核准人" 
var arrApproveName=strApproveName.split("|")
var strActorCode="Transactor Code|处理人" 
var arrActorCode=strActorCode.split("|")
var strApplyTime="Apply Time|申请时间" 
var arrApplyTime=strApplyTime.split("|")
var strMailTimes="Mail Times|邮件次数" 
var arrMailTimes=strMailTimes.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){} 
try{inner_BrowseUser.innerText=arrBrowseUser[<%=session("language")%>]}catch(e){} 
try{inner_Add.innerText=arrAdd[<%=session("language")%>]}catch(e){} 
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){} 
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_FormName.innerText=arrFormName[<%=session("language")%>]}catch(e){}
try{inner_FormStatus.innerText=arrFormStatus[<%=session("language")%>]}catch(e){}
try{inner_Param1.innerText=arrParam1[<%=session("language")%>]}catch(e){}
try{inner_Param2.innerText=arrParam2[<%=session("language")%>]}catch(e){}
try{inner_Param3.innerText=arrParam3[<%=session("language")%>]}catch(e){}
try{inner_ApproveName.innerText=arrApproveName[<%=session("language")%>]}catch(e){}
try{inner_ActorCode.innerText=arrActorCode[<%=session("language")%>]}catch(e){}
try{inner_ApplyTime.innerText=arrApplyTime[<%=session("language")%>]}catch(e){}
try{inner_MailTimes.innerText=arrMailTimes[<%=session("language")%>]}catch(e){}
}
</script>