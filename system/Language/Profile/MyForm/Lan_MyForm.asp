<script language="javascript">
var strBrowse="My Form List|�ҵı��б�"
var arrBrowse=strBrowse.split("|")
var strBrowseUser="User|�û�"
var arrBrowseUser=strBrowseUser.split("|")
var strAdd="Add a New Form|������" 
var arrAdd=strAdd.split("|")
var strNO="No|���" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strFormName="Form Name|������" 
var arrFormName=strFormName.split("|")
var strFormStatus="Form Status|��״̬" 
var arrFormStatus=strFormStatus.split("|")
var strParam1="Param 1|���� 1" 
var arrParam1=strParam1.split("|")
var strParam2="Param 2|���� 2" 
var arrParam2=strParam2.split("|")
var strParam3="Param 3|���� 3" 
var arrParam3=strParam3.split("|")
var strApproveName="Approve Names|��׼��" 
var arrApproveName=strApproveName.split("|")
var strActorCode="Transactor Code|������" 
var arrActorCode=strActorCode.split("|")
var strApplyTime="Apply Time|����ʱ��" 
var arrApplyTime=strApplyTime.split("|")
var strMailTimes="Mail Times|�ʼ�����" 
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