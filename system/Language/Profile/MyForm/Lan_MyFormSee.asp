<script language="javascript">
var strBrowse="Detail of System Form|系统表单明细"
var arrBrowse=strBrowse.split("|")
var strFormName="Form Name|表单名称" 
var arrFormName=strFormName.split("|")
var strApplicant="Applicant|申请人" 
var arrApplicant=strApplicant.split("|")
var strFormType="Form Type|表单类型" 
var arrFormType=strFormType.split("|")
var strFormStatus="Form Status|表单状态" 
var arrFormStatus=strFormStatus.split("|")
var strParam1="Parameters 1|参数 1" 
var arrParam1=strParam1.split("|")
var strParam2="Parameters 2|参数 2" 
var arrParam2=strParam2.split("|")
var strParam3="Parameters 3|参数 3" 
var arrParam3=strParam3.split("|")
var strParam4="Parameters 4|参数 4" 
var arrParam4=strParam4.split("|")
var strJobInfo="Job Info|工单信息" 
var arrJobInfo=strJobInfo.split("|")
var strApproveFlow="Approve Flow|核准流程" 
var arrApproveFlow=strApproveFlow.split("|")
var strActPerson="Act Person|处理人" 
var arrActPerson=strActPerson.split("|")
var strActionType="Action Type|处理方式" 
var arrActionType=strActionType.split("|")
var strDenyReason="Deny Reason to Approve|核准否决理由" 
var arrDenyReason=strDenyReason.split("|")
var strRejectReason="Reject Reason to Transact|拒绝处理理由" 
var arrRejectReason=strRejectReason.split("|")
var strButtonClose="Close|关闭" 
var arrButtonClose=strButtonClose.split("|")

function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){} 
try{inner_FormName.innerText=arrFormName[<%=session("language")%>]}catch(e){} 
try{inner_Applicant.innerText=arrApplicant[<%=session("language")%>]}catch(e){} 
try{inner_FormType.innerText=arrFormType[<%=session("language")%>]}catch(e){} 
try{inner_FormStatus.innerText=arrFormStatus[<%=session("language")%>]}catch(e){} 
try{inner_Param1.innerText=arrParam1[<%=session("language")%>]}catch(e){}
try{inner_Param2.innerText=arrParam2[<%=session("language")%>]}catch(e){}
try{inner_Param3.innerText=arrParam3[<%=session("language")%>]}catch(e){}
try{inner_Param4.innerText=arrParam4[<%=session("language")%>]}catch(e){}
try{inner_JobInfo.innerText=arrJobInfo[<%=session("language")%>]}catch(e){}
try{inner_ApproveFlow.innerText=arrApproveFlow[<%=session("language")%>]}catch(e){}
try{inner_ActPerson.innerText=arrActPerson[<%=session("language")%>]}catch(e){}
try{inner_ActionType.innerText=arrActionType[<%=session("language")%>]}catch(e){}
try{inner_DenyReason.innerText=arrDenyReason[<%=session("language")%>]}catch(e){}
try{inner_RejectReason.innerText=arrRejectReason[<%=session("language")%>]}catch(e){}
try{document.all.Close.value=arrButtonClose[<%=session("language")%>]}catch(e){}
}
</script>