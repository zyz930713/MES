<script language="javascript">
var strBrowse="Transact Form|�����"
var arrBrowse=strBrowse.split("|")
var strFormName="Form Name|������" 
var arrFormName=strFormName.split("|")
var strApplicant="Applicant|������" 
var arrApplicant=strApplicant.split("|")
var strFormType="Form Type|������" 
var arrFormType=strFormType.split("|")
var strFormStatus="Form Status|��״̬" 
var arrFormStatus=strFormStatus.split("|")
var strParam1="Parameters 1|���� 1" 
var arrParam1=strParam1.split("|")
var strParam2="Parameters 2|���� 2" 
var arrParam2=strParam2.split("|")
var strParam3="Parameters 3|���� 3" 
var arrParam3=strParam3.split("|")
var strJobInfo="Job Info|������Ϣ" 
var arrJobInfo=strJobInfo.split("|")
var strApproveFlow="Approve Flow|��׼����" 
var arrApproveFlow=strApproveFlow.split("|")
var strActPerson="Act Person|������" 
var arrActPerson=strActPerson.split("|")
var strActionType="Action Type|����ʽ" 
var arrActionType=strActionType.split("|")
var strActionAccept="Accept|����" 
var arrActionAccept=strActionAccept.split("|")
var strActionReject="Reject|�ܾ�" 
var arrActionReject=strActionReject.split("|")
var strRejectReason="Reject Reason|�ܾ�����" 
var arrRejectReason=strRejectReason.split("|")
var strButtonTransact="Transact|����" 
var arrButtonTransact=strButtonTransact.split("|")
var strButtonReset="Reset|����" 
var arrButtonReset=strButtonReset.split("|")

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
try{inner_JobInfo.innerText=arrJobInfo[<%=session("language")%>]}catch(e){}
try{inner_ApproveFlow.innerText=arrApproveFlow[<%=session("language")%>]}catch(e){}
try{inner_ActPerson.innerText=arrActPerson[<%=session("language")%>]}catch(e){}
try{inner_ActionType.innerText=arrActionType[<%=session("language")%>]}catch(e){}
try{inner_ActionAccept.innerText=arrActionAccept[<%=session("language")%>]}catch(e){}
try{inner_ActionReject.innerText=arrActionReject[<%=session("language")%>]}catch(e){}
try{inner_RejectReason.innerText=arrRejectReason[<%=session("language")%>]}catch(e){}
try{document.all.Transact.value=arrButtonTransact[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrButtonReset[<%=session("language")%>]}catch(e){}
}
</script>