<script language="javascript">
var strBrowse="Customize Task|�Զ�������"
var arrBrowse=strBrowse.split("|")
var strFormType="Form Type|������" 
var arrFormType=strFormType.split("|")
var strParam1="Parameters 1|���� 1" 
var arrParam1=strParam1.split("|")
var strParam2="Parameters 2|���� 2" 
var arrParam2=strParam2.split("|")
var strParam3="Parameters 3|���� 3" 
var arrParam3=strParam3.split("|")
var strParam4="Parameters 4|���� 4" 
var arrParam4=strParam4.split("|")
var strJobInfo="Job Info|������Ϣ" 
var arrJobInfo=strJobInfo.split("|")
var strApproveFlow="Approve Flow|��׼����" 
var arrApproveFlow=strApproveFlow.split("|")
var strActPerson="Act Person|������" 
var arrActPerson=strActPerson.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){} 
try{inner_FormType.innerText=arrFormType[<%=session("language")%>]}catch(e){} 
try{inner_Param1.innerText=arrParam1[<%=session("language")%>]}catch(e){}
try{inner_Param2.innerText=arrParam2[<%=session("language")%>]}catch(e){}
try{inner_Param3.innerText=arrParam3[<%=session("language")%>]}catch(e){}
try{inner_Param4.innerText=arrParam4[<%=session("language")%>]}catch(e){}
try{inner_JobInfo.innerText=arrJobInfo[<%=session("language")%>]}catch(e){}
try{inner_ApproveFlow.innerText=arrApproveFlow[<%=session("language")%>]}catch(e){}
try{inner_ActPerson.innerText=arrActPerson[<%=session("language")%>]}catch(e){}
}
</script>