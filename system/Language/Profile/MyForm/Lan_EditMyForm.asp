<script language="javascript">
var strBrowse="Customize Task|自定义任务"
var arrBrowse=strBrowse.split("|")
var strFormType="Form Type|表单类型" 
var arrFormType=strFormType.split("|")
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