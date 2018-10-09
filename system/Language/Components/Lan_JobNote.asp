<script language="javascript">
var strJobisopened="Job is opened|工单开放"
var arrJobisopened=strJobisopened.split("|")
var strJobispaused="Job is paused|工单暂停"
var arrJobispaused=strJobispaused.split("|")
var strJobisclosed="Job is closed|工单关闭"
var arrJobisclosed=strJobisclosed.split("|")
var strJobisshiftouted="Job is shift-out|工单停线"
var arrJobisshiftouted=strJobisshiftouted.split("|")
var strJobislocked="Job is locked|工单锁定"
var arrJobislocked=strJobislocked.split("|")
var strJobisaborted="Job is aborted|工单废除"
var arrJobisaborted=strJobisaborted.split("|")
var strStartjob="Start job|启动工单"
var arrStartjob=strStartjob.split("|")
var strPausejob="Pause job|暂停工单"
var arrPausejob=strPausejob.split("|")
var strAbortjob="Abort job|废除工单"
var arrAbortjob=strAbortjob.split("|")
var strRepeatjob="Repeat job|重做工单"
var arrRepeatjob=strRepeatjob.split("|")
var strHistory="History|控制记录"
var arrHistory=strHistory.split("|")
var strRepeated="Repeated|重做记录"
var arrRepeated=strRepeated.split("|")
var strGreen="Green|绿色"
var arrGreen=strGreen.split("|")
var strFinishedstation="Finished station|工站完成"
var arrFinishedstation=strFinishedstation.split("|")
var strRed="Green|红色"
var arrRed=strRed.split("|")
var strCurrentstation="Current station|当前工站"
var arrCurrentstation=strCurrentstation.split("|")
var strBlue="Green|蓝色"
var arrBlue=strBlue.split("|")
var strWaitingstation="Waiting station|未做工站"
var arrWaitingstation=strWaitingstation.split("|")
var strItalic="Italic|斜体"
var arrItalic=strItalic.split("|")
var strOptionalstations="Optional stations|可选站"
var arrOptionalstations=strOptionalstations.split("|")
var strConjuctive="[]|[]"
var arrConjuctive=strConjuctive.split("|")
var strConjuctivestations="Conjuctive stations|组合站"
var arrConjuctivestations=strConjuctivestations.split("|")
var strPractised="Practised|组合站"
var arrPractised=strPractised.split("|")
function language_jobnote()
{
try{inner_Jobisopened.innerText=arrJobisopened[<%=session("language")%>]}catch(e){} 
try{inner_Jobispaused.innerText=arrJobispaused[<%=session("language")%>]}catch(e){} 
try{inner_Jobisclosed.innerText=arrJobisclosed[<%=session("language")%>]}catch(e){} 
try{inner_Jobisshiftouted.innerText=arrJobisshiftouted[<%=session("language")%>]}catch(e){} 
try{inner_Jobislocked.innerText=arrJobislocked[<%=session("language")%>]}catch(e){} 
try{inner_Jobisaborted.innerText=arrJobisaborted[<%=session("language")%>]}catch(e){}
try{inner_Startjob.innerText=arrStartjob[<%=session("language")%>]}catch(e){}
try{inner_Pausejob.innerText=arrPausejob[<%=session("language")%>]}catch(e){}
try{inner_Abortjob.innerText=arrAbortjob[<%=session("language")%>]}catch(e){}
try{inner_Repeatjob.innerText=arrRepeatjob[<%=session("language")%>]}catch(e){}
try{inner_History.innerText=arrHistory[<%=session("language")%>]}catch(e){}
try{inner_Repeated.innerText=arrRepeated[<%=session("language")%>]}catch(e){}
try{inner_Green.innerText=arrGreen[<%=session("language")%>]}catch(e){}
try{inner_Finishedstation.innerText=arrFinishedstation[<%=session("language")%>]}catch(e){}
try{inner_Red.innerText=arrRed[<%=session("language")%>]}catch(e){}
try{inner_Currentstation.innerText=arrCurrentstation[<%=session("language")%>]}catch(e){}
try{inner_Blue.innerText=arrBlue[<%=session("language")%>]}catch(e){}
try{inner_Waitingstation.innerText=arrWaitingstation[<%=session("language")%>]}catch(e){}
try{inner_Italic.innerText=arrItalic[<%=session("language")%>]}catch(e){}
try{inner_Optionalstations.innerText=arrOptionalstations[<%=session("language")%>]}catch(e){}
try{inner_Conjuctivestations.innerText=arrConjuctivestations[<%=session("language")%>]}catch(e){}
try{inner_Finishedstation2.innerText=arrFinishedstation[<%=session("language")%>]}catch(e){}
try{inner_Currentstation2.innerText=arrCurrentstation[<%=session("language")%>]}catch(e){}
try{inner_Practised.innerText=arrPractised[<%=session("language")%>]}catch(e){}
}
</script>
