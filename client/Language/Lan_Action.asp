<script language="javascript">
var strLine="Family|家族" 
var arrLine=strLine.split("|")
var strTitle="NCMR Ticket(Short-term Action) |异常处理单记录单 （立即措施） "
var arrTitle=strTitle.split("|")
var strFactory ="Factory|专业厂" 
var arrFactory=strFactory.split("|")
var strKeyWord="Key Word|关键字" 
var arrKeyWord=strKeyWord.split("|")
var strKeyWord1="Key Word|关键字" 
var arrKeyWord1=strKeyWord1.split("|")
var strKeyWord2="Key Word|关键字" 
var arrKeyWord2=strKeyWord2.split("|")
var strCreateTime="Create Time|申请时间" 
var arrCreateTime=strCreateTime.split("|")
var strCreator="Creator|申请人" 
var arrCreator=strCreator.split("|")
var strNum="NCMR NO.|异常单编号" 
var arrNum=strNum.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strModel="Model|型号" 
var arrModel=strModel.split("|")
var strStation="Station|工站" 
var arrStation=strStation.split("|")
var strProblem="Problem Description|问题描述" 
var arrProblem=strProblem.split("|")
var strUpload="Upload File|上传文件" 
var arrUpload=strUpload.split("|")
var strOperator="Action Person|本次处理人" 
var arrOperator=strOperator.split("|")
var strAlert="(Please input your Badge Number)|(请输入您的工号)" 
var arrAlert=strAlert.split("|")
var strNo ="NO.|编号" 
var arrNo=strNo.split("|")
var strAction1="Short-term Action|立即措施" 
var arrAction1=strAction1.split("|")
var strDepartment="Department|责任部门" 
var arrDepartment=strDepartment.split("|")
var strAction="Action|采取措施" 
var arrAction=strAction.split("|")
var strPerson ="Person|负责人" 
var arrPerson=strPerson.split("|")
var strDueDate="Due Date|指定完成日期" 
var arrDueDate=strDueDate.split("|")
var strDueDate1="Due Date|完成日期" 
var arrDueDate1=strDueDate1.split("|")
var strDueDate2="Due Date|完成日期" 
var arrDueDate2=strDueDate2.split("|")
var strDueDate3="Due Date|完成日期" 
var arrDueDate3=strDueDate3.split("|")
var strDueDate4="Due Date|完成日期" 
var arrDueDate4=strDueDate4.split("|")
var strRestart="Forward to:|转发" 
var arrRestart=strRestart.split("|") 
var strDotice="Team Memebers|NCMR 成员" 
var arrDotice=strDotice.split("|")
var strSupervisor="Supervisor|主管" 
var arrSupervisor=strSupervisor.split("|")
var strKeyPerson1="Defect Disposal Person|不合格处理相关人" 
var arrKeyPerson1=strKeyPerson1.split("|")
var strKeyPerson2="Root Cause Person|原因分析相关人" 
var arrKeyPerson2=strKeyPerson2.split("|")
var strKeyPerson3="Long-term Action Person|纠正措施相关人" 
var arrKeyPerson3=strKeyPerson3.split("|")
var strKeyPerson4="QA|QA相关人" 
var arrKeyPerson4=strKeyPerson4.split("|")
var strBackUp1="Backup|备选相关人" 
var arrBackUp1=strBackUp1.split("|")
var strBackUp2="Backup|备选相关人" 
var arrBackUp2=strBackUp2.split("|")
var strBackUp3="Backup|备选相关人" 
var arrBackUp3=strBackUp3.split("|")
var strBackUp4="Backup|备选相关人" 
var arrBackUp4=strBackUp4.split("|")
var strGreen="Green:Finished;|绿色：完成；" 
var arrGreen=strGreen.split("|")
var strRed="Red:Doing;|红色：进行中；" 
var arrRed=strRed.split("|")
var strBlue="Blue:Waiting;|蓝色：未进行；" 
var arrBlue=strBlue.split("|")
var strIsClosed="Is Closed|此单是否关闭" 
var arrIsClosed=strIsClosed.split("|")
var strCloseAlert="Please input reason if you want to close this NCMR|*如果您认为此NCMR单无须再继续进行下去，请填写原因后，点击该按钮关闭此单"
var arrCloseAlert=strCloseAlert.split("|")
var strRisk="Risk Assessment|风险评估"
var arrRisk=strRisk.split("|")
var strDefect1="Defect Disposal|不合格处理"
var arrDefect1=strDefect1.split("|")
var strDefect2="Reject|批退"
var arrDefect2=strDefect2.split("|")
var strDefect3="Scrap|报废"
var arrDefect3=strDefect3.split("|")
var strDefect4="Rework /Sorting|返修/选别"
var arrDefect4=strDefect4.split("|")
var strDefect5="Deviation|变异接收"
var arrDefect5=strDefect5.split("|")
var strDefect6="Others|其他"
var arrDefect6=strDefect6.split("|")
var strRisk="Risk Assessment|风险评估"
var arrRisk=strRisk.split("|")
var strCause1="Root Cause|原因分析"
var arrCause1=strCause1.split("|")
var strCorrect1="Long-term Action|纠正措施"
var arrCorrect1=strCorrect1.split("|")
var strRejectHistory="Rejected History|拒绝历史"
var arrRejectHistory=strRejectHistory.split("|")
var strRejectStep="Rejected Step|拒绝环节"
var arrRejectStep=strRejectStep.split("|")
var strRejectCause="Reason|拒绝原因"
var arrRejectCause=strRejectCause.split("|")
var strRejectTime="Rejected Time|拒绝时间"
var arrRejectTime=strRejectTime.split("|")
var strRejectPerson="Rejected Person|拒绝人"
var arrRejectPerson=strRejectPerson.split("|")
var strRecords6="None|无"
var arrRecords6=strRecords6.split("|")

var strOwner="Owner|负责人"
var arrOwner=strOwner.split("|")
function language()
{
try{inner_Title.innerText=arrTitle[<%=session("language")%>]}catch(e){} 
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){} 
try{inner_KeyWord.innerText=arrKeyWord[<%=session("language")%>]}catch(e){}
try{inner_KeyWord1.innerText=arrKeyWord1[<%=session("language")%>]}catch(e){}
try{inner_KeyWord2.innerText=arrKeyWord2[<%=session("language")%>]}catch(e){}
try{inner_Green.innerText=arrGreen[<%=session("language")%>]}catch(e){}
try{inner_Red.innerText=arrRed[<%=session("language")%>]}catch(e){}
try{inner_Blue.innerText=arrBlue[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_CreateTime.innerText=arrCreateTime[<%=session("language")%>]}catch(e){}
try{inner_Creator.innerText=arrCreator[<%=session("language")%>]}catch(e){}
try{inner_Num.innerText=arrNum[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_Model.innerText=arrModel[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_Problem.innerText=arrProblem[<%=session("language")%>]}catch(e){}
try{inner_Upload.innerText=arrUpload[<%=session("language")%>]}catch(e){}
try{inner_Operator.innerText=arrOperator[<%=session("language")%>]}catch(e){}
try{inner_Alert.innerText=arrAlert[<%=session("language")%>]}catch(e){}
try{inner_IsClosed.innerText=arrIsClosed[<%=session("language")%>]}catch(e){}
try{inner_CloseAlert.innerText=arrCloseAlert[<%=session("language")%>]}catch(e){}
try{inner_IsSkip.innerText=arrIsSkip[<%=session("language")%>]}catch(e){}
try{inner_IsSkip1.innerText=arrIsSkip1[<%=session("language")%>]}catch(e){}
try{inner_IsSkip2.innerText=arrIsSkip2[<%=session("language")%>]}catch(e){}
try{inner_IsSkip3.innerText=arrIsSkip3[<%=session("language")%>]}catch(e){}
try{inner_No.innerText=arrNo[<%=session("language")%>]}catch(e){}
try{inner_No1.innerText=arrNo[<%=session("language")%>]}catch(e){}
try{inner_Action1.innerText=arrAction1[<%=session("language")%>]}catch(e){}
try{inner_Action2.innerText=arrAction1[<%=session("language")%>]}catch(e){}
try{inner_DueDate1.innerText=arrDueDate1[<%=session("language")%>]}catch(e){}
try{inner_DueDate2.innerText=arrDueDate2[<%=session("language")%>]}catch(e){}
try{inner_DueDate3.innerText=arrDueDate3[<%=session("language")%>]}catch(e){}
try{inner_DueDate4.innerText=arrDueDate4[<%=session("language")%>]}catch(e){}
try{inner_DueDate.innerText=arrDueDate[<%=session("language")%>]}catch(e){}
try{inner_Department.innerText=arrDepartment[<%=session("language")%>]}catch(e){}
try{inner_Department2.innerText=arrDepartment[<%=session("language")%>]}catch(e){}
try{inner_Person.innerText=arrPerson[<%=session("language")%>]}catch(e){}
try{inner_Person2.innerText=arrPerson[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Supervisor.innerText=arrSupervisor[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson1.innerText=arrKeyPerson1[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson2.innerText=arrKeyPerson2[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson3.innerText=arrKeyPerson3[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson4.innerText=arrKeyPerson4[<%=session("language")%>]}catch(e){}
try{inner_BackUp1.innerText=arrBackUp1[<%=session("language")%>]}catch(e){}
try{inner_BackUp2.innerText=arrBackUp2[<%=session("language")%>]}catch(e){}
try{inner_BackUp3.innerText=arrBackUp3[<%=session("language")%>]}catch(e){}
try{inner_BackUp4.innerText=arrBackUp4[<%=session("language")%>]}catch(e){}
try{inner_Restart.innerText=arrRestart[<%=session("language")%>]}catch(e){}
try{inner_Dotice.innerText=arrDotice[<%=session("language")%>]}catch(e){}
try{inner_CloseAlert.innerText=arrCloseAlert[<%=session("language")%>]}catch(e){}
try{inner_Risk.innerText=arrRisk[<%=session("language")%>]}catch(e){}
try{inner_Defect1.innerText=arrDefect1[<%=session("language")%>]}catch(e){}
try{inner_Defect2.innerText=arrDefect2[<%=session("language")%>]}catch(e){}
try{inner_Defect3.innerText=arrDefect3[<%=session("language")%>]}catch(e){}
try{inner_Defect4.innerText=arrDefect4[<%=session("language")%>]}catch(e){}
try{inner_Defect5.innerText=arrDefect5[<%=session("language")%>]}catch(e){}
try{inner_Defect6.innerText=arrDefect6[<%=session("language")%>]}catch(e){}
try{inner_Risk.innerText=arrRisk[<%=session("language")%>]}catch(e){}
try{inner_Cause1.innerText=arrCause1[<%=session("language")%>]}catch(e){} 
try{inner_Correct1.innerText=arrCorrect1[<%=session("language")%>]}catch(e){} 
try{inner_RejectHistory.innerText=arrRejectHistory[<%=session("language")%>]}catch(e){} 
try{inner_RejectStep.innerText=arrRejectStep[<%=session("language")%>]}catch(e){} 
try{inner_RejectCause.innerText=arrRejectCause[<%=session("language")%>]}catch(e){} 
try{inner_RejectTime.innerText=arrRejectTime[<%=session("language")%>]}catch(e){} 
try{inner_RejectPerson.innerText=arrRejectPerson[<%=session("language")%>]}catch(e){} 
try{inner_Records6.innerText=arrRecords6[<%=session("language")%>]}catch(e){}
try{inner_Owner.innerText=arrOwner[<%=session("language")%>]}catch(e){}
}
</script>