<script language="javascript">
var strLine="Family|家族" 
var arrLine=strLine.split("|")
var strTitle="NCMR Ticket ( Supervisor Confirm )|异常处理单记录单 （主管确认）"
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
var strProblem="Problem Description|问题描述" 
var arrProblem=strProblem.split("|")
var strUpload="Upload File|上传文件" 
var arrUpload=strUpload.split("|")
var strAction1="Short-term Action|立即措施" 
var arrAction1=strAction1.split("|")
var strAction2="Short-term Action|立即措施" 
var arrAction2=strAction2.split("|")
var strAction3="Short-term Action|立即措施" 
var arrAction3=strAction3.split("|")
var strAction4="Short-term Action|立即措施" 
var arrAction4=strAction4.split("|")
var strPerson1 ="Person|处理人" 
var arrPerson1=strPerson1.split("|")
var strPerson2 ="Person|处理人" 
var arrPerson2=strPerson2.split("|")
var strPerson3 ="Person|处理人" 
var arrPerson3=strPerson3.split("|")
var strPerson4 ="Person|处理人" 
var arrPerson4=strPerson4.split("|")
var strPerson5 ="Person|处理人" 
var arrPerson5=strPerson5.split("|")
var strPerson6 ="Person|负责人" 
var arrPerson6=strPerson6.split("|")
var strPerson7 ="Person|负责人" 
var arrPerson7=strPerson7.split("|")
var strDefect1="Defect Disposal|不合格处理"
var arrDefect1=strDefect1.split("|")
var strDefect2="Defect Disposal|不合格处理"
var arrDefect2=strDefect2.split("|")
var strDefect3="Defect Disposal|不合格处理"
var arrDefect3=strDefect3.split("|")
var strDefect4="Defect Disposal|不合格处理"
var arrDefect4=strDefect4.split("|")

var strStep1="Step|处理步骤" 
var arrStep1=strStep1.split("|")
var strStep2="Step|处理步骤" 
var arrStep2=strStep2.split("|")
var strStep3="Step|处理步骤" 
var arrStep3=strStep3.split("|")
var strStep4="Step|处理步骤" 
var arrStep4=strStep4.split("|")
var strStep5="Step|处理步骤" 
var arrStep5=strStep5.split("|")
var strTime1="Finish Date|处理日期" 
var arrTime1=strTime1.split("|")
var strTime2="Finish Date |处理日期" 
var arrTime2=strTime2.split("|")
var strTime3="Finish Date |处理日期" 
var arrTime3=strTime3.split("|")
var strTime4="Finish Date |处理日期" 
var arrTime4=strTime4.split("|")
var strTime5="Finish Date |处理日期" 
var arrTime5=strTime5.split("|")
var strRecords1="No Records|没有记录" 
var arrRecords1=strRecords1.split("|")
var strRecords2="None|无" 
var arrRecords2=strRecords2.split("|")
var strRecords3="None|无"
var arrRecords3=strRecords3.split("|")
var strRecords4="None|无"
var arrRecords4=strRecords4.split("|")
var strRecords5="None|无"
var arrRecords5=strRecords5.split("|")
var strRecords6="None|无"
var arrRecords6=strRecords6.split("|")

var strDepartment="Department|责任部门" 
var arrDepartment=strDepartment.split("|")
var strDepartment1="Department|责任部门" 
var arrDepartment1=strDepartment1.split("|")
var strDueDate="Due Date|指定完成日期" 
var arrDueDate=strDueDate.split("|")
var strDueDate1="Due Date|指定完成日期" 
var arrDueDate1=strDueDate1.split("|")
var strCancel="The Ticket was cancaled ,Reason as follow:|该单已被取消，原因如下：" 
var arrCancel=strCancel.split("|")
var strDefect="Defect Disposal|不合格处理"
var arrDefect=strDefect.split("|")
var strDefect2="Defect Disposal|不合格处理"
var arrDefect2=strDefect2.split("|")
var strDefect3="Defect Disposal|不合格处理"
var arrDefect3=strDefect3.split("|")
var strDefect4="Defect Disposal|不合格处理"
var arrDefect4=strDefect4.split("|")
var strRisk="Risk Assessment|风险评估"
var arrRisk=strRisk.split("|")
var strRemark="Remark|备注"
var arrRemark=strRemark.split("|")
var strRemark1="Remark|备注"
var arrRemark1=strRemark1.split("|")
var strRemark2="Remark|备注"
var arrRemark2=strRemark2.split("|")
var strRemark3="Remark|备注"
var arrRemark3=strRemark3.split("|")
var strOperator="Action Person|本次处理人" 
var arrOperator=strOperator.split("|")
var strAlert="(Please input your Badge Number)|(请输入您的工号)" 
var arrAlert=strAlert.split("|")
var strCause1="Root Cause|原因分析"
var arrCause1=strCause1.split("|")
var strCause2="Root Cause|原因分析"
var arrCause2=strCause2.split("|")
var strCause3="Root Cause|原因分析"
var arrCause3=strCause3.split("|")
var strCause4="Root Cause|原因分析"
var arrCause4=strCause4.split("|")
var strCause5="Root Cause|原因分析"
var arrCause5=strCause5.split("|")
var strAttachment="Attachment|附件"
var arrAttachment=strAttachment.split("|")
var strAttachment1="Attachment|附件"
var arrAttachment1=strAttachment1.split("|")
var strAttachment2="Attachment|附件"
var arrAttachment2=strAttachment2.split("|")

var strCorrect1="Long-term Action|纠正措施"
var arrCorrect1=strCorrect1.split("|")
var strCorrect2="Long-term Action|纠正措施"
var arrCorrect2=strCorrect2.split("|")
var strCorrect3="Long-term Action|纠正措施"
var arrCorrect3=strCorrect3.split("|")
var strCorrect4="Long-term Action|纠正措施"
var arrCorrect4=strCorrect4.split("|")
var strCorrect5="Long-term Action|纠正措施"
var arrCorrect5=strCorrect5.split("|")
var strCorrect6="Skip Long-term Action|纠正措施被跳过"
var arrCorrect6=strCorrect6.split("|")

var strQA1="Supervisor Confirm|主管确认"
var arrQA1=strQA1.split("|")
var strQA2="Supervisor Confirm|主管确认"
var arrQA2=strQA2.split("|")
var strQA3="Supervisor Confirm|主管确认"
var arrQA3=strQA3.split("|")
var strQA4="Supervisor Confirm|主管确认"
var arrQA4=strQA4.split("|")
var strQA5="Supervisor Confirm Passed|主管确认通过"
var arrQA5=strQA5.split("|")
var strQA6="Supervisor Confirm Failed|主管确认未通过"
var arrQA6=strQA6.split("|")
var strQA7="Supervisor Confirm|主管确认"
var arrQA7=strQA7.split("|")

var strFollow1="Follow Up|措施追踪"
var arrFollow1=strFollow1.split("|")
var strFollow2="Follow Up|措施追踪"
var arrFollow2=strFollow2.split("|")
var strFollow3="Follow Up|措施追踪"
var arrFollow3=strFollow3.split("|")
var strFollow4="Follow Up finished|措施追踪已完成"
var arrFollow4=strFollow4.split("|")
var strFollow5="Waiting Follow Up|措施追踪未完成"
var arrFollow5=strFollow5.split("|")
var strFollow6="Follow Up|措施追踪"
var arrFollow6=strFollow6.split("|")

var strConfirm1="Verification|效果确认"
var arrConfirm1=strConfirm1.split("|")
var strConfirm2="Verification|效果确认"
var arrConfirm2=strConfirm2.split("|")
var strConfirm3="Verification|效果确认"
var arrConfirm3=strConfirm3.split("|")
var strConfirm4="The ticket was closed|该单已关闭"
var arrConfirm4=strConfirm4.split("|")
var strConfirm5="Verification|效果确认"
var arrConfirm5=strConfirm5.split("|")

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
var strGreen="Green:Finished;|绿色：完成；" 
var arrGreen=strGreen.split("|")
var strRed="Red:Doing;|红色：进行中；" 
var arrRed=strRed.split("|")
var strBlue="Blue:Waiting;|蓝色：未进行；" 
var arrBlue=strBlue.split("|")
var strRepeat="Is Repeat|是否重复"
var arrRepeat=strRepeat.split("|")
var strRepeatAlert="If it had repeat, please input NCMR NO.|(如若重复，请填写重复的NCMR单号)"
var arrRepeatAlert=strRepeatAlert.split("|")
var strMember="NCMR Member|NCMR 成员"
var arrMember=strMember.split("|")
var strStation="Station|工站" 
var arrStation=strStation.split("|")
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
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_Operator.innerText=arrOperator[<%=session("language")%>]}catch(e){}
try{inner_Alert.innerText=arrAlert[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_CreateTime.innerText=arrCreateTime[<%=session("language")%>]}catch(e){}
try{inner_Creator.innerText=arrCreator[<%=session("language")%>]}catch(e){}
try{inner_Num.innerText=arrNum[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_Model.innerText=arrModel[<%=session("language")%>]}catch(e){}
try{inner_Problem.innerText=arrProblem[<%=session("language")%>]}catch(e){}
try{inner_Upload.innerText=arrUpload[<%=session("language")%>]}catch(e){}
try{inner_Action1.innerText=arrAction1[<%=session("language")%>]}catch(e){}
try{inner_Action2.innerText=arrAction2[<%=session("language")%>]}catch(e){}
try{inner_Action3.innerText=arrAction3[<%=session("language")%>]}catch(e){}
try{inner_Action4.innerText=arrAction4[<%=session("language")%>]}catch(e){}
try{inner_Person1.innerText=arrPerson1[<%=session("language")%>]}catch(e){} 
try{inner_Person2.innerText=arrPerson2[<%=session("language")%>]}catch(e){} 
try{inner_Person3.innerText=arrPerson3[<%=session("language")%>]}catch(e){} 
try{inner_Person4.innerText=arrPerson4[<%=session("language")%>]}catch(e){} 
try{inner_Person5.innerText=arrPerson5[<%=session("language")%>]}catch(e){} 
try{inner_Person6.innerText=arrPerson6[<%=session("language")%>]}catch(e){} 
try{inner_Person7.innerText=arrPerson6[<%=session("language")%>]}catch(e){} 
try{inner_Step1.innerText=arrStep1[<%=session("language")%>]}catch(e){} 
try{inner_Step2.innerText=arrStep2[<%=session("language")%>]}catch(e){} 
try{inner_Step3.innerText=arrStep3[<%=session("language")%>]}catch(e){} 
try{inner_Step4.innerText=arrStep4[<%=session("language")%>]}catch(e){} 
try{inner_Step5.innerText=arrStep5[<%=session("language")%>]}catch(e){} 
try{inner_Time1.innerText=arrTime1[<%=session("language")%>]}catch(e){}
try{inner_Time2.innerText=arrTime2[<%=session("language")%>]}catch(e){}
try{inner_Time3.innerText=arrTime3[<%=session("language")%>]}catch(e){}
try{inner_Time4.innerText=arrTime4[<%=session("language")%>]}catch(e){}
try{inner_Time5.innerText=arrTime5[<%=session("language")%>]}catch(e){}
try{inner_Department.innerText=arrDepartment[<%=session("language")%>]}catch(e){}
try{inner_Department1.innerText=arrDepartment1[<%=session("language")%>]}catch(e){}
try{inner_Defect1.innerText=arrDefect1[<%=session("language")%>]}catch(e){}
try{inner_Defect2.innerText=arrDefect2[<%=session("language")%>]}catch(e){}
try{inner_Defect3.innerText=arrDefect3[<%=session("language")%>]}catch(e){}
try{inner_Defect4.innerText=arrDefect4[<%=session("language")%>]}catch(e){}

try{inner_Records1.innerText=arrRecords1[<%=session("language")%>]}catch(e){}
try{inner_Records2.innerText=arrRecords2[<%=session("language")%>]}catch(e){}
try{inner_Records3.innerText=arrRecords3[<%=session("language")%>]}catch(e){}
try{inner_Records4.innerText=arrRecords4[<%=session("language")%>]}catch(e){}
try{inner_Records5.innerText=arrRecords5[<%=session("language")%>]}catch(e){}
try{inner_Records6.innerText=arrRecords6[<%=session("language")%>]}catch(e){}

try{inner_DueDate.innerText=arrDueDate[<%=session("language")%>]}catch(e){}
try{inner_DueDate1.innerText=arrDueDate1[<%=session("language")%>]}catch(e){}
try{inner_Defect.innerText=arrDefect[<%=session("language")%>]}catch(e){}
try{inner_Defect2.innerText=arrDefect2[<%=session("language")%>]}catch(e){}
try{inner_Defect3.innerText=arrDefect3[<%=session("language")%>]}catch(e){}
try{inner_Defect4.innerText=arrDefect4[<%=session("language")%>]}catch(e){}
try{inner_Risk.innerText=arrRisk[<%=session("language")%>]}catch(e){}
try{inner_Remark.innerText=arrRemark[<%=session("language")%>]}catch(e){} 
try{inner_Remark1.innerText=arrRemark1[<%=session("language")%>]}catch(e){} 
try{inner_Remark2.innerText=arrRemark2[<%=session("language")%>]}catch(e){} 

try{inner_Remark3.innerText=arrRemark3[<%=session("language")%>]}catch(e){} 

try{inner_Cause1.innerText=arrCause1[<%=session("language")%>]}catch(e){} 
try{inner_Cause2.innerText=arrCause2[<%=session("language")%>]}catch(e){} 
try{inner_Cause3.innerText=arrCause3[<%=session("language")%>]}catch(e){} 
try{inner_Cause4.innerText=arrCause4[<%=session("language")%>]}catch(e){} 
try{inner_Cause5.innerText=arrCause5[<%=session("language")%>]}catch(e){} 
try{inner_Attachment.innerText=arrAttachment[<%=session("language")%>]}catch(e){}
try{inner_Attachment1.innerText=arrAttachment1[<%=session("language")%>]}catch(e){}
try{inner_Attachment2.innerText=arrAttachment2[<%=session("language")%>]}catch(e){}

try{inner_Correct1.innerText=arrCorrect1[<%=session("language")%>]}catch(e){} 
try{inner_Correct2.innerText=arrCorrect2[<%=session("language")%>]}catch(e){} 
try{inner_Correct3.innerText=arrCorrect3[<%=session("language")%>]}catch(e){} 
try{inner_Correct4.innerText=arrCorrect4[<%=session("language")%>]}catch(e){} 
try{inner_Correct5.innerText=arrCorrect5[<%=session("language")%>]}catch(e){}
try{inner_Correct6.innerText=arrCorrect6[<%=session("language")%>]}catch(e){}  
try{inner_QA1.innerText=arrQA1[<%=session("language")%>]}catch(e){} 
try{inner_QA2.innerText=arrQA2[<%=session("language")%>]}catch(e){} 
try{inner_QA3.innerText=arrQA3[<%=session("language")%>]}catch(e){} 
try{inner_QA4.innerText=arrQA4[<%=session("language")%>]}catch(e){} 
try{inner_QA5.innerText=arrQA5[<%=session("language")%>]}catch(e){} 
try{inner_QA6.innerText=arrQA6[<%=session("language")%>]}catch(e){} 
try{inner_QA7.innerText=arrQA7[<%=session("language")%>]}catch(e){} 

try{inner_Follow1.innerText=arrFollow1[<%=session("language")%>]}catch(e){} 
try{inner_Follow2.innerText=arrFollow2[<%=session("language")%>]}catch(e){} 
try{inner_Follow3.innerText=arrFollow3[<%=session("language")%>]}catch(e){} 
try{inner_Follow4.innerText=arrFollow4[<%=session("language")%>]}catch(e){} 
try{inner_Follow5.innerText=arrFollow5[<%=session("language")%>]}catch(e){} 
try{inner_Follow6.innerText=arrFollow6[<%=session("language")%>]}catch(e){} 

try{inner_Confirm1.innerText=arrConfirm1[<%=session("language")%>]}catch(e){} 
try{inner_Confirm2.innerText=arrConfirm2[<%=session("language")%>]}catch(e){} 
try{inner_Confirm3.innerText=arrConfirm3[<%=session("language")%>]}catch(e){} 
try{inner_Confirm4.innerText=arrConfirm4[<%=session("language")%>]}catch(e){} 
try{inner_Confirm5.innerText=arrConfirm5[<%=session("language")%>]}catch(e){} 

try{inner_RejectHistory.innerText=arrRejectHistory[<%=session("language")%>]}catch(e){} 
try{inner_RejectStep.innerText=arrRejectStep[<%=session("language")%>]}catch(e){} 
try{inner_RejectCause.innerText=arrRejectCause[<%=session("language")%>]}catch(e){} 
try{inner_RejectTime.innerText=arrRejectTime[<%=session("language")%>]}catch(e){} 
try{inner_RejectPerson.innerText=arrRejectPerson[<%=session("language")%>]}catch(e){} 
try{inner_Repeat.innerText=arrRepeat[<%=session("language")%>]}catch(e){}
try{inner_RepeatAlert.innerText=arrRepeatAlert[<%=session("language")%>]}catch(e){}
try{inner_Member.innerText=arrMember[<%=session("language")%>]}catch(e){} 
try{inner_Owner.innerText=arrOwner[<%=session("language")%>]}catch(e){}
}
</script>