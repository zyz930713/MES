<script language="javascript">
var strLine="Line|线别" 
var arrLine=strLine.split("|")
var strTitle="NCMR Detail Information |异常处理单记录单 （不合格处理）"
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
var strIsClosed="Is Closed|此单是否关闭" 
var arrIsClosed=strIsClosed.split("|")
var strCloseAlert="Please input reason if you want to close this NCMR|*如果您认为此NCMR单无须再继续进行下去，请填写原因后，点击该按钮关闭此单" 
var arrCloseAlert=strCloseAlert.split("|")
var strIsSkip="Is skip Long-term Action|是否跳过纠正措施" 
var arrIsSkip=strIsSkip.split("|")
var strIsSkip1="Yes|跳过" 
var arrIsSkip1=strIsSkip1.split("|")
var strIsSkip2="No|不跳过" 
var arrIsSkip2=strIsSkip2.split("|")
var strIsSkip3="Please input skip reason!|（填写跳过原因）" 
var arrIsSkip3=strIsSkip3.split("|")

var strAction1="Short-term Action|立即措施" 
var arrAction1=strAction1.split("|")
var strAction2="Short-term Action|立即措施" 
var arrAction2=strAction2.split("|")
var strAction3="Short-term Action|立即措施" 
var arrAction3=strAction3.split("|")
var strAction4="Short-term Action|立即措施" 
var arrAction4=strAction4.split("|")
var strOperator="Defect Person|本次处理人" 
var arrOperator=strOperator.split("|")
var strAlert="(Please input your Badge Number)|(请输入您的工号)" 
var arrAlert=strAlert.split("|")
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
var strKeyPerson2="Root Cause Person|原因分析相关人" 
var arrKeyPerson2=strKeyPerson2.split("|")
var strKeyPerson3="Long-term Action Person|纠正措施相关人" 
var arrKeyPerson3=strKeyPerson3.split("|")
var strKeyPerson4="QA|QA相关人" 
var arrKeyPerson4=strKeyPerson4.split("|")
var strBackUp2="Backup|备选相关人" 
var arrBackUp2=strBackUp2.split("|")
var strBackUp3="Backup|备选相关人" 
var arrBackUp3=strBackUp3.split("|")
var strBackUp4="Backup|备选相关人" 
var arrBackUp4=strBackUp4.split("|")

var strAttachment="Attachment|附件"
var arrAttachment=strAttachment.split("|")
var strAttachment1="Attachment|附件"
var arrAttachment1=strAttachment1.split("|")
var strAttachment2="Attachment|附件"
var arrAttachment2=strAttachment2.split("|")

var strCorrect1="Long-term Action|纠正措施"
var arrCorrect1=strCorrect1.split("|")

var strCorrect6="Skip Long-term Action|纠正措施被跳过"
var arrCorrect6=strCorrect6.split("|")
var strDueDate2="Due Date|完成日期" 
var arrDueDate2=strDueDate2.split("|")
var strDueDate3="Due Date|完成日期" 
var arrDueDate3=strDueDate3.split("|")
var strDueDate4="Due Date|完成日期" 
var arrDueDate4=strDueDate4.split("|")


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
var strRestart="Redefine|重新定义负责人" 
var arrRestart=strRestart.split("|")
var strRemark="If need Redefine,don't input the follow content,only press the button.|如果重新定义，您无须填写下面内容，仅按“提交”按钮即可	" 
var arrRemark=strRemark.split("|") 

var strIsClosed="Need close|是否关闭N单" 
var arrIsClosed=strIsClosed.split("|") 
var strCloseAlert="If you need close NCMR ticket, please click this button.|如果您想关闭此N单，请单击此按钮。" 
var arrCloseAlert=strCloseAlert.split("|")
var strIsSkip="Need skip|是否跳过" 
var arrIsSkip=strIsSkip.split("|") 
var strIsSkip3="Please key in skip reason!|请填写跳过原因!" 
var arrIsSkip3=strIsSkip3.split("|") 
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
try{inner_Problem.innerText=arrProblem[<%=session("language")%>]}catch(e){}
try{inner_Upload.innerText=arrUpload[<%=session("language")%>]}catch(e){}
try{inner_Action1.innerText=arrAction1[<%=session("language")%>]}catch(e){}
try{inner_Action2.innerText=arrAction2[<%=session("language")%>]}catch(e){}
try{inner_Action3.innerText=arrAction3[<%=session("language")%>]}catch(e){}
try{inner_Action4.innerText=arrAction4[<%=session("language")%>]}catch(e){}
try{inner_Operator.innerText=arrOperator[<%=session("language")%>]}catch(e){}
try{inner_Alert.innerText=arrAlert[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson2.innerText=arrKeyPerson2[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson3.innerText=arrKeyPerson3[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson4.innerText=arrKeyPerson4[<%=session("language")%>]}catch(e){}
try{inner_BackUp2.innerText=arrBackUp2[<%=session("language")%>]}catch(e){}
try{inner_BackUp3.innerText=arrBackUp3[<%=session("language")%>]}catch(e){}
try{inner_BackUp4.innerText=arrBackUp4[<%=session("language")%>]}catch(e){}
try{inner_DueDate2.innerText=arrDueDate2[<%=session("language")%>]}catch(e){}
try{inner_DueDate3.innerText=arrDueDate3[<%=session("language")%>]}catch(e){}
try{inner_DueDate4.innerText=arrDueDate4[<%=session("language")%>]}catch(e){}
try{inner_IsClosed.innerText=arrIsClosed[<%=session("language")%>]}catch(e){}
try{inner_CloseAlert.innerText=arrCloseAlert[<%=session("language")%>]}catch(e){}
try{inner_IsSkip.innerText=arrIsSkip[<%=session("language")%>]}catch(e){}
try{inner_IsSkip1.innerText=arrIsSkip1[<%=session("language")%>]}catch(e){}
try{inner_IsSkip2.innerText=arrIsSkip2[<%=session("language")%>]}catch(e){}
try{inner_IsSkip3.innerText=arrIsSkip3[<%=session("language")%>]}catch(e){}
try{inner_Follow1.innerText=arrFollow1[<%=session("language")%>]}catch(e){} 
try{inner_Follow2.innerText=arrFollow2[<%=session("language")%>]}catch(e){} 
try{inner_Follow3.innerText=arrFollow3[<%=session("language")%>]}catch(e){} 
try{inner_Follow4.innerText=arrFollow4[<%=session("language")%>]}catch(e){} 
try{inner_Follow5.innerText=arrFollow5[<%=session("language")%>]}catch(e){} 
try{inner_Follow6.innerText=arrFollow6[<%=session("language")%>]}catch(e){} 

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

try{inner_Records1.innerText=arrRecords1[<%=session("language")%>]}catch(e){}
try{inner_Records2.innerText=arrRecords2[<%=session("language")%>]}catch(e){}
try{inner_Records3.innerText=arrRecords3[<%=session("language")%>]}catch(e){}
try{inner_Records4.innerText=arrRecords4[<%=session("language")%>]}catch(e){}
try{inner_Records5.innerText=arrRecords5[<%=session("language")%>]}catch(e){}
try{inner_Records6.innerText=arrRecords6[<%=session("language")%>]}catch(e){}

try{inner_DueDate.innerText=arrDueDate[<%=session("language")%>]}catch(e){}
try{inner_DueDate1.innerText=arrDueDate1[<%=session("language")%>]}catch(e){}
try{inner_Defect1.innerText=arrDefect1[<%=session("language")%>]}catch(e){}
try{inner_Defect2.innerText=arrDefect2[<%=session("language")%>]}catch(e){}
try{inner_Defect3.innerText=arrDefect3[<%=session("language")%>]}catch(e){}
try{inner_Defect4.innerText=arrDefect4[<%=session("language")%>]}catch(e){}
try{inner_Defect5.innerText=arrDefect5[<%=session("language")%>]}catch(e){}
try{inner_Defect6.innerText=arrDefect6[<%=session("language")%>]}catch(e){}
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
try{inner_Alert.innerText=arrAlert[<%=session("language")%>]}catch(e){}

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
try{inner_Restart.innerText=arrRestart[<%=session("language")%>]}catch(e){}
try{inner_Remark.innerText=arrRemark[<%=session("language")%>]}catch(e){}

try{inner_IsClosed.innerText=arrIsClosed[<%=session("language")%>]}catch(e){}
try{inner_CloseAlert.innerText=arrCloseAlert[<%=session("language")%>]}catch(e){}
try{inner_IsSkip.innerText=arrIsSkip[<%=session("language")%>]}catch(e){}
try{inner_IsSkip3.innerText=arrIsSkip3[<%=session("language")%>]}catch(e){}
}
</script>