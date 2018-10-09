<html>
   <head>
      <title>直接打印，选择打印机打印，打印预览</title>
      <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
      <!-- 插入打印控件 -->
<OBJECT ID="jatoolsPrinter" CLASSID="CLSID:B43D3361-D075-4BE2-87FE-057188254255" codebase="jatoolsPrinter.cab#version=8,3,0,0"></OBJECT>  
<script> 
	function doPrint(how) {
		myDoc = {
			documents: document,
			copyrights: '杰创软件拥有版权  www.jatools.com'    
				};
		if(how == '打印预览...')
			jatoolsPrinter.printPreview(myDoc );   // 打印预览
		else if(how == '打印...')
			jatoolsPrinter.print(myDoc ,true);   // 打印前弹出打印设置对话框
		else 
			jatoolsPrinter.print(myDoc ,false);       // 不弹出对话框打印
	}
</script> 
   </head>
   <body style='padding:10px;'>
      <span style='color:#B7C1D3;font-size:25px;font-family:"Microsoft YaHei","黑体",​"Segoe UI",​sans-serif'>直接打印，选择打印机打印，打印预览</span><br><br>
      <input type="button" value="打印预览..." onClick="doPrint('打印预览...')">
      <input type="button" value="打印..." onClick="doPrint('打印...')">
      <input type="button" value="打印" onClick="doPrint('打印')"><br><br>
      <div id='page1' style='width:300px;height:300px;border:2px solid green'>Hello world</div>
   </body>
</html>