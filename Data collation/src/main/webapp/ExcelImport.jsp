<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Excel匯入</title>
</head>
<body>
<form  method="post" name="form">
	<div class="col-sm-6 text-right">
		<label for="selectFiles" id="importBtnText" class=" btn-theme2 btn">Excel匯入</label>
		<input type="file" class="btn-theme2 btn" id="selectFiles" onchange="javascript:del();" value="Import" name="theFirstFile"/>
	</div>
</form>

<script>
//判斷是否為xls或xlsx檔，並判斷是xls還是xlsx檔
		//這裡控制要檢查的項目，true表示要檢查，false表示不檢查 
		var isCheckType = true;//是否檢查副檔名 
		var xlschekType = true;
		
		//點選提交按鈕觸發下面的函式
		function del(){
			var f = document.form;
			var re = /\.(xls|xlsx)$/i;
			var xe= /\.(xls)$/i;
			if (isCheckType && !re.test(f.theFirstFile.value)) { 
				alert("只允許上傳xls或xlsx檔"); 
			} 
			else if (xlschekType && xe.test(f.theFirstFile.value)) {
				document.form.action="ExcelxlsPr.jsp";
				document.form.enctype="multipart/form-data";
				document.form.submit();	
			}
			else{
				document.form.action="ExcelxlsxPr.jsp";
				document.form.enctype="multipart/form-data";
				document.form.submit();
			}
			
			}
</script>    
</body>
</html>
