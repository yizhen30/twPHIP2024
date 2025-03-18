<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Excel匯入</title>
<script>
function uploadFile() {
    var fileInput = document.getElementById("selectFiles");
    if (!fileInput.files.length) {
        alert("請選擇一個 Excel 檔案！");
        return;
    }

    var file = fileInput.files[0];
    var fileName = file.name;
    var formData = new FormData();
    formData.append("theFirstFile", file);

    // 驗證副檔名
    var re = /\.(xls|xlsx)$/i;
    if (!re.test(fileName)) {
        alert("只允許上傳 xls 或 xlsx 檔案");
        return;
    }

    // 根據副檔名決定要送到哪個 JSP
    var uploadUrl = fileName.endsWith(".xls") ? "ExcelxlsPr.jsp" : "ExcelxlsxPr.jsp";

    // 使用 AJAX 上傳
    fetch(uploadUrl, {
        method: "POST",
        body: formData
    })
    .then(response => response.text())
    .then(result => {
        alert("上傳成功！");
        console.log(result);
    })
    .catch(error => {
        alert("上傳失敗！");
        console.error(error);
    });
}
</script>
</head>
<body>
    <form id="uploadForm" enctype="multipart/form-data">
        <div class="col-sm-6 text-right">
            <label for="selectFiles" class="btn-theme2 btn">Excel匯入</label>
            <input type="file" id="selectFiles" accept=".xls,.xlsx" class="btn-theme2 btn" />
            <button type="button" class="btn-theme2 btn" onclick="uploadFile()">上傳</button>
        </div>
    </form>
</body>
</html>
