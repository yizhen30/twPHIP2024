<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>資料整理工具</title>
<style>
    #loading { display: none; font-size: 18px; color: blue; }
    #errorMessage { color: red; display: none; }
</style>
<script>
function uploadFile() {
    var fileInput = document.getElementById("selectFiles");
    var processingType = document.getElementById("processingType").value;
    var errorMessage = document.getElementById("errorMessage");
    var loading = document.getElementById("loading");
    errorMessage.style.display = "none";

    if (!fileInput.files.length) {
        errorMessage.innerText = "請選擇一個檔案！";
        errorMessage.style.display = "block";
        return;
    }

    var file = fileInput.files[0];
    var formData = new FormData();
    formData.append("file", file);
    loading.style.display = "block";

    var apiUrl = (processingType === "python") ? 
                 "http://127.0.0.1:5000/ml_process" : "upload.jsp";

    console.log("發送請求到:", apiUrl);  

    fetch(apiUrl, { method: "POST", body: formData })
    .then(response => response.json())
    .then(data => {
        loading.style.display = "none";
        console.log("伺服器回應資料:", data);  
        if (data.error) {
            errorMessage.innerText = "錯誤：" + data.error;
            errorMessage.style.display = "block";
            return;
        }

        var table = document.getElementById("previewTable");
        table.innerHTML = "<tr><th>整理後的資料</th></tr>";

        data.forEach(row => {
            var tr = document.createElement("tr");
            Object.values(row).forEach(cell => {
                var td = document.createElement("td");
                td.textContent = cell;
                tr.appendChild(td);
            });
            table.appendChild(tr);
        });

        document.getElementById("confirmBtn").style.display = "block";
    })
    .catch(error => {
        loading.style.display = "none";
        errorMessage.innerText = "上傳失敗！請檢查伺服器狀態";
        errorMessage.style.display = "block";
        console.error("錯誤發生:", error);
    });
}

function confirmDownload() {
    window.location.href = "process.jsp";
}
</script>
</head>
<body>
    <h2>資料整理工具</h2>
    <form enctype="multipart/form-data">
        <label>選擇處理方式：</label>
        <select id="processingType">
            <option value="jsp">JSP 內建處理</option>
            <option value="python">Python 機器學習</option>
        </select>
        <br><br>
        <input type="file" id="selectFiles" accept=".xls,.xlsx,.csv">
        <button type="button" onclick="uploadFile()">上傳並預覽</button>
    </form>

    <p id="loading">處理中，請稍候...</p>
    <p id="errorMessage"></p>

    <h3>預覽結果</h3>
    <table border="1" id="previewTable"></table>

    <button id="confirmBtn" style="display: none;" onclick="confirmDownload()">確認並下載</button>
</body>
</html>
