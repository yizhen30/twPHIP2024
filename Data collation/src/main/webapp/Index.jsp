<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>鄉鎮人口統計解析器</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: 'Microsoft JhengHei', Arial, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: #fff;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        }
        h1, h2 {
            color: #2c3e50;
        }
        .card {
            background-color: #fff;
            border-radius: 5px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        input[type="file"] {
            display: block;
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            margin-bottom: 10px;
        }
        button {
            background-color: #3498db;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }
        button:hover {
            background-color: #2980b9;
        }
        .result-summary {
            background-color: #edf7ff;
            padding: 10px 15px;
            border-radius: 4px;
            margin-bottom: 15px;
            font-weight: bold;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            font-size: 14px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>資料自動解析器</h1>
        
        <div class="card">
            <h2>選擇檔案</h2>
            <label for="dataFile">選擇資料檔案 (支援 .xls, .xlsx, .csv, .json, .txt)</label>
            <input type="file" id="dataFile" accept=".xls,.xlsx,.csv,.json,.txt">
            <button id="parseBtn">解析檔案</button>
        </div>
        
        <div id="resultContainer" class="card" style="display: none;">
            <h2>解析結果</h2>
            <div id="fileInfo" class="result-summary"></div>
            
            <h3>資料預覽</h3>
            <div style="overflow-x: auto;">
                <table id="resultTable">
                    <thead>
                        <tr>
                            <th>年分</th>
                            <th>月份</th>
                            <th>縣市別</th>
                            <th>鄉鎮別</th>
                            <th>性別</th>
                            <th>人口數</th>
                        </tr>
                    </thead>
                    <tbody id="resultBody"></tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        document.getElementById("parseBtn").addEventListener("click", handleFileSelect);

        function handleFileSelect() {
            const fileInput = document.getElementById("dataFile");
            const file = fileInput.files[0];
            if (!file) {
                alert("請選擇檔案");
                return;
            }

            const fileExt = file.name.split('.').pop().toLowerCase();

            if (fileExt === "xls" || fileExt === "xlsx") {
                parseExcelFile(file);
            } else if (fileExt === "csv") {
                parseCSVFile(file);
            } else if (fileExt === "json") {
                parseJSONFile(file);
            } else if (fileExt === "txt") {
                parseTextFile(file);
            } else {
                alert("不支援的檔案格式");
            }
        }

        function parseExcelFile(file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                processParsedData(jsonData);
            };
            reader.readAsArrayBuffer(file);
        }

        function parseCSVFile(file) {
            const reader = new FileReader();
            reader.onload = function(event) {
                const text = event.target.result;
                const rows = text.split("\n").map(row => row.split(","));
                processParsedData(rows);
            };
            reader.readAsText(file);
        }

        function parseJSONFile(file) {
            const reader = new FileReader();
            reader.onload = function(event) {
                try {
                    const jsonData = JSON.parse(event.target.result);
                    processParsedData(jsonData);
                } catch (error) {
                    alert("JSON 解析錯誤：" + error.message);
                }
            };
            reader.readAsText(file);
        }

        function parseTextFile(file) {
            const reader = new FileReader();
            reader.onload = function(event) {
                const text = event.target.result;
                const rows = text.split("\n").map(row => row.split("\t"));
                processParsedData(rows);
            };
            reader.readAsText(file);
        }

        function processParsedData(data) {
            if (!Array.isArray(data) || data.length === 0) {
                alert("檔案內容錯誤或無法解析");
                return;
            }

            console.log("🔍 原始解析數據:", data); // DEBUG: 印出原始資料

            // 確保標題行是字串陣列
            let headers = data[0];
            if (!Array.isArray(headers) || headers.every(h => h === undefined || h === null || h.toString().trim() === "")) {
                alert("無法偵測到有效的標題行，請確認檔案格式");
                return;
            }
            headers = headers.map(h => (typeof h === "string" ? h.trim() : h?.toString().trim() || "欄位" + Math.random().toString(36).substr(2, 5)));

            console.log("✅ 確認標題行:", headers);

            // 解析資料行
            const parsedData = data.slice(1).filter(row => Array.isArray(row) && row.some(cell => cell !== undefined && cell !== null && cell.toString().trim() !== "")).map(row => {
                let obj = {};
                headers.forEach((header, index) => {
                    let cellValue = row[index];

                    // **避免 undefined、null、false 被當成數據**
                    if (cellValue === undefined || cellValue === null || cellValue === false) {
                        cellValue = "";
                    }

                    // **如果是數字，轉為字串**
                    if (typeof cellValue === "number") {
                        cellValue = cellValue.toString();
                    }

                    // **確保值是字串後才 `.trim()`**
                    obj[header] = (typeof cellValue === "string" ? cellValue.trim() : cellValue);
                });
                return obj;
            });

            console.log("✅ 處理後的數據:", parsedData); // DEBUG: 印出修正後的數據

            displayResults(parsedData);
        }


        function displayResults(data) {
            document.getElementById("fileInfo").innerHTML = `資料總筆數: ${data.length} 筆`;
            document.getElementById("resultBody").innerHTML = "";

            data.forEach((item, index) => {
                if (index < 50) {
                    const row = document.createElement("tr");
                    row.innerHTML = `
                        <td>${item["年分"] || ""}</td>
                        <td>${item["月份"] || ""}</td>
                        <td>${item["縣市別"] || ""}</td>
                        <td>${item["鄉鎮別"] || ""}</td>
                        <td>${item["性別"] || ""}</td>
                        <td>${item["人口數"] || ""}</td>
                    `;
                    document.getElementById("resultBody").appendChild(row);
                }
            });

            document.getElementById("resultContainer").style.display = "block";
        }
    </script>
</body>
</html>
